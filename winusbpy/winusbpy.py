import struct
import ctypes
from enum import Enum
from dataclasses import dataclass
from ctypes import c_byte, byref, sizeof, c_ulong, resize, wstring_at, c_void_p, c_ubyte, create_string_buffer, WinError
from ctypes.wintypes import DWORD, HANDLE
from typing import Optional
from pathlib import Path
from .winusb import WinUSBApi
from .winusbclasses import GUID, DIGCF_ALLCLASSES, DIGCF_DEFAULT, DIGCF_PRESENT, DIGCF_PROFILE, DIGCF_DEVICE_INTERFACE, \
    SpDeviceInterfaceData,  SpDeviceInterfaceDetailData, SpDevinfoData, GENERIC_WRITE, GENERIC_READ, FILE_SHARE_WRITE, \
    FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, FILE_FLAG_OVERLAPPED, INVALID_HANDLE_VALUE, \
    UsbInterfaceDescriptor, PipeInfo, ERROR_IO_INCOMPLETE, ERROR_IO_PENDING, Overlapped
from .winusbutils import SetupDiGetClassDevs, SetupDiEnumDeviceInterfaces, SetupDiGetDeviceInterfaceDetail, is_device, \
    CreateFile, WinUsb_Initialize, Close_Handle, WinUsb_Free, GetLastError, WinUsb_QueryDeviceInformation, \
    WinUsb_GetAssociatedInterface, WinUsb_QueryInterfaceSettings, WinUsb_QueryPipe, WinUsb_ControlTransfer, \
    WinUsb_WritePipe, WinUsb_ReadPipe, WinUsb_GetOverlappedResult, SetupDiGetDeviceRegistryProperty, \
    WinUsb_SetPipePolicy, WinUsb_FlushPipe, SPDRP_FRIENDLYNAME
from .logger import Logger, logging

def is_64bit():
    return struct.calcsize('P') * 8 == 64

@dataclass
class UsbDevice:
    name:str
    path:str
    api:WinUSBApi
    handle_file:HANDLE = INVALID_HANDLE_VALUE
    handle_winusb:c_void_p = c_void_p()
    interface_index:int = -1
    logging_level:int = logging.INFO
    logging_filepath:Optional[Path] = None

    def __post_init__(self):
        self.logger = Logger(self.name, level=self.logging_level, log_file=self.logging_filepath).get_logger()
        self.logger.debug(f"UsbDevice initialized: {self}")

    def init_device(self) -> bool:
        # Open the device
        self.handle_file = self.api.exec_function_kernel32(CreateFile, self.path, 
                                                            GENERIC_WRITE | GENERIC_READ,
                                                            FILE_SHARE_WRITE | FILE_SHARE_READ, None, 
                                                            OPEN_EXISTING,
                                                            FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED, None)

        # Check if the device handle is valid
        if self.handle_file == INVALID_HANDLE_VALUE:
            self.logger.error(f"Failed to open device at {self.path}. Invalid handle.")
            return False  # Failed to open device

        self.logger.debug(f"Device opened successfully at {self.path}")

        # Initialize WinUSB handle
        result = self.api.exec_function_winusb(WinUsb_Initialize, self.handle_file, byref(self.handle_winusb))
        
        if result == 0:
            err = self.get_last_error_code()
            raise ctypes.WinError()

        self.logger.debug("WinUSB handle initialized successfully.")

        return True
    
    def get_last_error_code(self):
        return self.api.exec_function_kernel32(GetLastError)

    def close_device(self):
        result_file = self.api.exec_function_kernel32(Close_Handle, self.handle_file)
        result_winusb = self.api.exec_function_winusb(WinUsb_Free, self.handle_winusb)
        return result_file != 0 and result_winusb != 0

    def query_device_info(self, query=1):
        info_type = c_ulong(query)
        buff = (c_void_p * 1)()
        buff_length = c_ulong(sizeof(c_void_p))
        result = self.api.exec_function_winusb(WinUsb_QueryDeviceInformation, self.handle_winusb, info_type,
                                        byref(buff_length), buff)
        if result != 0:
            return buff[0]
        else:
            return -1

    def query_interface_settings(self, index):
        if self.interface_index != -1:
            temp_handle_winusb = self.handle_winusb
            if self.interface_index != 0:
                result = self.api.exec_function_winusb(WinUsb_GetAssociatedInterface, self.handle_winusb,
                                                    c_ubyte(index), byref(temp_handle_winusb))
                if result == 0:
                    return False
            interface_descriptor = UsbInterfaceDescriptor()
            result = self.api.exec_function_winusb(WinUsb_QueryInterfaceSettings, temp_handle_winusb, c_ubyte(0),
                                                byref(interface_descriptor))
            if result != 0:
                return interface_descriptor
            else:
                return None
        else:
            return None

    def change_interface(self, index):
        result = self.api.exec_function_winusb(WinUsb_GetAssociatedInterface, self.handle_winusb, c_ubyte(index),
                                            byref(self.handle_winusb))
        if result != 0:
            self.interface_index = index
            return True
        else:
            return False

    def query_pipe(self, pipe_index):
        pipe_info = PipeInfo()
        result = self.api.exec_function_winusb(WinUsb_QueryPipe, self.handle_winusb, c_ubyte(0), pipe_index,
                                            byref(pipe_info))
        if result != 0:
            return pipe_info
        else:
            return None

    def control_transfer(self, setup_packet:bytes, buff=None):
        if buff != None:
            if setup_packet.length > 0:  # Host 2 Device
                buff = (c_ubyte * setup_packet.length)(*buff)
                buffer_length = setup_packet.length
            else:  # Device 2 Host
                buff = (c_ubyte * setup_packet.length)()
                buffer_length = setup_packet.length
        else:
            buff = c_ubyte()
            buffer_length = 0

        result = self.api.exec_function_winusb(WinUsb_ControlTransfer, self.handle_winusb, setup_packet, byref(buff),
                                            c_ulong(buffer_length), byref(c_ulong(0)), None)
        return {"result": result != 0, "buffer": [buff]}

    def write(self, pipe_id:int, write_buffer:bytearray):
        buffer_type = c_ubyte * len(write_buffer)
        write_buffer = buffer_type(*write_buffer)
        written = c_ulong(0)

        if not self.handle_winusb:
            raise ValueError("WinUSB handle is not initialized")

        result = self.api.exec_function_winusb(
            WinUsb_WritePipe, self.handle_winusb, c_ubyte(pipe_id), 
            write_buffer, c_ulong(len(write_buffer)), byref(written), None
        )

        if result == 0:
            raise ctypes.WinError(self.get_last_error_code())

        return written.value

    def read(self, pipe_id: int, length_buffer: int) -> bytearray:
        read_buffer = (c_ubyte * length_buffer)()
        read = c_ulong(0)

        # Call WinUsb_ReadPipe
        result = self.api.exec_function_winusb(
            WinUsb_ReadPipe, self.handle_winusb, c_ubyte(pipe_id),
            read_buffer, c_ulong(length_buffer), byref(read), None
        )
        
        if result != 0:
            if read.value != length_buffer:
                return bytearray(read_buffer[:read.value])  # Return only the valid bytes
            else:
                return bytearray(read_buffer)  # Return the entire buffer
        else:
            error_code = self.get_last_error_code()
            if error_code != 0:
                raise ctypes.WinError(error_code)
            else:
                raise RuntimeError("WinUsb_ReadPipe failed but no error code was returned.")

    def set_timeout(self, pipe_id:int, timeout_ms:int) -> int:
        class POLICY_TYPE:
            SHORT_PACKET_TERMINATE = 1
            AUTO_CLEAR_STALL = 2
            PIPE_TRANSFER_TIMEOUT = 3
            IGNORE_SHORT_PACKETS = 4
            ALLOW_PARTIAL_READS = 5
            AUTO_FLUSH = 6
            RAW_IO = 7

        policy_type = c_ulong(POLICY_TYPE.PIPE_TRANSFER_TIMEOUT)
        value_length = c_ulong(4)
        value = c_ulong(timeout_ms)
        result = self.api.exec_function_winusb(WinUsb_SetPipePolicy, self.handle_winusb, c_ubyte(pipe_id),
                                            policy_type, value_length, byref(value))
        return result

    def flush(self, pipe_id:int):
        result = self.api.exec_function_winusb(WinUsb_FlushPipe, self.handle_winusb, c_ubyte(pipe_id))
        return result

    def _overlapped_read_do(self, pipe_id:int) -> bool:
        self.olread_ol.Internal = 0
        self.olread_ol.InternalHigh = 0
        self.olread_ol.Offset = 0
        self.olread_ol.OffsetHigh = 0
        self.olread_ol.Pointer = 0
        self.olread_ol.hEvent = 0                
        result = self.api.exec_function_winusb(WinUsb_ReadPipe, self.handle_winusb, c_ubyte(pipe_id), self.olread_buf, 
                                            c_ulong(self.olread_buflen), byref(c_ulong(0)), byref(self.olread_ol))
        if result != 0:
            return True
        else:
            return False
                
    def overlapped_read_init(self, pipe_id, length_buffer) -> bool:
        self.olread_ol = Overlapped()
        self.olread_buf = create_string_buffer(length_buffer)
        self.olread_buflen = length_buffer
        return self._overlapped_read_do(pipe_id)

    def overlapped_read(self, pipe_id:int) -> str:
        """ keep on reading overlapped, return bytearray, empty if nothing to read, None if err"""
        rl = c_ulong(0)
        result = self.api.exec_function_winusb(WinUsb_GetOverlappedResult, self.handle_winusb, byref(self.olread_ol),byref(rl),True)
        if result == 0:
            last_error = self.get_last_error_code()
            if last_error == ERROR_IO_PENDING or \
                last_error == ERROR_IO_INCOMPLETE:
                return str(last_error)
            else:
                return None
        else:
            ret = self.olread_buf[0:rl.value]
            # self._overlapped_read_do(pipe_id)
            return ret

def is_64bit():
    return struct.calcsize('P') * 8 == 64

byte_array = c_byte * 8
class WinUsbPy(object):
    class GUIDEnum(Enum):
        USB_DEVICE = GUID(0xA5DCBF10, 0x6530, 0x11D2, byte_array(0x90, 0x1F, 0x00, 0xC0, 0x4F, 0xB9, 0x51, 0xED))
        USB_WINUSB = GUID(0xdee824ef, 0x729b, 0x4a0e, byte_array(0x9c, 0x14, 0xb7, 0x11, 0x7d, 0x33, 0xa8, 0x17))
        USB_COMPOSITE = GUID(0x36FC9E60, 0xC465, 0x11CF, byte_array(0x80, 0x56, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00))
        
    def __init__(self, logging_level=logging.INFO, log_filepath: Optional[Path]=None):
        self.logging_level = logging_level
        self.log_filepath = log_filepath
        self.api = WinUSBApi()

    def get_usb_devices(self, guid:GUIDEnum, **kwargs) -> list[UsbDevice]:
        """Retrieve a dictionary of connected USB devices with their paths."""
        flags = self._compute_flags(**kwargs)
        handle = self.api.exec_function_setupapi(
            SetupDiGetClassDevs, byref(guid.value), None, None, flags
        )

        return self._enumerate_usb_devices(handle, guid)
    
    def get_usb_devices_filtered(self, guid:GUIDEnum, vid: str, pid: str, **kwargs) -> list[UsbDevice]:
        devices:list[UsbDevice] = self.get_usb_devices(guid, **kwargs)
        return [dev for dev in devices if is_device(vid, pid, dev.path)]

    def _compute_flags(self, **kwargs) -> DWORD:
        """Compute flag values based on provided keyword arguments."""
        flag_map = {
            "default": DIGCF_DEFAULT,
            "present": DIGCF_PRESENT,
            "allclasses": DIGCF_ALLCLASSES,
            "profile": DIGCF_PROFILE,
            "deviceinterface": DIGCF_DEVICE_INTERFACE,
        }
        value = sum(flag for key, flag in flag_map.items() if kwargs.get(key)) or 0x00000010
        return DWORD(value)

    def _enumerate_usb_devices(self, handle, guid:GUIDEnum) -> list[UsbDevice]:
        """Enumerate USB devices and return a dictionary of device names and paths."""
        devices:list[UsbDevice] = []

        # Type hinting for these variables
        sp_device_interface_data: SpDeviceInterfaceData = SpDeviceInterfaceData()
        sp_device_interface_data.cb_size = sizeof(sp_device_interface_data)
        
        sp_device_interface_detail_data: SpDeviceInterfaceDetailData = SpDeviceInterfaceDetailData()
        
        sp_device_info_data: SpDevinfoData = SpDevinfoData()
        sp_device_info_data.cb_size = sizeof(sp_device_info_data)

        i = 0
        required_size = DWORD(0)
        member_index = DWORD(i)
        cb_sizes = (8, 6, 5)  # Different sizes for different architectures

        while self.api.exec_function_setupapi(
            SetupDiEnumDeviceInterfaces, handle, None, byref(guid.value),
            member_index, byref(sp_device_interface_data)
        ):
            self.api.exec_function_setupapi(
                SetupDiGetDeviceInterfaceDetail, handle,
                byref(sp_device_interface_data), None, 0, byref(required_size), None
            )
            resize(sp_device_interface_detail_data, required_size.value)

            path: str = None  # Initialize path variable with type hinting
            for cb_size in cb_sizes:
                sp_device_interface_detail_data.cb_size = cb_size
                ret = self.api.exec_function_setupapi(
                    SetupDiGetDeviceInterfaceDetail, handle,
                    byref(sp_device_interface_data), byref(sp_device_interface_detail_data),
                    required_size, byref(required_size), byref(sp_device_info_data)
                )
                if ret:
                    cb_sizes = (cb_size,)
                    path = wstring_at(byref(sp_device_interface_detail_data, sizeof(DWORD)))
                    break
            if path is None:
                raise ctypes.WinError()

            name = self._get_device_friendly_name(handle, sp_device_info_data, path)
            devices.append(UsbDevice(name, path, self.api, logging_level=self.logging_level, 
                                        logging_filepath=self.log_filepath))

            i += 1
            member_index = DWORD(i)
            required_size = c_ulong(0)
            resize(sp_device_interface_detail_data, sizeof(SpDeviceInterfaceDetailData))

        return devices

    def _get_device_friendly_name(self, handle, sp_device_info_data, default_name: str) -> str:
        """Retrieve the friendly name of the USB device if available."""
        buff_friendly_name = ctypes.create_unicode_buffer(250)
        if self.api.exec_function_setupapi(
            SetupDiGetDeviceRegistryProperty, handle,
            byref(sp_device_info_data), SPDRP_FRIENDLYNAME,
            None, ctypes.byref(buff_friendly_name),
            ctypes.sizeof(buff_friendly_name) - 1, None
        ):
            return buff_friendly_name.value
        return default_name
    
    def get_last_error_code(self):
        return self.api.exec_function_kernel32(GetLastError)