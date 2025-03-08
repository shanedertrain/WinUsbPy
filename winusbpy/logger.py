import logging
import sys
from pathlib import Path

FMT_LOGGING = '%(asctime)s|%(name)s|%(filename)s|%(levelname)s: %(message)s'

class Logger:
    def __init__(self, name: str, log_file: Path | None = None, level: int = logging.INFO) -> None:
        self.logger = logging.getLogger(name)
        if not self.logger.handlers:  # Prevent duplicate handlers
            self.logger.setLevel(level)

            formatter = logging.Formatter(FMT_LOGGING)

            # Console handler
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)

            # File handler (if provided)
            if log_file:
                file_handler = logging.FileHandler(log_file)
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)

    def get_logger(self) -> logging.Logger:
        return self.logger
