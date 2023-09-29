import sys
import logging
from typing import Optional

class Logger:
    """Logs info, warning and error messages"""
    def __init__(self, name: Optional[str]=None) -> None:
        name = __class__.__name__ if name is None else name

        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.INFO)

        s_handler = logging.StreamHandler()
        f_handler = logging.FileHandler("./logs/logs.log", "w")

        fmt = logging.Formatter("%(name)s:%(levelname)s - %(message)s")

        s_handler.setFormatter(fmt)
        f_handler.setFormatter(fmt)

        s_handler.setLevel(logging.INFO)
        f_handler.setLevel(logging.INFO)

        self.logger.addHandler(s_handler)
        self.logger.addHandler(f_handler)

    def info(self, message: str) -> None:
        """Logs info messages"""
        self.logger.info(message)

    def warn(self, message: str) -> None:
        """Logs warning messages"""
        self.logger.warning(message)

    def error(self, message: str) -> None:
        """Logs error messages"""
        self.logger.error(message, exc_info=True)

        sys.exit(1)