"""
Logger Implementation - Handles logging to file and console
"""
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Optional

from app.interfaces import ILogger, LogLevel


class Logger(ILogger):
    """
    Logger implementation that logs to both console and file.
    Follows Single Responsibility Principle - only responsible for logging.
    """

    def __init__(self, name: str, log_dir: str = "logs"):
        """
        Initialize logger.
        
        Args:
            name: Logger name
            log_dir: Directory for log files
        """
        self.name = name
        self.log_dir = log_dir
        self._ensure_log_dir()
        
        # Create logger
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.DEBUG)
        
        # Remove existing handlers
        self.logger.handlers = []
        
        # Create formatter with timestamp
        self.formatter = logging.Formatter(
            '[%(asctime)s] [%(levelname)s] %(message)s',
            datefmt='%H:%M:%S'
        )
        
        # Add console handler
        self._add_console_handler()
        
        # Add file handler
        self._add_file_handler()

    def _ensure_log_dir(self) -> None:
        """Create log directory if it doesn't exist."""
        Path(self.log_dir).mkdir(parents=True, exist_ok=True)

    def _add_console_handler(self) -> None:
        """Add console logging handler."""
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        console_handler.setFormatter(self.formatter)
        self.logger.addHandler(console_handler)

    def _add_file_handler(self) -> None:
        """Add file logging handler."""
        log_file = self._get_log_filename()
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(self.formatter)
        self.logger.addHandler(file_handler)

    def _get_log_filename(self) -> str:
        """
        Generate log filename based on date.
        Format: debug_ddmm.txt
        
        Returns:
            Full path to log file
        """
        today = datetime.now()
        date_str = today.strftime('%d%m')
        filename = f"debug_{date_str}.txt"
        return os.path.join(self.log_dir, filename)

    def debug(self, message: str) -> None:
        """Log debug message."""
        self.logger.debug(message)

    def info(self, message: str) -> None:
        """Log info message."""
        self.logger.info(message)

    def warning(self, message: str) -> None:
        """Log warning message."""
        self.logger.warning(message)

    def error(self, message: str) -> None:
        """Log error message."""
        self.logger.error(message)

    def critical(self, message: str) -> None:
        """Log critical message."""
        self.logger.critical(message)

    def set_level(self, level: LogLevel) -> None:
        """Set minimum log level."""
        level_map = {
            LogLevel.DEBUG: logging.DEBUG,
            LogLevel.INFO: logging.INFO,
            LogLevel.WARNING: logging.WARNING,
            LogLevel.ERROR: logging.ERROR,
            LogLevel.CRITICAL: logging.CRITICAL,
        }
        self.logger.setLevel(level_map[level])
