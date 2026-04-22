"""
Logger Interface - Abstract interface for logging
"""
from abc import ABC, abstractmethod
from enum import Enum


class LogLevel(Enum):
    """Log level enumeration."""
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"


class ILogger(ABC):
    """
    Interface for logging operations.
    Implements Dependency Inversion Principle - abstraction for logging.
    """

    @abstractmethod
    def debug(self, message: str) -> None:
        """Log debug message."""
        pass

    @abstractmethod
    def info(self, message: str) -> None:
        """Log info message."""
        pass

    @abstractmethod
    def warning(self, message: str) -> None:
        """Log warning message."""
        pass

    @abstractmethod
    def error(self, message: str) -> None:
        """Log error message."""
        pass

    @abstractmethod
    def critical(self, message: str) -> None:
        """Log critical message."""
        pass

    @abstractmethod
    def set_level(self, level: LogLevel) -> None:
        """Set minimum log level."""
        pass
