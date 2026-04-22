"""
Operation Result Model - Represents the result of an operation
"""
from dataclasses import dataclass
from enum import Enum
from typing import Any, Optional


class ResultStatus(Enum):
    """Status of an operation."""
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"
    INFO = "info"


@dataclass
class OperationResult:
    """
    Represents the result of an operation.
    
    Attributes:
        status: Operation status (success, warning, error, info)
        message: Human-readable message
        data: Optional data payload
        error: Optional error details
    """
    status: ResultStatus
    message: str
    data: Optional[Any] = None
    error: Optional[str] = None

    @classmethod
    def success(cls, message: str, data: Any = None) -> 'OperationResult':
        """Create a successful result."""
        return cls(ResultStatus.SUCCESS, message, data)

    @classmethod
    def error(cls, message: str, error: str = None, data: Any = None) -> 'OperationResult':
        """Create an error result."""
        return cls(ResultStatus.ERROR, message, data, error)

    @classmethod
    def warning(cls, message: str, data: Any = None) -> 'OperationResult':
        """Create a warning result."""
        return cls(ResultStatus.WARNING, message, data)

    @classmethod
    def info(cls, message: str, data: Any = None) -> 'OperationResult':
        """Create an info result."""
        return cls(ResultStatus.INFO, message, data)

    def is_success(self) -> bool:
        """Check if operation was successful."""
        return self.status == ResultStatus.SUCCESS

    def is_error(self) -> bool:
        """Check if operation resulted in error."""
        return self.status == ResultStatus.ERROR

    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            'status': self.status.value,
            'message': self.message,
            'data': self.data,
            'error': self.error
        }
