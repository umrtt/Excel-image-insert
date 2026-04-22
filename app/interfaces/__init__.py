"""
Interfaces Package - Abstract base classes for dependency inversion
"""
from .excel_handler_interface import IExcelHandler
from .file_handler_interface import IFileHandler
from .logger_interface import ILogger, LogLevel

__all__ = ['IExcelHandler', 'IFileHandler', 'ILogger', 'LogLevel']
