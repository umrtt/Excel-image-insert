"""
File Handler Interface - Abstract interface for file operations
"""
from abc import ABC, abstractmethod
from typing import List


class IFileHandler(ABC):
    """
    Interface for file system operations.
    Implements Dependency Inversion Principle - abstraction for file operations.
    """

    @abstractmethod
    def select_file(self, file_types: str = "All files (*.*)|*.*") -> str:
        """
        Open file selection dialog.
        
        Args:
            file_types: File type filter
            
        Returns:
            Selected file path or empty string if cancelled
        """
        pass

    @abstractmethod
    def select_files(self, file_types: str = "All files (*.*)|*.*") -> List[str]:
        """
        Open file selection dialog for multiple files.
        
        Args:
            file_types: File type filter
            
        Returns:
            List of selected file paths
        """
        pass

    @abstractmethod
    def select_directory(self) -> str:
        """
        Open directory selection dialog.
        
        Returns:
            Selected directory path or empty string if cancelled
        """
        pass

    @abstractmethod
    def file_exists(self, file_path: str) -> bool:
        """
        Check if file exists.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if file exists, False otherwise
        """
        pass

    @abstractmethod
    def get_file_extension(self, file_path: str) -> str:
        """
        Get file extension.
        
        Args:
            file_path: Path to file
            
        Returns:
            File extension (e.g., '.xlsx')
        """
        pass

    @abstractmethod
    def get_filename(self, file_path: str) -> str:
        """
        Get filename from path.
        
        Args:
            file_path: Path to file
            
        Returns:
            Filename with extension
        """
        pass
