"""
File Utilities - File system helper functions
"""
from pathlib import Path
from typing import List


class FileUtils:
    """
    File system utilities.
    Follows Single Responsibility Principle - file operations helper.
    """

    @staticmethod
    def file_exists(file_path: str) -> bool:
        """
        Check if file exists.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if file exists, False otherwise
        """
        return Path(file_path).exists()

    @staticmethod
    def get_extension(file_path: str) -> str:
        """
        Get file extension.
        
        Args:
            file_path: Path to file
            
        Returns:
            File extension (e.g., '.xlsx')
        """
        return Path(file_path).suffix.lower()

    @staticmethod
    def get_filename(file_path: str) -> str:
        """
        Get filename from path.
        
        Args:
            file_path: Path to file
            
        Returns:
            Filename with extension
        """
        return Path(file_path).name

    @staticmethod
    def get_absolute_path(file_path: str) -> str:
        """
        Get absolute path.
        
        Args:
            file_path: Path to file
            
        Returns:
            Absolute path
        """
        return str(Path(file_path).resolve())

    @staticmethod
    def is_valid_image_extension(file_path: str) -> bool:
        """
        Check if file has valid image extension.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if valid image extension, False otherwise
        """
        image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.ico'}
        ext = Path(file_path).suffix.lower()
        return ext in image_extensions

    @staticmethod
    def is_valid_excel_extension(file_path: str) -> bool:
        """
        Check if file has valid Excel extension.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if valid Excel extension, False otherwise
        """
        excel_extensions = {'.xlsx', '.xls', '.xlsm'}
        ext = Path(file_path).suffix.lower()
        return ext in excel_extensions
