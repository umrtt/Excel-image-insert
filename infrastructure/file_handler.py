"""
File Handler Implementation
"""
from typing import List
from pathlib import Path
from tkinter import filedialog, Tk
import tkinter as tk

from app.interfaces import IFileHandler
from app.utils import Logger


class FileHandler(IFileHandler):
    """
    File handler using tkinter for dialogs.
    Implements IFileHandler interface for dependency injection.
    Follows Single Responsibility Principle - file dialog operations.
    """

    def __init__(self, logger: Logger = None):
        """
        Initialize file handler.
        
        Args:
            logger: Optional logger instance
        """
        self.logger = logger or Logger(__name__)

    def select_file(self, file_types: str = "All files (*.*)|*.*") -> str:
        """
        Open file selection dialog.
        
        Args:
            file_types: File type filter (tkinter format)
            
        Returns:
            Selected file path or empty string if cancelled
        """
        try:
            # Create hidden root window
            root = Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            
            file_path = filedialog.askopenfilename(
                title="Select File",
                filetypes=[("All files", "*.*")]
            )
            
            root.destroy()
            
            if file_path:
                self.logger.info(f"File selected: {file_path}")
            else:
                self.logger.debug("File selection cancelled")
            
            return file_path
            
        except Exception as e:
            self.logger.error(f"Failed to select file: {str(e)}")
            return ""

    def select_files(self, file_types: str = "All files (*.*)|*.*") -> List[str]:
        """
        Open file selection dialog for multiple files.
        
        Args:
            file_types: File type filter
            
        Returns:
            List of selected file paths
        """
        try:
            root = Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            
            file_paths = filedialog.askopenfilenames(
                title="Select Images",
                filetypes=[
                    ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                    ("All files", "*.*")
                ]
            )
            
            root.destroy()
            
            if file_paths:
                self.logger.info(f"Selected {len(file_paths)} file(s)")
            else:
                self.logger.debug("File selection cancelled")
            
            return list(file_paths)
            
        except Exception as e:
            self.logger.error(f"Failed to select files: {str(e)}")
            return []

    def select_directory(self) -> str:
        """
        Open directory selection dialog.
        
        Returns:
            Selected directory path or empty string if cancelled
        """
        try:
            root = Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            
            directory = filedialog.askdirectory(title="Select Directory")
            
            root.destroy()
            
            if directory:
                self.logger.info(f"Directory selected: {directory}")
            else:
                self.logger.debug("Directory selection cancelled")
            
            return directory
            
        except Exception as e:
            self.logger.error(f"Failed to select directory: {str(e)}")
            return ""

    def file_exists(self, file_path: str) -> bool:
        """
        Check if file exists.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if file exists, False otherwise
        """
        return Path(file_path).exists()

    def get_file_extension(self, file_path: str) -> str:
        """
        Get file extension.
        
        Args:
            file_path: Path to file
            
        Returns:
            File extension (e.g., '.xlsx')
        """
        return Path(file_path).suffix.lower()

    def get_filename(self, file_path: str) -> str:
        """
        Get filename from path.
        
        Args:
            file_path: Path to file
            
        Returns:
            Filename with extension
        """
        return Path(file_path).name
