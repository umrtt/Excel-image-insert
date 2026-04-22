"""
Validator - Input validation utilities
"""
from pathlib import Path
from typing import List

from app.models import CellPosition


class Validator:
    """
    Validator for input data.
    Follows Single Responsibility Principle - only responsible for validation.
    """

    @staticmethod
    def validate_excel_file(file_path: str) -> bool:
        """
        Validate Excel file.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if valid Excel file, False otherwise
        """
        if not file_path:
            return False
        
        path = Path(file_path)
        if not path.exists():
            return False
        
        # Check if file has Excel extension
        excel_extensions = ['.xlsx', '.xls', '.xlsm']
        return path.suffix.lower() in excel_extensions

    @staticmethod
    def validate_image_file(file_path: str) -> bool:
        """
        Validate image file.
        
        Args:
            file_path: Path to file
            
        Returns:
            True if valid image file, False otherwise
        """
        if not file_path:
            return False
        
        path = Path(file_path)
        if not path.exists():
            return False
        
        # Check if file has image extension
        image_extensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff']
        return path.suffix.lower() in image_extensions

    @staticmethod
    def validate_cell_position(position: str) -> bool:
        """
        Validate cell position format.
        Accepts single cells (A1) or cell ranges (A1:C5).
        
        Args:
            position: Cell position (e.g., 'A1') or range (e.g., 'A1:C5')
            
        Returns:
            True if valid format, False otherwise
        """
        try:
            # Check if it's a range (contains colon)
            if ':' in position:
                parts = position.split(':')
                if len(parts) != 2:
                    return False
                # Validate both parts
                for part in parts:
                    CellPosition.from_string(part.strip())
            else:
                # Single cell
                CellPosition.from_string(position)
            return True
        except ValueError:
            return False

    @staticmethod
    def validate_sheet_name(sheet_name: str, available_sheets: List[str]) -> bool:
        """
        Validate sheet name exists.
        
        Args:
            sheet_name: Sheet name to validate
            available_sheets: List of available sheet names
            
        Returns:
            True if sheet name is valid, False otherwise
        """
        return sheet_name in available_sheets

    @staticmethod
    def validate_image_dimensions(width: int = None, height: int = None) -> bool:
        """
        Validate image dimensions.
        
        Args:
            width: Width in pixels
            height: Height in pixels
            
        Returns:
            True if valid, False otherwise
        """
        if width is not None and width <= 0:
            return False
        if height is not None and height <= 0:
            return False
        return True
