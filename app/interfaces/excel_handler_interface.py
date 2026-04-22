"""
Excel Handler Interface - Abstract interface for Excel operations
"""
from abc import ABC, abstractmethod
from typing import List, Tuple, Optional


class IExcelHandler(ABC):
    """
    Interface for Excel file operations.
    Implements Dependency Inversion Principle - high-level modules depend on abstractions.
    """

    @abstractmethod
    def open_workbook(self, file_path: str) -> bool:
        """
        Open an Excel workbook.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            True if successful, False otherwise
        """
        pass

    @abstractmethod
    def close_workbook(self) -> bool:
        """
        Close the current workbook.
        
        Returns:
            True if successful, False otherwise
        """
        pass

    @abstractmethod
    def get_sheets(self) -> List[str]:
        """
        Get list of sheet names.
        
        Returns:
            List of sheet names
        """
        pass

    @abstractmethod
    def insert_image(
        self,
        sheet_name: str,
        cell_position: str,
        image_path: str,
        width: Optional[int] = None,
        height: Optional[int] = None
    ) -> bool:
        """
        Insert image into Excel sheet at specified position.
        
        Args:
            sheet_name: Target sheet name
            cell_position: Cell position (e.g., 'A1')
            image_path: Path to image file
            width: Optional width in pixels
            height: Optional height in pixels
            
        Returns:
            True if successful, False otherwise
        """
        pass

    @abstractmethod
    def delete_images(self, sheet_name: str, cell_position: str) -> bool:
        """
        Delete images at specific cell position.
        
        Args:
            sheet_name: Sheet name
            cell_position: Cell position
            
        Returns:
            True if successful, False otherwise
        """
        pass

    @abstractmethod
    def get_merged_cell_range(self, sheet_name: str, cell_position: str) -> Optional[Tuple[int, int, int, int]]:
        """
        Get merged cell range coordinates if cell is part of merged cells.
        
        Args:
            sheet_name: Sheet name
            cell_position: Cell position
            
        Returns:
            Tuple of (min_row, min_col, max_row, max_col) or None if not merged
        """
        pass

    @abstractmethod
    def get_cell_dimensions(self, sheet_name: str, cell_position: str) -> Tuple[float, float]:
        """
        Get cell dimensions in pixels.
        
        Args:
            sheet_name: Sheet name
            cell_position: Cell position
            
        Returns:
            Tuple of (width, height) in pixels
        """
        pass

    @abstractmethod
    def save_workbook(self) -> bool:
        """
        Save the current workbook.
        
        Returns:
            True if successful, False otherwise
        """
        pass

    @abstractmethod
    def is_workbook_open(self) -> bool:
        """
        Check if workbook is currently open.
        
        Returns:
            True if workbook is open, False otherwise
        """
        pass

    @abstractmethod
    def get_selected_range(self) -> Optional[str]:
        """
        Get the currently selected cell range in Excel.
        
        Returns:
            Cell range as string (e.g., 'A1', 'A1:C5') or None if nothing selected
        """
        pass
