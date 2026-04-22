"""
Range Utilities - Helper functions for Excel range parsing
"""
from typing import List
import re


class RangeUtils:
    """
    Utility class for Excel range operations.
    """

    @staticmethod
    def expand_range(cell_range: str) -> List[str]:
        """
        Expand a cell range into individual cells.
        
        Examples:
            "A1:D1" -> ["A1", "B1", "C1", "D1"]
            "A1:A5" -> ["A1", "A2", "A3", "A4", "A5"]
            "A1" -> ["A1"]
            "A1:C3" -> ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"]
        
        Args:
            cell_range: Cell range string (e.g., "A1:D5" or single cell "A1")
            
        Returns:
            List of individual cell positions
        """
        # Handle single cell
        if ':' not in cell_range:
            return [cell_range.strip()]
        
        # Parse range
        parts = cell_range.split(':')
        if len(parts) != 2:
            return [cell_range.strip()]
        
        start_cell = parts[0].strip()
        end_cell = parts[1].strip()
        
        # Parse cell coordinates
        start_coords = RangeUtils._parse_cell(start_cell)
        end_coords = RangeUtils._parse_cell(end_cell)
        
        if not start_coords or not end_coords:
            return [cell_range.strip()]
        
        start_col, start_row = start_coords
        end_col, end_row = end_coords
        
        # Ensure correct order
        min_col = min(start_col, end_col)
        max_col = max(start_col, end_col)
        min_row = min(start_row, end_row)
        max_row = max(start_row, end_row)
        
        # Generate all cells in range
        cells = []
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cell = RangeUtils._col_to_letters(col) + str(row)
                cells.append(cell)
        
        return cells

    @staticmethod
    def _parse_cell(cell: str) -> tuple:
        """
        Parse cell reference to (column_number, row_number).
        
        Args:
            cell: Cell reference (e.g., "A1", "Z99", "AA1")
            
        Returns:
            Tuple of (column_number, row_number) or None
        """
        match = re.match(r'^([A-Z]+)(\d+)$', cell.strip().upper())
        if not match:
            return None
        
        col_letters, row_str = match.groups()
        col_num = RangeUtils._letters_to_col(col_letters)
        row_num = int(row_str)
        
        return (col_num, row_num)

    @staticmethod
    def _letters_to_col(letters: str) -> int:
        """
        Convert column letters to column number.
        A=1, B=2, ..., Z=26, AA=27, AB=28, ..., AZ=52, BA=53, etc.
        
        Args:
            letters: Column letters (e.g., 'A', 'AA', 'ABC')
            
        Returns:
            Column number (1-based)
        """
        result = 0
        for char in letters:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    @staticmethod
    def _col_to_letters(col_num: int) -> str:
        """
        Convert column number to letters.
        1=A, 2=B, ..., 26=Z, 27=AA, 28=AB, etc.
        
        Args:
            col_num: Column number (1-based)
            
        Returns:
            Column letters
        """
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
