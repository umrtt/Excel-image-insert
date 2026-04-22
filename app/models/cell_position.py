"""
Cell Position Model - Represents Excel cell coordinates
"""
import re
from dataclasses import dataclass
from typing import Tuple


@dataclass
class CellPosition:
    """
    Represents an Excel cell position using column and row.
    Converts between letter notation (A1) and numeric coordinates.
    
    Attributes:
        column: Column number (1-based)
        row: Row number (1-based)
    """
    column: int
    row: int

    @classmethod
    def from_string(cls, position_str: str) -> 'CellPosition':
        """
        Parse cell position from string format (e.g., 'A1', 'Z99').
        
        Args:
            position_str: Cell position as string (e.g., 'A1')
            
        Returns:
            CellPosition instance
            
        Raises:
            ValueError: If position string format is invalid
        """
        # Remove whitespace and convert to uppercase
        position_str = position_str.strip().upper()
        
        # Match pattern: letters followed by numbers
        match = re.match(r'^([A-Z]+)(\d+)$', position_str)
        if not match:
            raise ValueError(f"Invalid cell position format: {position_str}")
        
        col_letters, row_str = match.groups()
        column = cls._letters_to_column(col_letters)
        row = int(row_str)
        
        if row < 1:
            raise ValueError(f"Row must be >= 1, got {row}")
        
        return cls(column, row)

    @staticmethod
    def _letters_to_column(letters: str) -> int:
        """
        Convert column letters to column number.
        A=1, B=2, ..., Z=26, AA=27, etc.
        
        Args:
            letters: Column letters (e.g., 'A', 'AA')
            
        Returns:
            Column number (1-based)
        """
        result = 0
        for char in letters:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    @staticmethod
    def _column_to_letters(column: int) -> str:
        """
        Convert column number to letters.
        1=A, 2=B, ..., 26=Z, 27=AA, etc.
        
        Args:
            column: Column number (1-based)
            
        Returns:
            Column letters
        """
        result = ""
        while column > 0:
            column -= 1
            result = chr(column % 26 + ord('A')) + result
            column //= 26
        return result

    def to_string(self) -> str:
        """
        Convert to cell position string format.
        
        Returns:
            Cell position as string (e.g., 'A1')
        """
        col_letters = self._column_to_letters(self.column)
        return f"{col_letters}{self.row}"

    def __str__(self) -> str:
        """String representation."""
        return self.to_string()

    def __repr__(self) -> str:
        """Developer representation."""
        return f"CellPosition({self.column}, {self.row})"
