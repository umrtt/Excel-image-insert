"""
Excel Handler Implementation using xlwings
"""
from typing import List, Optional, Tuple
from pathlib import Path
import xlwings as xw
from PIL import Image

from app.interfaces import IExcelHandler
from app.utils import Logger


class ExcelHandler(IExcelHandler):
    """
    Excel file handler using xlwings library.
    Implements IExcelHandler interface for dependency injection.
    Follows Single Responsibility Principle - only handles Excel operations.
    """

    def __init__(self, logger: Logger = None):
        """
        Initialize Excel handler.
        
        Args:
            logger: Optional logger instance
        """
        self.logger = logger or Logger(__name__)
        self.workbook = None
        self.app = None

    def open_workbook(self, file_path: str) -> bool:
        """
        Open Excel workbook.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Check if file exists
            if not Path(file_path).exists():
                self.logger.error(f"File not found: {file_path}")
                return False
            
            # Close existing workbook first
            if self.workbook:
                self.workbook.close()
            
            # Use existing Excel app if available, otherwise create new one
            try:
                # Try to get active Excel app
                self.app = xw.apps.active
            except Exception:
                # No active app, create new one
                self.app = xw.App(visible=True)
            
            # Open workbook
            self.workbook = xw.Book(file_path)
            self.logger.info(f"Workbook opened: {file_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to open workbook: {str(e)}")
            return False

    def close_workbook(self) -> bool:
        """
        Close the current workbook.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if self.workbook:
                self.workbook.close()
                self.logger.info("Workbook closed")
                self.workbook = None
            
            if self.app:
                self.app.quit()
                self.app = None
            
            return True
        except Exception as e:
            self.logger.error(f"Failed to close workbook: {str(e)}")
            return False

    def is_workbook_open(self) -> bool:
        """
        Check if workbook is open.
        
        Returns:
            True if workbook is open, False otherwise
        """
        return self.workbook is not None

    def get_sheets(self) -> List[str]:
        """
        Get list of sheet names.
        
        Returns:
            List of sheet names
        """
        try:
            if not self.workbook:
                self.logger.warning("No workbook open")
                return []
            
            sheets = [sheet.name for sheet in self.workbook.sheets]
            return sheets
        except Exception as e:
            self.logger.error(f"Failed to get sheets: {str(e)}")
            return []

    def insert_image(
        self,
        sheet_name: str,
        cell_position: str,
        image_path: str,
        width: Optional[int] = None,
        height: Optional[int] = None
    ) -> bool:
        """
        Insert image into Excel sheet.
        If cell_position is a range (e.g., A1:C5), image is sized to fit the range.
        
        Args:
            sheet_name: Target sheet name
            cell_position: Cell position (e.g., 'A1') or range (e.g., 'A1:C5')
            image_path: Path to image file
            width: Optional width in pixels
            height: Optional height in pixels
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.workbook:
                self.logger.error("No workbook open")
                return False
            
            # Convert image path to absolute path string
            image_path = str(Path(image_path).absolute())
            if not Path(image_path).exists():
                self.logger.error(f"Image file not found: {image_path}")
                return False
            
            # Get sheet
            sheet = self.workbook.sheets[sheet_name]
            
            # Parse cell position - handle both single cells and ranges
            if ':' in cell_position:
                # It's a range (e.g., A1:C5)
                range_parts = cell_position.split(':')
                first_cell = range_parts[0].strip()
                last_cell = range_parts[1].strip()
                cell_range = sheet.range(f"{first_cell}:{last_cell}")
                insert_cell = sheet.range(first_cell)
                width_px = cell_range.width if not width else width
                height_px = cell_range.height if not height else height
                
                self.logger.debug(f"Range selection detected: {cell_position}, size: {width_px}x{height_px}")
            else:
                # Single cell
                cell = sheet.range(cell_position)
                insert_cell = cell
                width_px, height_px = self.get_cell_dimensions(sheet_name, cell_position)
                width_px = width or width_px
                height_px = height or height_px
            
            # Check if cell is part of merged cells
            merged_range = self.get_merged_cell_range(sheet_name, cell_position.split(':')[0].strip())
            
            if merged_range:
                # If merged, get dimensions of entire merged range
                min_row, min_col, max_row, max_col = merged_range
                width_px, height_px = self._get_merged_cell_dimensions(
                    sheet, min_row, min_col, max_row, max_col
                )
            
            # Insert picture
            picture = sheet.pictures.add(
                image_path,
                left=insert_cell.left,
                top=insert_cell.top,
                width=width_px,
                height=height_px
            )
            
            self.logger.info(
                f"Image inserted at {sheet_name}!{cell_position}: {Path(image_path).name}"
            )
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to insert image: {str(e)}")
            return False

    def delete_images(self, sheet_name: str, cell_position: str) -> bool:
        """
        Delete images at specific cell position or range.
        
        Args:
            sheet_name: Sheet name
            cell_position: Cell position or range (e.g., 'A1' or 'A1:C5')
            
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.workbook:
                self.logger.error("No workbook open")
                return False
            
            sheet = self.workbook.sheets[sheet_name]
            
            # Handle range or single cell
            if ':' in cell_position:
                # It's a range
                cell_range = sheet.range(cell_position)
            else:
                # Single cell
                cell_range = sheet.range(cell_position)
            
            # Find pictures that overlap with this cell/range
            pictures_to_delete = []
            for pic in sheet.pictures:
                # Check if picture overlaps with cell/range
                if (pic.left < cell_range.left + cell_range.width and
                    pic.left + pic.width > cell_range.left and
                    pic.top < cell_range.top + cell_range.height and
                    pic.top + pic.height > cell_range.top):
                    pictures_to_delete.append(pic)
            
            # Delete pictures
            for pic in pictures_to_delete:
                pic.delete()
            
            if pictures_to_delete:
                self.logger.info(
                    f"Deleted {len(pictures_to_delete)} image(s) at {sheet_name}!{cell_position}"
                )
            
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to delete images: {str(e)}")
            return False

    def get_merged_cell_range(
        self,
        sheet_name: str,
        cell_position: str
    ) -> Optional[Tuple[int, int, int, int]]:
        """
        Get merged cell range if cell is part of merged cells.
        
        Args:
            sheet_name: Sheet name
            cell_position: Cell position
            
        Returns:
            Tuple of (min_row, min_col, max_row, max_col) or None
        """
        try:
            if not self.workbook:
                return None
            
            sheet = self.workbook.sheets[sheet_name]
            cell = sheet.range(cell_position)
            
            # Check if cell is merged using xlwings API
            try:
                # Access merged cells through the API
                merged_areas = sheet.api.MergedAreas
                for merged_range in merged_areas:
                    # Check if our cell is in this merged range
                    if (cell.row >= merged_range.Row and
                        cell.row <= merged_range.Row + merged_range.Rows.Count - 1 and
                        cell.column >= merged_range.Column and
                        cell.column <= merged_range.Column + merged_range.Columns.Count - 1):
                        
                        min_row = merged_range.Row
                        min_col = merged_range.Column
                        max_row = merged_range.Row + merged_range.Rows.Count - 1
                        max_col = merged_range.Column + merged_range.Columns.Count - 1
                        
                        return (min_row, min_col, max_row, max_col)
            except Exception:
                # No merged cells or API error, continue
                pass
            
            return None
            
        except Exception as e:
            self.logger.debug(f"Failed to get merged cell range: {str(e)}")
            return None

    def get_cell_dimensions(self, sheet_name: str, cell_position: str) -> Tuple[float, float]:
        """
        Get cell dimensions in pixels.
        
        Args:
            sheet_name: Sheet name
            cell_position: Cell position
            
        Returns:
            Tuple of (width, height) in pixels
        """
        try:
            if not self.workbook:
                return (200, 200)  # Default size
            
            sheet = self.workbook.sheets[sheet_name]
            cell = sheet.range(cell_position)
            
            width = cell.width
            height = cell.height
            
            # Ensure minimum dimensions
            width = max(width, 50) if width else 200
            height = max(height, 50) if height else 200
            
            return (width, height)
            
        except Exception as e:
            self.logger.error(f"Failed to get cell dimensions: {str(e)}")
            return (200, 200)

    def divide_range_equally(self, sheet_name: str, cell_range: str, num_divisions: int) -> List[str]:
        """
        Divide a cell range equally into N parts.
        Used for distributing multiple images across a selected range.
        
        Examples:
            "A1:L1" divided by 4 -> ["A1:D1", "E1:H1", "I1:J1", "K1:L1"]
            "A1:L4" divided by 2 -> ["A1:D4", "E1:L4"]
        
        Args:
            sheet_name: Sheet name
            cell_range: Cell range to divide (e.g., "A1:L1")
            num_divisions: Number of parts to divide into
            
        Returns:
            List of sub-ranges
        """
        try:
            from app.utils import RangeUtils
            
            if num_divisions <= 0:
                return [cell_range]
            
            if ':' not in cell_range:
                return [cell_range]
            
            # Parse the range
            parts = cell_range.split(':')
            start_cell = parts[0].strip()
            end_cell = parts[1].strip()
            
            start_coords = RangeUtils._parse_cell(start_cell)
            end_coords = RangeUtils._parse_cell(end_cell)
            
            if not start_coords or not end_coords:
                return [cell_range]
            
            start_col, start_row = start_coords
            end_col, end_row = end_coords
            
            # Ensure correct order
            min_col = min(start_col, end_col)
            max_col = max(start_col, end_col)
            min_row = min(start_row, end_row)
            max_row = max(start_row, end_row)
            
            # Calculate division width (in columns)
            total_cols = max_col - min_col + 1
            cols_per_division = max(1, total_cols // num_divisions)
            
            result_ranges = []
            
            for i in range(num_divisions):
                # Calculate column range for this division
                div_start_col = min_col + (i * cols_per_division)
                
                # Last division gets any remaining columns
                if i == num_divisions - 1:
                    div_end_col = max_col
                else:
                    div_end_col = div_start_col + cols_per_division - 1
                
                # Create range for this division
                start_cell_div = RangeUtils._col_to_letters(div_start_col) + str(min_row)
                end_cell_div = RangeUtils._col_to_letters(div_end_col) + str(max_row)
                
                div_range = f"{start_cell_div}:{end_cell_div}"
                result_ranges.append(div_range)
                
                self.logger.debug(f"Division {i+1}/{num_divisions}: {div_range}")
            
            return result_ranges
            
        except Exception as e:
            self.logger.error(f"Failed to divide range: {str(e)}")
            return [cell_range]

    def _get_merged_cell_dimensions(
        self,
        sheet,
        min_row: int,
        min_col: int,
        max_row: int,
        max_col: int
    ) -> Tuple[float, float]:
        """
        Calculate dimensions of merged cell range.
        
        Args:
            sheet: Excel sheet object
            min_row: Minimum row number
            min_col: Minimum column number
            max_row: Maximum row number
            max_col: Maximum column number
            
        Returns:
            Tuple of (width, height)
        """
        try:
            merged_range = sheet.range((min_row, min_col), (max_row, max_col))
            width = merged_range.width
            height = merged_range.height
            
            return (max(width, 50), max(height, 50))
        except Exception as e:
            self.logger.error(f"Failed to get merged cell dimensions: {str(e)}")
            return (200, 200)

    def save_workbook(self) -> bool:
        """
        Save the current workbook.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.workbook:
                self.logger.error("No workbook open")
                return False
            
            self.workbook.save()
            self.logger.info("Workbook saved")
            return True
        except Exception as e:
            self.logger.error(f"Failed to save workbook: {str(e)}")
            return False

    def get_selected_range(self) -> Optional[str]:
        """
        Get the currently selected range in Excel.
        
        Returns:
            Cell range as string (e.g., 'A1' or 'A1:C5') or None
        """
        try:
            # Get the active Excel application
            try:
                app = xw.apps.active
            except Exception:
                self.logger.debug("No active Excel application found")
                return None
            
            if not app:
                self.logger.debug("No active Excel app for range selection")
                return None
            
            # Get the currently selected range from Excel
            selected = app.selection
            
            if selected:
                # Get the address of the selection (returns absolute reference like $A$1:$L$1)
                address = selected.address
                # Strip dollar signs to convert to relative reference format (A1:L1)
                clean_address = address.replace('$', '')
                self.logger.debug(f"Selected range from Excel: {clean_address}")
                return clean_address
            
            return None
            
        except Exception as e:
            self.logger.error(f"Failed to get selected range: {str(e)}")
            return None
