"""
Excel Service - Business logic for Excel operations
"""
from typing import List, Optional

from app.interfaces import IExcelHandler
from app.models import OperationResult
from app.utils import Logger, Validator


class ExcelService:
    """
    Excel service for Excel operations.
    Follows Single Responsibility Principle and Dependency Inversion Principle.
    Depends on IExcelHandler abstraction, not concrete implementation.
    """

    def __init__(self, excel_handler: IExcelHandler, logger: Logger = None):
        """
        Initialize Excel service.
        
        Args:
            excel_handler: Implementation of IExcelHandler interface
            logger: Optional logger instance
        """
        self.excel_handler = excel_handler
        self.logger = logger or Logger(__name__)
        self.current_file = None

    def open_file(self, file_path: str) -> OperationResult:
        """
        Open Excel file.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            OperationResult
        """
        try:
            # Validate Excel file
            if not Validator.validate_excel_file(file_path):
                msg = f"Invalid Excel file: {file_path}"
                self.logger.warning(msg)
                return OperationResult.error(msg)
            
            # Close existing workbook if any
            if self.current_file:
                self.excel_handler.close_workbook()
            
            # Open new workbook
            if not self.excel_handler.open_workbook(file_path):
                msg = "Failed to open workbook"
                return OperationResult.error(msg)
            
            self.current_file = file_path
            
            msg = f"Excel file opened: {file_path}"
            self.logger.info(msg)
            
            # Get sheet names
            sheets = self.excel_handler.get_sheets()
            
            return OperationResult.success(
                msg,
                data={'sheets': sheets}
            )
            
        except Exception as e:
            msg = f"Failed to open Excel file: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def close_file(self) -> OperationResult:
        """
        Close current Excel file.
        
        Returns:
            OperationResult
        """
        try:
            if not self.excel_handler.is_workbook_open():
                return OperationResult.warning("No workbook open")
            
            if not self.excel_handler.close_workbook():
                msg = "Failed to close workbook"
                return OperationResult.error(msg)
            
            self.current_file = None
            msg = "Excel file closed"
            self.logger.info(msg)
            
            return OperationResult.success(msg)
            
        except Exception as e:
            msg = f"Failed to close Excel file: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def get_sheets(self) -> OperationResult:
        """
        Get list of sheet names.
        
        Returns:
            OperationResult with sheet names
        """
        try:
            if not self.excel_handler.is_workbook_open():
                msg = "No workbook open"
                return OperationResult.error(msg)
            
            sheets = self.excel_handler.get_sheets()
            
            if not sheets:
                msg = "No sheets found"
                return OperationResult.warning(msg)
            
            msg = f"Retrieved {len(sheets)} sheet(s)"
            self.logger.info(msg)
            
            return OperationResult.success(msg, data={'sheets': sheets})
            
        except Exception as e:
            msg = f"Failed to get sheets: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def insert_image(
        self,
        sheet_name: str,
        cell_position: str,
        image_path: str,
        width: Optional[int] = None,
        height: Optional[int] = None
    ) -> OperationResult:
        """
        Insert image into Excel sheet.
        
        Args:
            sheet_name: Target sheet name
            cell_position: Cell position
            image_path: Path to image file
            width: Optional width
            height: Optional height
            
        Returns:
            OperationResult
        """
        try:
            if not self.excel_handler.is_workbook_open():
                msg = "No workbook open"
                return OperationResult.error(msg)
            
            # Validate sheet name
            sheets = self.excel_handler.get_sheets()
            if not Validator.validate_sheet_name(sheet_name, sheets):
                msg = f"Invalid sheet name: {sheet_name}"
                self.logger.warning(msg)
                return OperationResult.error(msg)
            
            # Validate cell position
            if not Validator.validate_cell_position(cell_position):
                msg = f"Invalid cell position: {cell_position}"
                self.logger.warning(msg)
                return OperationResult.error(msg)
            
            # Delete existing images at this position
            self.excel_handler.delete_images(sheet_name, cell_position)
            
            # Insert image
            if not self.excel_handler.insert_image(
                sheet_name, cell_position, image_path, width, height
            ):
                msg = f"Failed to insert image at {sheet_name}!{cell_position}"
                return OperationResult.error(msg)
            
            msg = f"Image inserted at {sheet_name}!{cell_position}"
            self.logger.info(msg)
            
            return OperationResult.success(msg)
            
        except Exception as e:
            msg = f"Failed to insert image: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def insert_images_batch(
        self,
        operations: List[dict]
    ) -> OperationResult:
        """
        Insert multiple images in batch.
        
        Args:
            operations: List of operation dicts with keys:
                        sheet_name, cell_position, image_path, width, height
            
        Returns:
            OperationResult with summary
        """
        try:
            if not self.excel_handler.is_workbook_open():
                msg = "No workbook open"
                return OperationResult.error(msg)
            
            if not operations:
                return OperationResult.warning("No operations provided")
            
            success_count = 0
            failed_count = 0
            
            for op in operations:
                result = self.insert_image(
                    sheet_name=op.get('sheet_name'),
                    cell_position=op.get('cell_position'),
                    image_path=op.get('image_path'),
                    width=op.get('width'),
                    height=op.get('height')
                )
                
                if result.is_success():
                    success_count += 1
                else:
                    failed_count += 1
            
            msg = f"Inserted {success_count} image(s), failed {failed_count}"
            
            if failed_count > 0:
                return OperationResult.warning(msg)
            else:
                return OperationResult.success(msg)
                
        except Exception as e:
            msg = f"Failed to insert images: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def save_file(self) -> OperationResult:
        """
        Save current Excel file.
        
        Returns:
            OperationResult
        """
        try:
            if not self.excel_handler.is_workbook_open():
                msg = "No workbook open"
                return OperationResult.error(msg)
            
            if not self.excel_handler.save_workbook():
                msg = "Failed to save workbook"
                return OperationResult.error(msg)
            
            msg = "Excel file saved"
            self.logger.info(msg)
            
            return OperationResult.success(msg)
            
        except Exception as e:
            msg = f"Failed to save Excel file: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def is_file_open(self) -> bool:
        """
        Check if Excel file is open.
        
        Returns:
            True if file is open, False otherwise
        """
        return self.excel_handler.is_workbook_open()

    def get_current_file(self) -> Optional[str]:
        """
        Get path to current open file.
        
        Returns:
            File path or None
        """
        return self.current_file

    def get_selected_range(self) -> OperationResult:
        """
        Get the currently selected range from Excel.
        
        Returns:
            OperationResult with selected range
        """
        try:
            if not self.excel_handler.is_workbook_open():
                msg = "No workbook open"
                return OperationResult.error(msg)
            
            selected_range = self.excel_handler.get_selected_range()
            
            if not selected_range:
                msg = "No cells selected in Excel"
                return OperationResult.warning(msg)
            
            msg = f"Selected range: {selected_range}"
            self.logger.info(msg)
            
            return OperationResult.success(
                msg,
                data={'selected_range': selected_range}
            )
            
        except Exception as e:
            msg = f"Failed to get selected range: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))
