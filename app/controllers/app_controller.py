"""
Application Controller - Handles API endpoints for Eel frontend
"""
from typing import List, Dict, Any

from app.services import ImageService, ExcelService
from app.utils import Logger, FileUtils, RangeUtils
from app.interfaces import IExcelHandler, IFileHandler
from infrastructure.excel_handler import ExcelHandler
from infrastructure.file_handler import FileHandler


class AppController:
    """
    Application controller handling API endpoints.
    Follows Single Responsibility Principle - coordinates between services.
    Uses dependency injection for services.
    """

    def __init__(self):
        """Initialize application controller."""
        self.logger = Logger(__name__)
        
        # Initialize handlers with dependency injection
        self.excel_handler: IExcelHandler = ExcelHandler(self.logger)
        self.file_handler: IFileHandler = FileHandler(self.logger)
        
        # Initialize services with dependency injection
        self.excel_service = ExcelService(self.excel_handler, self.logger)
        self.image_service = ImageService(self.logger)

    # ===================== Excel Operations =====================

    def select_excel_file(self) -> Dict[str, Any]:
        """
        Open file selection dialog for Excel file.
        
        Returns:
            Dict with success status and file path
        """
        try:
            file_path = self.file_handler.select_file()
            
            if not file_path:
                return {
                    'success': False,
                    'message': 'No file selected'
                }
            
            # Try to open the file
            result = self.excel_service.open_file(file_path)
            
            if not result.is_success():
                return {
                    'success': False,
                    'message': result.message,
                    'error': result.error
                }
            
            return {
                'success': True,
                'message': result.message,
                'file_path': file_path,
                'sheets': result.data.get('sheets', [])
            }
            
        except Exception as e:
            self.logger.error(f"Error in select_excel_file: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def get_sheets(self) -> Dict[str, Any]:
        """
        Get list of sheets from current workbook.
        
        Returns:
            Dict with sheets list
        """
        try:
            result = self.excel_service.get_sheets()
            
            return {
                'success': result.is_success(),
                'message': result.message,
                'sheets': result.data.get('sheets', []) if result.data else []
            }
            
        except Exception as e:
            self.logger.error(f"Error in get_sheets: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def close_excel_file(self) -> Dict[str, Any]:
        """
        Close current Excel file.
        
        Returns:
            Dict with success status
        """
        try:
            result = self.excel_service.close_file()
            
            return {
                'success': result.is_success(),
                'message': result.message
            }
            
        except Exception as e:
            self.logger.error(f"Error in close_excel_file: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    # ===================== Image Operations =====================

    def select_images(self) -> Dict[str, Any]:
        """
        Open file selection dialog for images.
        
        Returns:
            Dict with success status and file paths
        """
        try:
            file_paths = self.file_handler.select_files()
            
            if not file_paths:
                return {
                    'success': False,
                    'message': 'No files selected'
                }
            
            return {
                'success': True,
                'message': f'Selected {len(file_paths)} image(s)',
                'file_paths': file_paths
            }
            
        except Exception as e:
            self.logger.error(f"Error in select_images: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def add_image(
        self,
        file_path: str,
        sheet_name: str,
        cell_position: str,
        width: int = None,
        height: int = None
    ) -> Dict[str, Any]:
        """
        Add image to list.
        
        Args:
            file_path: Path to image file
            sheet_name: Target sheet name
            cell_position: Target cell position
            width: Optional width
            height: Optional height
            
        Returns:
            Dict with operation result
        """
        try:
            result = self.image_service.add_image(
                file_path, sheet_name, cell_position, width, height
            )
            
            return {
                'success': result.is_success(),
                'message': result.message,
                'data': result.data
            }
            
        except Exception as e:
            self.logger.error(f"Error in add_image: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def add_images_batch(
        self,
        file_paths: List[str],
        sheet_name: str,
        cell_positions: List[str] = None
    ) -> Dict[str, Any]:
        """
        Add multiple images.
        
        Args:
            file_paths: List of image file paths
            sheet_name: Target sheet name
            cell_positions: Optional list of cell positions
            
        Returns:
            Dict with operation result
        """
        try:
            result = self.image_service.add_images(
                file_paths, sheet_name, cell_positions
            )
            
            return {
                'success': not result.is_error(),
                'message': result.message,
                'images': self.image_service.get_images()
            }
            
        except Exception as e:
            self.logger.error(f"Error in add_images_batch: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def remove_image(self, index: int) -> Dict[str, Any]:
        """
        Remove image from list.
        
        Args:
            index: Image index
            
        Returns:
            Dict with operation result
        """
        try:
            result = self.image_service.remove_image(index)
            
            return {
                'success': result.is_success(),
                'message': result.message,
                'images': self.image_service.get_images()
            }
            
        except Exception as e:
            self.logger.error(f"Error in remove_image: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def clear_images(self) -> Dict[str, Any]:
        """
        Clear all images.
        
        Returns:
            Dict with operation result
        """
        try:
            result = self.image_service.clear_images()
            
            return {
                'success': result.is_success(),
                'message': result.message
            }
            
        except Exception as e:
            self.logger.error(f"Error in clear_images: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def get_images(self) -> Dict[str, Any]:
        """
        Get all images.
        
        Returns:
            Dict with images list
        """
        try:
            images = self.image_service.get_images()
            
            return {
                'success': True,
                'images': images,
                'count': len(images)
            }
            
        except Exception as e:
            self.logger.error(f"Error in get_images: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}',
                'images': []
            }

    # ===================== Insertion Operations =====================

    def insert_images(self) -> Dict[str, Any]:
        """
        Insert all images into Excel file.
        If all images have the same range position, divides the range equally among images.
        Each sub-range is sized to fit its image perfectly.
        
        Returns:
            Dict with operation result
        """
        try:
            if not self.excel_service.is_file_open():
                return {
                    'success': False,
                    'message': 'No Excel file open'
                }
            
            if self.image_service.get_image_count() == 0:
                return {
                    'success': False,
                    'message': 'No images to insert'
                }
            
            # Prepare batch operations
            operations = []
            
            # Check if all images have the same cell_position and it's a range
            images = self.image_service.images
            if len(images) > 1:
                first_position = images[0].cell_position
                first_sheet = images[0].sheet_name
                all_same_position = all(
                    img.cell_position == first_position and img.sheet_name == first_sheet
                    for img in images
                )
                
                # If all have same position and it's a range, divide it among images
                if all_same_position and ':' in first_position:
                    self.logger.info(
                        f"Dividing range {first_position} equally among {len(images)} images"
                    )
                    
                    # Divide the range equally
                    sub_ranges = self.excel_handler.divide_range_equally(
                        first_sheet, first_position, len(images)
                    )
                    
                    self.logger.info(f"Created {len(sub_ranges)} sub-ranges: {sub_ranges}")
                    
                    # Create operations for each image with its sub-range
                    for i, img in enumerate(images):
                        if i < len(sub_ranges):
                            sub_range = sub_ranges[i]
                            operations.append({
                                'sheet_name': img.sheet_name,
                                'cell_position': sub_range,
                                'image_path': img.file_path,
                                'width': None,  # Let Excel size it based on the range
                                'height': None
                            })
                    
                else:
                    # Different positions for each image, use as-is
                    for img in images:
                        operations.append({
                            'sheet_name': img.sheet_name,
                            'cell_position': img.cell_position,
                            'image_path': img.file_path,
                            'width': img.width,
                            'height': img.height
                        })
            else:
                # Single image, use as-is
                for img in images:
                    operations.append({
                        'sheet_name': img.sheet_name,
                        'cell_position': img.cell_position,
                        'image_path': img.file_path,
                        'width': img.width,
                        'height': img.height
                    })
            
            # Insert images
            result = self.excel_service.insert_images_batch(operations)
            
            if result.is_success():
                # Save file
                save_result = self.excel_service.save_file()
                
                if save_result.is_success():
                    # Clear images after successful insertion
                    self.image_service.clear_images()
            
            return {
                'success': result.is_success(),
                'message': result.message,
                'images': self.image_service.get_images()
            }
            
        except Exception as e:
            self.logger.error(f"Error in insert_images: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def insert_single_image(
        self,
        file_path: str,
        sheet_name: str,
        cell_position: str,
        width: int = None,
        height: int = None
    ) -> Dict[str, Any]:
        """
        Insert single image into Excel file.
        
        Args:
            file_path: Path to image file
            sheet_name: Target sheet name
            cell_position: Target cell position
            width: Optional width
            height: Optional height
            
        Returns:
            Dict with operation result
        """
        try:
            if not self.excel_service.is_file_open():
                return {
                    'success': False,
                    'message': 'No Excel file open'
                }
            
            result = self.excel_service.insert_image(
                sheet_name, cell_position, file_path, width, height
            )
            
            if result.is_success():
                self.excel_service.save_file()
            
            return {
                'success': result.is_success(),
                'message': result.message
            }
            
        except Exception as e:
            self.logger.error(f"Error in insert_single_image: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    # ===================== Status Operations =====================

    def get_status(self) -> Dict[str, Any]:
        """
        Get application status.
        
        Returns:
            Dict with status information
        """
        try:
            return {
                'success': True,
                'excel_open': self.excel_service.is_file_open(),
                'current_file': self.excel_service.get_current_file(),
                'images_count': self.image_service.get_image_count(),
                'images': self.image_service.get_images()
            }
            
        except Exception as e:
            self.logger.error(f"Error in get_status: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }

    def select_range_from_excel(self) -> Dict[str, Any]:
        """
        Get the currently selected cell range from Excel.
        User clicks a button to capture the selection.
        
        Returns:
            Dict with selected range
        """
        try:
            if not self.excel_service.is_file_open():
                return {
                    'success': False,
                    'message': 'No Excel file open'
                }
            
            result = self.excel_service.get_selected_range()
            
            return {
                'success': result.is_success(),
                'message': result.message,
                'selected_range': result.data.get('selected_range') if result.data else None
            }
            
        except Exception as e:
            self.logger.error(f"Error in select_range_from_excel: {str(e)}")
            return {
                'success': False,
                'message': f'Error: {str(e)}'
            }
