"""
Image Service - Business logic for image operations
"""
from typing import List
from pathlib import Path

from app.models import ImageModel, OperationResult
from app.utils import Logger, Validator, FileUtils


class ImageService:
    """
    Image service for image operations.
    Follows Single Responsibility Principle - handles image validation and processing.
    """

    def __init__(self, logger: Logger = None):
        """
        Initialize image service.
        
        Args:
            logger: Optional logger instance
        """
        self.logger = logger or Logger(__name__)
        self.images: List[ImageModel] = []

    def add_image(
        self,
        file_path: str,
        sheet_name: str,
        cell_position: str,
        width: int = None,
        height: int = None
    ) -> OperationResult:
        """
        Add image to the list.
        
        Args:
            file_path: Path to image file
            sheet_name: Target sheet name
            cell_position: Target cell position
            width: Optional width in pixels
            height: Optional height in pixels
            
        Returns:
            OperationResult indicating success or failure
        """
        try:
            # Validate image file
            if not Validator.validate_image_file(file_path):
                msg = f"Invalid image file: {file_path}"
                self.logger.warning(msg)
                return OperationResult.error(msg)
            
            # Validate cell position
            if not Validator.validate_cell_position(cell_position):
                msg = f"Invalid cell position: {cell_position}"
                self.logger.warning(msg)
                return OperationResult.error(msg)
            
            # Validate dimensions
            if not Validator.validate_image_dimensions(width, height):
                msg = "Invalid image dimensions"
                self.logger.warning(msg)
                return OperationResult.error(msg)
            
            # Create image model
            image = ImageModel(
                file_path=file_path,
                sheet_name=sheet_name,
                cell_position=cell_position,
                width=width,
                height=height
            )
            
            self.images.append(image)
            
            msg = f"Image added: {FileUtils.get_filename(file_path)}"
            self.logger.info(msg)
            
            return OperationResult.success(
                msg,
                data=image.to_dict()
            )
            
        except Exception as e:
            msg = f"Failed to add image: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def add_images(
        self,
        file_paths: List[str],
        sheet_name: str,
        cell_positions: List[str] = None
    ) -> OperationResult:
        """
        Add multiple images.
        
        Args:
            file_paths: List of image file paths
            sheet_name: Target sheet name
            cell_positions: Optional list of cell positions (must match file count)
            
        Returns:
            OperationResult with summary
        """
        try:
            if not file_paths:
                return OperationResult.warning("No images provided")
            
            # If cell positions provided, must match file count
            if cell_positions and len(cell_positions) != len(file_paths):
                msg = "Number of cell positions must match number of images"
                return OperationResult.error(msg)
            
            added = 0
            failed = 0
            
            for i, file_path in enumerate(file_paths):
                # Use provided cell position or request it from UI
                cell_pos = cell_positions[i] if cell_positions else None
                
                if cell_pos:
                    result = self.add_image(file_path, sheet_name, cell_pos)
                    if result.is_success():
                        added += 1
                    else:
                        failed += 1
            
            msg = f"Added {added} image(s), failed {failed}"
            if failed > 0:
                return OperationResult.warning(msg)
            else:
                return OperationResult.success(msg)
                
        except Exception as e:
            msg = f"Failed to add images: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def remove_image(self, index: int) -> OperationResult:
        """
        Remove image from list.
        
        Args:
            index: Image index
            
        Returns:
            OperationResult
        """
        try:
            if index < 0 or index >= len(self.images):
                msg = "Invalid image index"
                return OperationResult.error(msg)
            
            removed = self.images.pop(index)
            msg = f"Image removed: {removed.filename}"
            self.logger.info(msg)
            
            return OperationResult.success(msg)
            
        except Exception as e:
            msg = f"Failed to remove image: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def clear_images(self) -> OperationResult:
        """
        Clear all images.
        
        Returns:
            OperationResult
        """
        try:
            count = len(self.images)
            self.images.clear()
            
            msg = f"Cleared {count} image(s)"
            self.logger.info(msg)
            
            return OperationResult.success(msg)
            
        except Exception as e:
            msg = f"Failed to clear images: {str(e)}"
            self.logger.error(msg)
            return OperationResult.error(msg, error=str(e))

    def get_images(self) -> List[dict]:
        """
        Get all images as dictionaries.
        
        Returns:
            List of image dictionaries
        """
        return [img.to_dict() for img in self.images]

    def get_image_count(self) -> int:
        """
        Get number of images.
        
        Returns:
            Image count
        """
        return len(self.images)
