"""
Image Model - Represents image data and metadata
"""
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ImageModel:
    """
    Represents an image file with metadata.
    
    Attributes:
        file_path: Path to the image file
        sheet_name: Target Excel sheet name
        cell_position: Target cell position (e.g., 'A1')
        width: Optional width in pixels (None = auto-fit)
        height: Optional height in pixels (None = auto-fit)
    """
    file_path: str
    sheet_name: str
    cell_position: str
    width: int = None
    height: int = None

    def __post_init__(self):
        """Validate image file exists."""
        if not Path(self.file_path).exists():
            raise FileNotFoundError(f"Image file not found: {self.file_path}")

    @property
    def filename(self) -> str:
        """Get the filename without path."""
        return Path(self.file_path).name

    @property
    def file_size_mb(self) -> float:
        """Get file size in megabytes."""
        return Path(self.file_path).stat().st_size / (1024 * 1024)

    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            'file_path': self.file_path,
            'sheet_name': self.sheet_name,
            'cell_position': self.cell_position,
            'width': self.width,
            'height': self.height,
            'filename': self.filename,
            'file_size_mb': round(self.file_size_mb, 2)
        }
