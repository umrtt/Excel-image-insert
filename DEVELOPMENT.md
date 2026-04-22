# Development Guide and Examples

## Example Usage Scenarios

### Scenario 1: Simple Image Insertion

```python
from app.controllers import AppController

# Initialize controller
controller = AppController()

# Select Excel file
result = controller.select_excel_file()
# User selects: C:\files\report.xlsx

# Get available sheets
sheets_result = controller.get_sheets()
# Returns: ['Sheet1', 'Sheet2', 'Report']

# Add image to specific cell
image_result = controller.add_image(
    file_path='C:\\images\\logo.png',
    sheet_name='Sheet1',
    cell_position='A1'
)

# Insert all images
insert_result = controller.insert_images()
# File is saved automatically
```

### Scenario 2: Batch Image Insertion

```python
# Add multiple images
controller.add_image('img1.png', 'Sheet1', 'A1')
controller.add_image('img2.png', 'Sheet1', 'B1')
controller.add_image('img3.png', 'Sheet2', 'A1')

# Insert all at once
result = controller.insert_images()
# Batch processing handles all insertions
```

### Scenario 3: With Custom Dimensions

```python
# Add image with specific size
controller.add_image(
    file_path='photo.jpg',
    sheet_name='Sheet1',
    cell_position='A1',
    width=200,  # pixels
    height=150   # pixels
)

# Insert
controller.insert_images()
```

### Scenario 4: Merged Cell Handling

```python
# Add image to merged cell range
# If cells A1:C3 are merged, this automatically fits to that range
controller.add_image(
    file_path='chart.png',
    sheet_name='Sheet1',
    cell_position='A1'  # Part of merged range A1:C3
)

# Image will be sized to fit entire merged range
controller.insert_images()
```

## Architecture Diagram

```
User Interface (Web Frontend)
    ↓
Eel (Desktop GUI Wrapper)
    ↓
AppController (API Endpoints)
    ├─→ ImageService (Image Logic)
    │   ├─→ Validator (Input Validation)
    │   ├─→ Logger (Logging)
    │   └─→ ImageModel (Data Model)
    │
    └─→ ExcelService (Excel Logic)
        ├─→ IExcelHandler (Interface)
        │   └─→ ExcelHandler (xlwings Implementation)
        ├─→ Validator
        └─→ Logger

[Database/File System]
    ↓
Excel Files & Image Files
```

## Testing Examples

### Test Image Service

```python
import pytest
from app.services import ImageService
from app.utils import Logger

def test_add_valid_image():
    """Test adding a valid image"""
    service = ImageService(Logger(__name__))
    
    result = service.add_image(
        'test.png',
        'Sheet1',
        'A1'
    )
    
    assert result.is_success()
    assert service.get_image_count() == 1

def test_invalid_image_file():
    """Test adding invalid image file"""
    service = ImageService(Logger(__name__))
    
    result = service.add_image(
        'nonexistent.png',
        'Sheet1',
        'A1'
    )
    
    assert result.is_error()
    assert service.get_image_count() == 0

def test_invalid_cell_position():
    """Test invalid cell position"""
    service = ImageService(Logger(__name__))
    
    result = service.add_image(
        'test.png',
        'Sheet1',
        'INVALID'
    )
    
    assert result.is_error()
```

### Test Cell Position Model

```python
from app.models import CellPosition

def test_parse_cell_position():
    """Test parsing cell position"""
    cell = CellPosition.from_string('A1')
    assert cell.column == 1
    assert cell.row == 1
    assert str(cell) == 'A1'

def test_large_cell_position():
    """Test large column numbers"""
    cell = CellPosition.from_string('AA100')
    assert cell.column == 27
    assert cell.row == 100

def test_invalid_position():
    """Test invalid position"""
    with pytest.raises(ValueError):
        CellPosition.from_string('INVALID')
```

### Test Excel Service

```python
from app.services import ExcelService
from infrastructure.excel_handler import ExcelHandler

def test_open_excel_file():
    """Test opening Excel file"""
    handler = ExcelHandler()
    service = ExcelService(handler)
    
    result = service.open_file('test.xlsx')
    
    assert result.is_success()
    assert service.is_file_open()
    
    # Cleanup
    service.close_file()

def test_get_sheets():
    """Test getting sheet names"""
    handler = ExcelHandler()
    service = ExcelService(handler)
    
    service.open_file('test.xlsx')
    result = service.get_sheets()
    
    assert result.is_success()
    assert 'Sheet1' in result.data['sheets']
    
    service.close_file()
```

## Error Handling Patterns

### Pattern 1: Result-Based Error Handling

```python
def process_image(file_path):
    result = image_service.add_image(file_path, 'Sheet1', 'A1')
    
    if result.is_success():
        logger.info(f"Success: {result.message}")
    else:
        logger.error(f"Error: {result.message}")
        if result.error:
            logger.error(f"Details: {result.error}")
```

### Pattern 2: Exception-Based Error Handling

```python
try:
    excel_service.open_file(file_path)
except FileNotFoundError:
    logger.error(f"File not found: {file_path}")
except Exception as e:
    logger.error(f"Unexpected error: {str(e)}")
```

### Pattern 3: Validation-Based Error Handling

```python
if not Validator.validate_excel_file(file_path):
    logger.warning(f"Invalid Excel file: {file_path}")
    return OperationResult.error("Invalid file")

if not Validator.validate_cell_position(cell_pos):
    logger.warning(f"Invalid cell position: {cell_pos}")
    return OperationResult.error("Invalid position")
```

## Extending the Application

### Adding a New Image Format Support

1. **Update Validator:**
```python
# In app/utils/validator.py
@staticmethod
def validate_image_file(file_path: str) -> bool:
    # ... existing code ...
    image_extensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp']  # Added .webp
    return path.suffix.lower() in image_extensions
```

2. **Update File Handler:**
```python
# In infrastructure/file_handler.py
filetypes=[
    ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff *.webp"),  # Added .webp
    ("All files", "*.*")
]
```

3. **Test the new format:**
```python
def test_webp_image_support():
    result = image_service.add_image('test.webp', 'Sheet1', 'A1')
    assert result.is_success()
```

### Adding a New Excel Handler (e.g., openpyxl)

1. **Create new implementation:**
```python
# infrastructure/openpyxl_handler.py
from app.interfaces import IExcelHandler
from openpyxl import load_workbook

class OpenpyxlHandler(IExcelHandler):
    def __init__(self, logger=None):
        self.logger = logger or Logger(__name__)
        self.workbook = None
    
    def open_workbook(self, file_path: str) -> bool:
        try:
            self.workbook = load_workbook(file_path)
            return True
        except Exception as e:
            self.logger.error(f"Failed: {e}")
            return False
    
    # Implement other methods...
```

2. **Use the new handler:**
```python
# In app_controller.py
from infrastructure.openpyxl_handler import OpenpyxlHandler

self.excel_handler: IExcelHandler = OpenpyxlHandler(self.logger)
```

### Adding Image Resizing Feature

```python
# app/services/image_service.py (extended)
from PIL import Image as PILImage

def resize_image(self, image_path: str, width: int, height: int) -> str:
    """Resize image to specified dimensions"""
    img = PILImage.open(image_path)
    img = img.resize((width, height), PILImage.Resampling.LANCZOS)
    
    # Save resized image
    resized_path = image_path.replace('.', '_resized.')
    img.save(resized_path)
    
    return resized_path
```

## Performance Optimization

### Batch Processing Optimization

```python
# Optimized batch insertion
def insert_images_batch(self, operations: List[dict]) -> OperationResult:
    """
    Insert multiple images efficiently.
    Minimizes Excel file writes.
    """
    # Open file once
    success_count = 0
    failed_count = 0
    
    for op in operations:
        # Process each image
        if self.insert_image(**op):
            success_count += 1
        else:
            failed_count += 1
    
    # Save file once at the end
    self.excel_handler.save_workbook()
    
    return OperationResult.success(f"Inserted {success_count}")
```

### Memory-Efficient Logging

```python
# app/utils/logger.py - Add log rotation
import logging.handlers

def _add_file_handler(self) -> None:
    log_file = self._get_log_filename()
    
    # Use RotatingFileHandler for memory efficiency
    handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,  # 5 MB
        backupCount=5,  # Keep 5 backup files
        encoding='utf-8'
    )
    
    handler.setFormatter(self.formatter)
    self.logger.addHandler(handler)
```

## Debugging Tips

### Enable Debug Logging

```python
from app.utils import Logger
from app.interfaces import LogLevel

logger = Logger(__name__)
logger.set_level(LogLevel.DEBUG)
```

### Add Tracing to Services

```python
def add_image(self, file_path: str, ...):
    self.logger.debug(f"add_image called with: {file_path}")
    
    # Validation
    if not Validator.validate_image_file(file_path):
        self.logger.debug(f"Validation failed for: {file_path}")
        return OperationResult.error(...)
    
    self.logger.debug("Validation passed, creating ImageModel")
    # ... rest of code
```

### Monitor Excel Operations

```python
def insert_image(self, sheet_name: str, ...):
    self.logger.debug(f"Attempting to insert image at {sheet_name}!{cell_position}")
    
    sheet = self.workbook.sheets[sheet_name]
    self.logger.debug(f"Sheet reference obtained: {sheet}")
    
    # Check merged cells
    merged = self.get_merged_cell_range(sheet_name, cell_position)
    self.logger.debug(f"Merged cell check result: {merged}")
    
    # ... rest of code
```

---

This development guide provides patterns and examples for using and extending the Excel Image Insert application.
