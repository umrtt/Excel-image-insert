"""
Configuration Module - Application settings
"""

# Application Settings
APP_NAME = "Excel Image Insert"
APP_VERSION = "1.0.0"
APP_DESCRIPTION = "Desktop application for inserting images into Excel files"

# UI Settings
UI_WINDOW_WIDTH = 1000
UI_WINDOW_HEIGHT = 900
UI_WINDOW_POSITION_X = 100
UI_WINDOW_POSITION_Y = 100
UI_PORT = 8000
UI_MODE = 'chrome'

# Logging Settings
LOG_DIRECTORY = 'logs'
LOG_LEVEL = 'DEBUG'  # DEBUG, INFO, WARNING, ERROR, CRITICAL
LOG_FORMAT = '[%(asctime)s] [%(levelname)s] %(message)s'
LOG_DATE_FORMAT = '%H:%M:%S'

# File Settings
SUPPORTED_EXCEL_FORMATS = ['.xlsx', '.xls', '.xlsm']
SUPPORTED_IMAGE_FORMATS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.ico']

# Image Settings
DEFAULT_IMAGE_WIDTH = 200  # pixels
DEFAULT_IMAGE_HEIGHT = 150  # pixels
MAX_IMAGE_SIZE_MB = 10  # Maximum image file size

# Excel Settings
DEFAULT_CELL_POSITION = 'A1'
MAX_BATCH_SIZE = 100  # Maximum images per batch operation

# Performance Settings
ENABLE_CACHE = False
CONNECTION_TIMEOUT = 30  # seconds

# Debug Settings
DEBUG_MODE = False
VERBOSE_LOGGING = False

# Feature Flags
FEATURE_DRAG_DROP = True
FEATURE_IMAGE_PREVIEW = True
FEATURE_BATCH_PROCESSING = True
FEATURE_MERGED_CELLS = True
FEATURE_AUTO_SCALING = True
