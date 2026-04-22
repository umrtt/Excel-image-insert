"""
Excel Image Insert - Main Entry Point
Desktop GUI application for inserting images into Excel files using Eel
"""
import eel
import os
from pathlib import Path

# Import application controller
from app.controllers import AppController
from app.utils import Logger

# Initialize logger
logger = Logger(__name__)

# Initialize application controller (dependency injection)
app_controller = AppController()

# ===== Eel API Endpoints =====
# These methods are exposed to the JavaScript frontend

# ===== File Operations =====

@eel.expose
def select_excel_file():
    """
    Eel endpoint: Select Excel file
    Exposed to JavaScript frontend
    """
    logger.info("select_excel_file endpoint called")
    return app_controller.select_excel_file()


@eel.expose
def get_sheets():
    """
    Eel endpoint: Get available sheets from open workbook
    """
    logger.info("get_sheets endpoint called")
    return app_controller.get_sheets()


@eel.expose
def close_excel_file():
    """
    Eel endpoint: Close current Excel file
    """
    logger.info("close_excel_file endpoint called")
    return app_controller.close_excel_file()


# ===== Image Operations =====

@eel.expose
def select_images():
    """
    Eel endpoint: Select image files
    """
    logger.info("select_images endpoint called")
    return app_controller.select_images()


@eel.expose
def add_image(file_path, sheet_name, cell_position, width=None, height=None):
    """
    Eel endpoint: Add image to insertion list
    
    Args:
        file_path: Path to image file
        sheet_name: Target sheet name
        cell_position: Target cell position
        width: Optional width in pixels
        height: Optional height in pixels
    """
    logger.info(f"add_image endpoint called: {file_path} at {sheet_name}!{cell_position}")
    return app_controller.add_image(
        file_path, sheet_name, cell_position, width, height
    )


@eel.expose
def remove_image(index):
    """
    Eel endpoint: Remove image from list
    
    Args:
        index: Image index
    """
    logger.info(f"remove_image endpoint called: index {index}")
    return app_controller.remove_image(index)


@eel.expose
def clear_images():
    """
    Eel endpoint: Clear all images
    """
    logger.info("clear_images endpoint called")
    return app_controller.clear_images()


@eel.expose
def get_images():
    """
    Eel endpoint: Get all images in list
    """
    logger.debug("get_images endpoint called")
    return app_controller.get_images()


# ===== Insertion Operations =====

@eel.expose
def insert_images():
    """
    Eel endpoint: Insert all images into Excel file
    """
    logger.info("insert_images endpoint called")
    return app_controller.insert_images()


@eel.expose
def insert_single_image(file_path, sheet_name, cell_position, width=None, height=None):
    """
    Eel endpoint: Insert single image directly
    
    Args:
        file_path: Path to image file
        sheet_name: Target sheet name
        cell_position: Target cell position
        width: Optional width in pixels
        height: Optional height in pixels
    """
    logger.info(f"insert_single_image endpoint called: {file_path} at {sheet_name}!{cell_position}")
    return app_controller.insert_single_image(
        file_path, sheet_name, cell_position, width, height
    )


# ===== Status Operations =====

@eel.expose
def get_status():
    """
    Eel endpoint: Get application status
    """
    logger.debug("get_status endpoint called")
    return app_controller.get_status()


@eel.expose
def select_range_from_excel():
    """
    Eel endpoint: Get currently selected range from Excel
    User clicks button to capture the selection
    """
    logger.info("select_range_from_excel endpoint called")
    return app_controller.select_range_from_excel()


# ===== Application Startup =====

def get_web_path():
    """Get path to web directory."""
    current_dir = Path(__file__).parent
    web_dir = current_dir / 'web'
    return str(web_dir)


def main():
    """
    Main entry point for the application.
    Initializes Eel and starts the GUI.
    """
    try:
        # Set up Eel
        web_path = get_web_path()
        eel.init(web_path)
        
        logger.info("=" * 60)
        logger.info("Excel Image Insert Application Starting")
        logger.info("=" * 60)
        logger.info(f"Web directory: {web_path}")
        logger.info("Starting Eel GUI application...")
        
        # Start the application
        # chrome_args = ['--headless', '--disable-gpu']
        eel.start(
            'index.html',
            mode='chrome',
            port=8000,
            size=(1000, 900),
            position=(100, 100),
            disable_cache=True
        )
        
    except Exception as e:
        logger.critical(f"Failed to start application: {str(e)}")
        raise


if __name__ == '__main__':
    main()
