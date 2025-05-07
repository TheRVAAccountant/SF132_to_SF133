"""
Compatibility layer for the refactored SF132 to SF133 reconciliation tool.

This module allows existing code to import the original module names
while using the refactored package implementation.
"""

import sys
import logging
import warnings

# Setup basic logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("compat")

# Check if running on Windows - this application requires Windows
if not sys.platform.startswith('win'):
    logger.warning("This application is designed specifically for Windows and may not function correctly on other platforms.")
    print("WARNING: This application requires Windows to function properly.")
    print("Running on a non-Windows platform may result in limited functionality or errors.")

# Import refactored components
try:
    # Core components
    from src.sf132_sf133_recon.core.excel_processor import ExcelProcessor
    
    # Module components
    from src.sf132_sf133_recon.modules.gui import ExcelProcessorGUI
    from src.sf132_sf133_recon.modules.excel_handler import close_excel_instances
    from src.sf132_sf133_recon.modules.excel_recovery import fix_file_in_use_error
    from src.sf132_sf133_recon.modules.file_operations import (
        create_verified_copy, safe_file_copy, cleanup_temp_files
    )
    
    # Utility components
    from src.sf132_sf133_recon.utils.logger import (
        setup_logging, performance_logger, setup_exception_handling, log_exception
    )
    
    # Config components
    from src.sf132_sf133_recon.config import app_config, ProcessingConfig, FileHandlingConfig, ExcelConfig
    
    # Main application
    from src.sf132_sf133_recon.main import ExcelProcessingApplication
    
    logger.info("Successfully loaded refactored components")
    
except ImportError as e:
    logger.error(f"Error importing refactored components: {e}")
    
    # On Windows, Windows-specific modules are required
    if sys.platform.startswith('win'):
        logger.critical("Windows-specific modules are required but not found")
        print("CRITICAL ERROR: Windows-specific modules required but not found")
        print("Please install the required dependencies for Windows operation:")
        print("pip install pywin32 openpyxl pandas psutil")
        raise ImportError("Windows-specific modules are required but not found") from e
    
    logger.warning("Falling back to original modules")
    
    # Fall back to original imports if refactored modules aren't available
    try:
        from excel_processor import ExcelProcessor
        from gui import ExcelProcessorGUI
        from main import ExcelProcessingApplication
        from logger_config import LoggerSetup, ErrorHandler, performance_logger
        from config import app_config
        from excel_processor_config import ProcessingConfig, FileHandlingConfig, ExcelConfig
        
        warnings.warn(
            "Using original module imports. Consider updating your code to use the "
            "refactored package structure for better maintainability.",
            DeprecationWarning, stacklevel=2
        )
    except ImportError as original_import_error:
        logger.critical(f"Failed to import required modules: {original_import_error}")
        print(f"CRITICAL ERROR: Required modules missing: {original_import_error}")
        print("Please install all required dependencies:")
        print("pip install -r requirements.txt")
        raise

# Provide backward compatibility for specific functions
def process_excel_file(file_path, password=None):
    """
    Process an Excel file using the current implementation.
    
    This function provides backward compatibility with code that
    might have directly called a function to process files.
    
    Args:
        file_path: Path to the Excel file
        password: Optional password for protected sheets
        
    Returns:
        bool: True if successful, False otherwise
    """
    app = ExcelProcessingApplication()
    return app.process_file(file_path, password)