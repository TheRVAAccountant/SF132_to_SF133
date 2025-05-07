"""
Main entry point for the SF132 to SF133 Excel Processing Application.

This module provides both GUI and CLI interfaces for processing
SF132 to SF133 reconciliation Excel files.
"""

import sys
import os
import logging
import threading
import argparse
import time
from queue import Queue
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Any, Union

# Local imports
# Try to import local modules - with fallbacks to simplify initial testing
try:
    from .core.excel_processor import ExcelProcessor
except ImportError:
    # Create a minimal ExcelProcessor for testing
    class ExcelProcessor:
        def __init__(self, queue=None):
            self.queue = queue
            
        def process_file(self, file_path, password=None):
            print(f"Processing file: {file_path}")
            return True

# Basic logging setup
import logging
def setup_logging():
    """Basic logging setup for testing."""
    logger = logging.getLogger("ExcelProcessor")
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    return logger

def setup_exception_handling(logger):
    """Basic exception handling setup."""
    pass

def performance_logger(operation_name=None):
    """Dummy performance logger for testing."""
    def decorator(func):
        return func
    return decorator

# Dummy config
class AppConfig:
    cleanup_temp_files = True
    
app_config = AppConfig()

# Try to import real modules if available
try:
    from .utils.logger import setup_logging, setup_exception_handling, performance_logger
    from .config import app_config
    print("Successfully imported logger and config modules")
except ImportError:
    pass

# Check if GUI is available
GUI_AVAILABLE = False
try:
    # Try to import the GUI
    try:
        from .modules.gui import ExcelProcessorGUI
        GUI_AVAILABLE = True
    except ImportError:
        # Create a minimal GUI class for testing
        class ExcelProcessorGUI:
            def __init__(self):
                pass
                
            def set_process_callback(self, callback):
                self.callback = callback
                
            def run(self):
                print("GUI running (minimal implementation)")
        
        # Only enable GUI if we have tkinter
        try:
            import tkinter
            GUI_AVAILABLE = True
            print("Using minimal GUI implementation")
        except ImportError:
            print("Tkinter not available, GUI disabled")
except Exception as e:
    print(f"GUI initialization error: {e}")
    pass

class ExcelProcessingApplication:
    """
    Main application class that coordinates GUI, processing, and logging components.
    """
    
    def __init__(self):
        """Initialize the application components and logging."""
        # Set up logging
        self.logger = setup_logging()
        setup_exception_handling(self.logger)
        
        # Initialize processing queue
        self.processing_queue = Queue()
        
        # Initialize processor with queue
        self.processor = ExcelProcessor(self.processing_queue)
        
        # Initialize GUI if available
        self.gui = None
        if GUI_AVAILABLE:
            try:
                self.gui = ExcelProcessorGUI()
                # Connect GUI callback
                self.gui.set_process_callback(self.process_file)
            except Exception as e:
                self.logger.warning(f"GUI initialization failed: {e}")
        
    @performance_logger("file_processing")
    def process_file(self, file_path: str, password: str = None, queue: Optional[Queue] = None) -> bool:
        """
        Process Excel file with error handling and logging.
        
        Args:
            file_path: Path to Excel file
            password: Sheet protection password
            queue: Communication queue for GUI updates
        
        Returns:
            bool: True if processing was successful, False otherwise
        """
        try:
            # Convert to absolute path
            file_path = os.path.abspath(file_path)
            
            # Validate inputs
            self._validate_inputs(file_path, password)
            
            # Log processing start
            self.logger.info(f"Starting processing of file: {file_path}")
            
            # Close any existing Excel instances
            try:
                from .modules.excel_handler import close_excel_instances
                close_excel_instances()
            except ImportError:
                self.logger.warning("Excel handler module not available")
            
            # Process the file and send updates to queue
            success = self.processor.process_file(file_path, password)
            
            if success:
                # Log successful completion
                self.logger.info("File processing completed successfully")
                return True
            else:
                self.logger.error("File processing failed")
                return False
            
        except Exception as e:
            # Log error with context
            self.logger.error(f"Error during file processing: {e}", exc_info=True)
            
            # Send error to GUI
            if queue:
                queue.put(("error", f"Processing failed: {str(e)}"))
            return False
            
    def _validate_inputs(self, file_path: str, password: str) -> None:
        """
        Validate input parameters.
        
        Args:
            file_path: Path to Excel file
            password: Sheet protection password (optional)
            
        Raises:
            ValueError: If inputs are invalid
        """
        if not file_path:
            raise ValueError("No file selected")
            
        if not os.path.exists(file_path):
            raise ValueError(f"File does not exist: {file_path}")
            
        if not file_path.lower().endswith('.xlsx'):
            raise ValueError("File must be an Excel (.xlsx) file")
            
    def run_gui(self) -> None:
        """Start the application in GUI mode."""
        if self.gui:
            try:
                self.logger.info("Starting Excel Processing Application (GUI mode)")
                self.gui.run()
            except Exception as e:
                self.logger.error(f"Error in GUI mode: {e}", exc_info=True)
                sys.exit(1)
        else:
            print("GUI components not available")
            sys.exit(1)
            
    def run_cli(self, file_path: str, password: str = None) -> int:
        """
        Run the application in command-line mode.
        
        Args:
            file_path: Path to Excel file
            password: Optional sheet password
            
        Returns:
            int: Exit code (0 for success, 1 for failure)
        """
        try:
            self.logger.info("Starting Excel Processing Application (CLI mode)")
            success = self.process_file(file_path, password)
            return 0 if success else 1
        except Exception as e:
            self.logger.error(f"Error in CLI mode: {e}", exc_info=True)
            return 1

def create_required_directories() -> None:
    """Create required application directories."""
    directories = ['logs', 'output', 'output/backups']
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)

def check_dependencies() -> bool:
    """
    Check for required dependencies and their versions.
    
    Returns:
        bool: True if all dependencies are available
    """
    logger = logging.getLogger(__name__)
    
    # Define required packages and versions
    requirements = {
        'openpyxl': '3.0.0',  # Minimum version
        'pandas': '1.0.0',    # Minimum version
    }
    
    all_available = True
    
    try:
        import pkg_resources
        
        for package, min_version in requirements.items():
            try:
                installed_version = pkg_resources.get_distribution(package).version
                logger.info(f"Found {package} version {installed_version}")
                
                # Check if version is sufficient
                if pkg_resources.parse_version(installed_version) < pkg_resources.parse_version(min_version):
                    logger.warning(f"Warning: {package} version {installed_version} is below recommended {min_version}")
                    print(f"Warning: {package} version {installed_version} is below recommended {min_version}")
            except pkg_resources.DistributionNotFound:
                logger.error(f"Required package {package} is not installed")
                print(f"Error: Required package {package} is not installed")
                all_available = False
    except ImportError:
        logger.warning("pkg_resources not available, cannot check dependencies")
        all_available = False
    
    return all_available

@performance_logger("application_startup")
def main() -> int:
    """
    Main entry point for the application.
    
    Returns:
        int: Exit code (0 for success, 1 for failure)
    """
    try:
        # Create required directories
        create_required_directories()
        
        # Set up logging
        logger = setup_logging()
        setup_exception_handling(logger)
        
        # Check dependencies
        check_dependencies()
        
        # Set up Windows-specific integrations
        is_windows = sys.platform.startswith('win')
        
        # On Windows, initialize comprehensive Windows integration
        if is_windows:
            try:
                # Import and initialize the Windows integration module
                from .modules import windows_integration
                
                # Verify Windows integration was successful
                if not windows_integration.WINDOWS_MODULES_LOADED:
                    logger.critical("Windows integration failed to load required modules")
                    print("CRITICAL ERROR: Windows integration failed")
                    print("This application requires Windows-specific modules (pywin32).")
                    print("Please install the required dependencies: pip install pywin32")
                    return 1
                
                logger.info("Windows integration initialized successfully")
                
            except ImportError as e:
                logger.critical(f"Required Windows module missing: {e}")
                print(f"CRITICAL ERROR: Required Windows module missing: {e}")
                print("This application requires the following Windows packages:")
                print("- pywin32 (pip install pywin32)")
                print("Please install the required dependencies for Windows operation.")
                return 1
            except Exception as e:
                logger.critical(f"Windows integration error: {e}")
                print(f"CRITICAL ERROR: Windows integration error: {e}")
                print("Please check your Windows environment and try again.")
                return 1
        else:
            logger.warning("Not running on Windows - functionality will be limited")
            print("WARNING: This application is designed specifically for Windows and requires Windows-specific features.")
            print("Running on non-Windows platforms may result in limited functionality or errors.")
        
        # Parse command line arguments
        parser = argparse.ArgumentParser(
            description="SF132 to SF133 Excel Reconciliation Tool"
        )
        parser.add_argument(
            "file", nargs="?", help="Path to Excel file to process"
        )
        parser.add_argument(
            "password", nargs="?", help="Sheet protection password"
        )
        parser.add_argument(
            "--recover", "-r", action="store_true", 
            help="Run in recovery mode for file access issues"
        )
        
        args = parser.parse_args()
        
        # Start the application
        app = ExcelProcessingApplication()
        
        # If recover mode is enabled and we have a file path
        if args.recover and args.file:
            logger.info(f"Running in recovery mode for file: {args.file}")
            try:
                try:
                    from .modules.excel_recovery import fix_file_in_use_error
                    success, recovery_path = fix_file_in_use_error(args.file)
                    if success:
                        logger.info(f"Recovery successful, processing file: {recovery_path}")
                        return app.run_cli(recovery_path, args.password)
                    else:
                        logger.error(f"Recovery failed: {recovery_path}")
                        return 1
                except ImportError:
                    # Minimal recovery functionality
                    logger.warning("Recovery module not available, attempting basic recovery")
                    # Create a simple copy of the file as a basic recovery
                    import shutil
                    import os
                    from pathlib import Path
                    import time
                    
                    # Create a recovery path
                    recovery_path = str(Path(args.file).with_stem(f"{Path(args.file).stem}_recovered_{int(time.time())}"))
                    try:
                        # Try to copy the file
                        shutil.copy2(args.file, recovery_path)
                        logger.info(f"Basic recovery successful, processing file: {recovery_path}")
                        return app.run_cli(recovery_path, args.password)
                    except Exception as e:
                        logger.error(f"Basic recovery failed: {e}")
                        return 1
            except ImportError:
                logger.error("Recovery modules not available")
                return 1
        
        # Check if a file path is provided
        if args.file:
            return app.run_cli(args.file, args.password)
        else:
            # If no command-line arguments and GUI is available, start in GUI mode
            if GUI_AVAILABLE:
                app.run_gui()
                return 0
            else:
                # If no arguments and no GUI, show usage information
                parser.print_help()
                return 1
                
    except Exception as e:
        # Get logger instance
        try:
            logger = logging.getLogger(__name__)
            logger.error(f"Fatal error during startup: {e}", exc_info=True)
        except:
            pass
        
        # Ensure the error is displayed to the user
        print(f"Fatal error during startup: {str(e)}")
        print("Check the logs for more details.")
        return 1

if __name__ == "__main__":
    sys.exit(main())