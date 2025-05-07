import sys
import os
from pathlib import Path
from typing import Optional
import threading
from queue import Queue
import traceback
import logging
import time
import json
import pkg_resources

try:
    from gui import ExcelProcessorGUI
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

from excel_processor import ExcelProcessor
from logger_config import LoggerSetup, ErrorHandler, performance_logger

class ExcelProcessingApplication:
    """
    Main application class that coordinates GUI, processing, and logging components.
    """
    
    def __init__(self):
        """Initialize the application components and logging."""
        # Set up logging
        self.logger_setup = LoggerSetup()
        self.logger = self.logger_setup.setup_logging()
        self.error_handler = ErrorHandler(self.logger)
        
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
    def process_file(self, file_path: str, password: str = None, queue: Queue = None) -> bool:
        """
        Process Excel file with error handling and logging.
        
        Args:
            file_path (str): Path to Excel file
            password (str): Sheet protection password
            queue (Queue): Communication queue for GUI updates
        
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
            self.processor.close_excel_instances()
            
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
            self.error_handler.log_exception(
                e,
                {
                    'file_path': file_path,
                    'operation': 'file_processing'
                }
            )
            # Send error to GUI
            if queue:
                queue.put(("error", f"Processing failed: {str(e)}"))
            return False
            
    def _validate_inputs(self, file_path: str, password: str) -> None:
        """
        Validate input parameters.
        
        Args:
            file_path (str): Path to Excel file
            password (str): Sheet protection password (optional)
            
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
                self.error_handler.log_exception(e, {'operation': 'application_startup_gui'})
                sys.exit(1)
        else:
            print("GUI components not available")
            sys.exit(1)
            
    def run_cli(self, file_path: str, password: str = None) -> int:
        """
        Run the application in command-line mode.
        
        Args:
            file_path (str): Path to Excel file
            password (str): Optional sheet password
            
        Returns:
            int: Exit code (0 for success, 1 for failure)
        """
        try:
            self.logger.info("Starting Excel Processing Application (CLI mode)")
            success = self.process_file(file_path, password)
            return 0 if success else 1
        except Exception as e:
            self.error_handler.log_exception(e, {'operation': 'application_startup_cli'})
            return 1
            
def create_required_directories() -> None:
    """Create required application directories."""
    directories = ['logs', 'output', 'output/backups']
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)
        
def setup_exception_handling() -> None:
    """Set up global exception handling."""
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            # Handle keyboard interrupt differently
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
            
        logger = LoggerSetup().setup_logging()
        logger.critical("Uncaught exception:", exc_info=(exc_type, exc_value, exc_traceback))
        
    sys.excepthook = handle_exception
    
    def handle_thread_exception(args):
        logger = LoggerSetup().setup_logging()
        logger.critical("Uncaught thread exception:", exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        
    threading.excepthook = handle_thread_exception

def check_dependencies():
    """Check for required dependencies and their versions."""
    logger = logging.getLogger("ExcelProcessor")
    
    # Define required packages and versions
    requirements = {
        'openpyxl': '3.0.0',  # Minimum version
        'pandas': '1.0.0',    # Minimum version
    }
    
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
            return False
    
    return True

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
        
        # Set up global exception handling
        setup_exception_handling()
        
        # Check dependencies
        check_dependencies()
        
        # Setup minimal logging for CLI mode
        logger = logging.getLogger("ExcelProcessor")
        if not logger.handlers:
            logger.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            
            # Add file handler
            file_handler = logging.FileHandler('excel_processor.log')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        # Set up Windows-specific integrations if running on Windows
        if sys.platform.startswith('win'):
            try:
                import windows_integration
                logger.info("Windows-specific enhancements enabled")
            except ImportError:
                logger.warning("Windows integration module not available")
        
        # Parse command line arguments
        recover_mode = False
        file_path = None
        password = None
        
        # Process command line arguments
        args = sys.argv[1:]
        i = 0
        while i < len(args):
            if args[i] in ('--recover', '-r'):
                recover_mode = True
                i += 1
            elif args[i].startswith('--'):
                # Skip unknown options and their values
                i += 2
            else:
                # First non-option argument is the file path
                file_path = args[i]
                # If there's another argument, it's the password
                if i + 1 < len(args) and not args[i+1].startswith('--'):
                    password = args[i+1]
                break
        
        # Start the application
        app = ExcelProcessingApplication()
        
        # If recover mode is enabled and we have a file path
        if recover_mode and file_path:
            logger.info(f"Running in recovery mode for file: {file_path}")
            try:
                from excel_file_recovery import fix_excel_file_in_use_error
                success, recovery_path = fix_excel_file_in_use_error(file_path)
                if success:
                    logger.info(f"Recovery successful, processing file: {recovery_path}")
                    return app.run_cli(recovery_path, password)
                else:
                    logger.error(f"Recovery failed: {recovery_path}")
                    return 1
            except ImportError:
                logger.error("Recovery modules not available")
                return 1
        
        # Check if a file path is provided as argument
        if file_path:
            return app.run_cli(file_path, password)
        else:
            # If no command-line arguments and GUI is available, start in GUI mode
            if GUI_AVAILABLE:
                app.run_gui()
                return 0
            else:
                # If no arguments and no GUI, show usage information
                print("Usage: python main.py [--recover] [excel_file_path] [optional_password]")
                return 1
                
    except Exception as e:
        # Get logger instance
        logger = LoggerSetup().setup_logging()
        error_handler = ErrorHandler(logger)
        
        # Log the error
        error_handler.log_exception(
            e,
            {'operation': 'application_startup'}
        )
        
        # Ensure the error is displayed to the user
        print(f"Fatal error during startup: {str(e)}")
        print("Check the logs for more details.")
        return 1
        
if __name__ == "__main__":
    sys.exit(main())