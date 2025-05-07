import sys
import os
from pathlib import Path
from typing import Optional
import threading
from queue import Queue
import traceback

from gui import ExcelProcessorGUI
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
        
        # Initialize GUI
        self.gui = ExcelProcessorGUI()
        
        # Set up processor with queue
        self.processor = ExcelProcessor(self.processing_queue)
        
        # Connect GUI callback
        self.gui.set_process_callback(self.process_file)
        
    @performance_logger("file_processing")
    def process_file(self, file_path: str, password: str, queue: Queue) -> None:
        """
        Process Excel file with error handling and logging.
        
        Args:
            file_path (str): Path to Excel file
            password (str): Sheet protection password
            queue (Queue): Communication queue for GUI updates
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
            self.processor.process_file(file_path, password)
            
            # Log successful completion
            self.logger.info("File processing completed successfully")
            
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
            queue.put(("error", f"Processing failed: {str(e)}"))
            
    def _validate_inputs(self, file_path: str, password: str) -> None:
        """
        Validate input parameters.
        
        Args:
            file_path (str): Path to Excel file
            password (str): Sheet protection password
            
        Raises:
            ValueError: If inputs are invalid
        """
        if not file_path:
            raise ValueError("No file selected")
            
        if not password:
            raise ValueError("Password is required")
            
        if not os.path.exists(file_path):
            raise ValueError(f"File does not exist: {file_path}")
            
        if not file_path.lower().endswith('.xlsx'):
            raise ValueError("File must be an Excel (.xlsx) file")
            
    def run(self) -> None:
        """Start the application."""
        try:
            self.logger.info("Starting Excel Processing Application")
            self.gui.run()
        except Exception as e:
            self.error_handler.log_exception(e, {'operation': 'application_startup'})
            sys.exit(1)
            
def create_required_directories() -> None:
    """Create required application directories."""
    directories = ['logs', 'output']
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
    
@performance_logger("application_startup")
def main() -> None:
    """Main entry point for the application."""
    try:
        # Create required directories
        create_required_directories()
        
        # Set up global exception handling
        setup_exception_handling()
        
        # Start the application
        app = ExcelProcessingApplication()
        app.run()
        
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
        sys.exit(1)
        
if __name__ == "__main__":
    main()