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
from .core.excel_processor import ExcelProcessor
from .utils.logger import setup_logging, setup_exception_handling, performance_logger
from .config import app_config

# Check if GUI is available
GUI_AVAILABLE = False
try:
    from .modules.gui import ExcelProcessorGUI
    GUI_AVAILABLE = True
except ImportError:
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
            from .modules.excel_handler import close_excel_instances
            close_excel_instances()
            
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
        
        # Set up Windows-specific integrations if running on Windows
        if sys.platform.startswith('win'):
            try:
                from .utils.win_path_handler import normalize_windows_path
                logger.info("Windows-specific path handling enabled")
            except ImportError:
                logger.warning("Windows path handling not available")
        
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
                from .modules.excel_recovery import fix_file_in_use_error
                success, recovery_path = fix_file_in_use_error(args.file)
                if success:
                    logger.info(f"Recovery successful, processing file: {recovery_path}")
                    return app.run_cli(recovery_path, args.password)
                else:
                    logger.error(f"Recovery failed: {recovery_path}")
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