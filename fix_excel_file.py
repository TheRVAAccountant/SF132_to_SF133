"""
Command-line utility for fixing Excel file access issues on Windows.
This script provides a direct interface to the recovery functionality.
"""

import os
import sys
import argparse
import logging
import time
from typing import Optional

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('excel_recovery.log')
    ]
)

logger = logging.getLogger("ExcelRecovery")

def setup_modules():
    """Import all required modules and set up resources."""
    # Add current directory to path to ensure imports work
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
        
    # Import the recovery modules
    try:
        from file_operations_win import setup_resource_tracking, close_excel_instances
        from win_path_handler import normalize_windows_path
        from excel_file_recovery import fix_excel_file_in_use_error
        
        # Register cleanup handlers
        setup_resource_tracking()
        
        return True
    except ImportError as e:
        logger.error(f"Failed to import required modules: {e}")
        return False

def fix_excel_file(file_path: str, output_dir: Optional[str] = None) -> bool:
    """
    Fix Excel file access issues.
    
    Args:
        file_path: Path to Excel file with access issues
        output_dir: Optional output directory for fixed file
        
    Returns:
        bool: Whether the fix was successful
    """
    logger.info(f"Starting Excel file fix for: {file_path}")
    
    # First make sure the file exists
    if not os.path.exists(file_path):
        logger.error(f"File does not exist: {file_path}")
        return False
        
    # Import recovery modules
    if not setup_modules():
        return False
    
    # Now import specific functions (after setup)
    from excel_file_recovery import fix_excel_file_in_use_error
    from win_path_handler import normalize_windows_path
    from file_operations_win import close_excel_instances
    
    # Normalize file path
    file_path = normalize_windows_path(file_path)
    
    # Close any Excel instances first
    logger.info("Closing Excel instances...")
    close_excel_instances()
    
    # Apply the fix
    logger.info("Applying file access fix...")
    success, result_path = fix_excel_file_in_use_error(file_path, output_dir)
    
    if success:
        logger.info(f"Successfully fixed file. Output at: {result_path}")
        return True
    else:
        logger.error(f"Failed to fix file: {result_path}")
        return False

def main():
    """Main entry point for the command-line utility."""
    parser = argparse.ArgumentParser(description="Fix Excel file access issues on Windows")
    parser.add_argument("file_path", help="Path to the Excel file to fix")
    parser.add_argument("--output-dir", "-o", help="Optional output directory for fixed file")
    
    args = parser.parse_args()
    
    # Start timer
    start_time = time.time()
    
    # Run the fix
    success = fix_excel_file(args.file_path, args.output_dir)
    
    # Report elapsed time
    elapsed = time.time() - start_time
    logger.info(f"Process completed in {elapsed:.2f} seconds")
    
    # Return appropriate exit code
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())