"""
File operations for Excel processing.

This module provides robust file handling functions for Excel files,
with special attention to Windows-specific file locking issues.
"""

import os
import sys
import time
import shutil
import tempfile
import uuid
import logging
import gc
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple, Union
import contextlib

# Import platform-specific modules
try:
    import pythoncom
    import win32com.client
    WINDOWS_MODULES_AVAILABLE = True
except ImportError:
    WINDOWS_MODULES_AVAILABLE = False

# Type aliases
PathLike = Union[str, Path]

# Initialize logger
logger = logging.getLogger(__name__)

# Global tracking of temporary files
_temp_files = []

def register_temp_file(file_path: str) -> None:
    """
    Register a temporary file for later cleanup.
    
    Args:
        file_path: Path to the temporary file
    """
    if file_path not in _temp_files:
        _temp_files.append(file_path)

def cleanup_temp_files() -> None:
    """Clean up all registered temporary files."""
    for temp_file in list(_temp_files):
        if os.path.exists(temp_file):
            try:
                os.unlink(temp_file)
                logger.debug(f"Cleaned up temp file: {temp_file}")
                _temp_files.remove(temp_file)
            except Exception as e:
                logger.warning(f"Failed to clean up temp file {temp_file}: {e}")

def get_temp_file_path(prefix: str = "excel") -> str:
    """
    Generate a temporary file path.
    
    Args:
        prefix: Prefix for temporary file
        
    Returns:
        str: Temporary file path
    """
    temp_dir = tempfile.gettempdir()
    unique_id = str(uuid.uuid4())
    temp_file = os.path.join(temp_dir, f"{prefix}_temp_{unique_id}.xlsx")
    register_temp_file(temp_file)  # Register for cleanup
    return temp_file

def create_backup_file(file_path: str, backup_dir: Optional[str] = None) -> str:
    """
    Create a backup copy of the file.
    
    Args:
        file_path: Path to file to backup
        backup_dir: Optional backup directory path
        
    Returns:
        str: Path to backup file
    """
    if not backup_dir:
        backup_dir = os.path.join(os.path.dirname(file_path), "backups")
    
    # Create backup directory
    os.makedirs(backup_dir, exist_ok=True)
    
    # Create backup file path with timestamp
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    filename = os.path.basename(file_path)
    backup_path = os.path.join(backup_dir, 
                              f"{Path(filename).stem}_backup_{timestamp}{Path(filename).suffix}")
    
    # Copy the file
    shutil.copy2(file_path, backup_path)
    return backup_path

def close_excel_instances() -> List[int]:
    """
    Terminate all existing Excel processes to prevent file locking.
    
    Returns:
        List[int]: List of terminated process IDs
    """
    terminated_pids = []
    
    # Import psutil here to avoid dependency issues
    try:
        import psutil
        
        # First find all Excel processes
        excel_pids = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
                    excel_pids.append(proc.info['pid'])
            except Exception:
                pass
        
        if excel_pids:
            logger.info(f"Found {len(excel_pids)} Excel processes to close")
            
        # Then terminate each process with proper cleanup
        for pid in excel_pids:
            try:
                proc = psutil.Process(pid)
                proc.terminate()
                proc.wait(timeout=5)
                terminated_pids.append(pid)
            except Exception as e:
                logger.warning(f"Failed to terminate Excel process (PID {pid}): {e}")
                # Try more aggressive termination
                try:
                    if psutil.pid_exists(pid):
                        os.kill(pid, 9)  # SIGKILL
                        terminated_pids.append(pid)
                except Exception:
                    pass
        
        # Verify all processes are terminated
        remaining = []
        for pid in excel_pids:
            if psutil.pid_exists(pid):
                remaining.append(pid)
                
        if remaining:
            logger.warning(f"Could not terminate {len(remaining)} Excel processes: {remaining}")
        else:
            logger.info("All Excel processes successfully closed")
            
    except ImportError:
        logger.warning("psutil module not available, cannot close Excel instances")
    
    return terminated_pids

def safe_file_copy(source_path: str, dest_path: str, retries: int = 3, delay: float = 1.0) -> bool:
    """
    Safely copy a file with retries.
    
    Args:
        source_path: Source file path
        dest_path: Destination file path
        retries: Number of retry attempts
        delay: Delay between retries
        
    Returns:
        bool: Whether the copy was successful
    """
    # Ensure destination directory exists
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    
    for attempt in range(retries):
        try:
            # Close any Excel instances before attempting copy
            if WINDOWS_MODULES_AVAILABLE and sys.platform.startswith('win'):
                close_excel_instances()
                
            # Perform the copy
            shutil.copy2(source_path, dest_path)
            
            # Verify the copy succeeded
            if not os.path.exists(dest_path):
                raise IOError(f"Failed to copy file to {dest_path}")
                
            file_size = os.path.getsize(dest_path)
            if file_size == 0:
                raise IOError(f"Copied file has zero size: {dest_path}")
                
            logger.info(f"Successfully copied file to {dest_path} ({file_size} bytes)")
            return True
            
        except Exception as e:
            logger.warning(f"Copy attempt {attempt + 1}/{retries} failed: {e}")
            if attempt < retries - 1:
                # Wait before retrying
                time.sleep(delay)
                # Increase delay for next retry
                delay *= 1.5
            else:
                logger.error(f"Failed to copy file after {retries} attempts: {e}")
                return False

def create_verified_copy(original_file: str) -> str:
    """
    Create a verified copy of the original file to prevent corruption.
    
    Args:
        original_file: Path to original Excel file
        
    Returns:
        str: Path to verified copy
    """
    temp_copy = get_temp_file_path("verified_copy")
    logger.info(f"Creating verified copy at {temp_copy}...")
    
    # Try creating a clean copy using different methods
    
    # Method 1: Try using COM if available on Windows
    if WINDOWS_MODULES_AVAILABLE and sys.platform.startswith('win'):
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Start Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Open workbook with recovery options
            wb = excel.Workbooks.Open(
                original_file,
                UpdateLinks=0,
                ReadOnly=True,
                CorruptLoad=2  # xlRepairFile (better corruption handling)
            )
            
            # Save as a new clean file
            wb.SaveAs(
                temp_copy,
                FileFormat=51,  # xlOpenXMLWorkbook
                CreateBackup=False
            )
            
            # Clean close
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            # Force cleanup
            del wb
            del excel
            gc.collect()
            pythoncom.CoUninitialize()
            
            # Verify file exists and has content
            if not os.path.exists(temp_copy) or os.path.getsize(temp_copy) == 0:
                raise ValueError("Failed to create valid copy")
                
            return temp_copy
            
        except Exception as e:
            logger.warning(f"COM copy failed: {e}, falling back to direct copy")
    
    # Method 2: Direct copy as fallback
    safe_file_copy(original_file, temp_copy)
    return temp_copy

def validate_excel_file(file_path: str) -> bool:
    """
    Validate an Excel file by trying to open it and check for errors.
    
    Args:
        file_path: Path to Excel file to validate
        
    Returns:
        bool: Whether the file is valid
    """
    try:
        # Method 1: Try opening with openpyxl
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True)
        wb.close()
        
        # Method 2: Verify with Excel COM if possible
        if WINDOWS_MODULES_AVAILABLE and sys.platform.startswith('win'):
            try:
                # Initialize COM
                pythoncom.CoInitialize()
                
                # Start Excel
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                try:
                    # Try to open the file
                    wb = excel.Workbooks.Open(
                        file_path,
                        UpdateLinks=0,
                        ReadOnly=True
                    )
                    
                    # Check if there are any error messages
                    has_errors = excel.ErrorCheckingStatus if hasattr(excel, 'ErrorCheckingStatus') else False
                    
                    # Close without saving
                    wb.Close(SaveChanges=False)
                    excel.Quit()
                    
                    # Cleanup COM objects
                    del wb
                    del excel
                    gc.collect()
                    pythoncom.CoUninitialize()
                    
                    if has_errors:
                        logger.warning(f"Excel detected errors in {file_path}")
                        return False
                        
                except Exception as e:
                    logger.warning(f"COM validation failed: {e}")
                    return False
                    
        # If we got here, file seems valid
        return True
        
    except Exception as e:
        logger.warning(f"File validation failed: {e}")
        return False

@contextlib.contextmanager
def safe_open_excel_file(file_path: str, read_only: bool = True):
    """
    Safely open an Excel file with proper error handling.
    
    Args:
        file_path: Path to Excel file
        read_only: Whether to open in read-only mode
        
    Yields:
        openpyxl.Workbook: Open workbook or None if error
    """
    wb = None
    try:
        import openpyxl
        # Wait a moment for filesystem to settle
        time.sleep(0.2)
        wb = openpyxl.load_workbook(file_path, read_only=read_only)
        yield wb
    except Exception as e:
        logger.warning(f"Failed to open workbook {file_path}: {e}")
        yield None
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass