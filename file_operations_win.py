"""
Enhanced file operations module for Windows compatibility with robust file handle cleanup.
This module specifically addresses issues with files being locked by system processes.
"""

import os
import sys
import time
import shutil
import logging
import traceback
import tempfile
import psutil
import pythoncom
import win32com.client
import openpyxl
import gc
import atexit
import contextlib
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

# Initialize logger
logger = logging.getLogger(__name__)

# Global tracking of file handles and resources
_temp_files = []
_open_handles = {}

def setup_resource_tracking():
    """Register cleanup handlers for application exit."""
    atexit.register(cleanup_all_resources)
    
def cleanup_all_resources():
    """Clean up all tracked resources when the application exits."""
    # Close COM objects first
    for handle_id, handle_info in list(_open_handles.items()):
        close_resource(handle_id)
    
    # Then clean temp files
    cleanup_temp_files()

def track_resource(resource_type: str, resource_obj: Any, resource_path: Optional[str] = None) -> str:
    """
    Track a resource for later cleanup.
    
    Args:
        resource_type: Type of resource (e.g., 'excel', 'workbook', 'file')
        resource_obj: The resource object to track
        resource_path: Optional file path associated with the resource
        
    Returns:
        str: Unique ID for the resource
    """
    import uuid
    handle_id = str(uuid.uuid4())
    _open_handles[handle_id] = {
        'type': resource_type,
        'object': resource_obj,
        'path': resource_path,
        'created': time.time()
    }
    return handle_id

def close_resource(handle_id: str) -> bool:
    """
    Close and clean up a tracked resource.
    
    Args:
        handle_id: Resource handle ID
        
    Returns:
        bool: Whether cleanup was successful
    """
    if handle_id not in _open_handles:
        return False
        
    handle_info = _open_handles.pop(handle_id)
    resource_type = handle_info['type']
    resource_obj = handle_info['object']
    
    try:
        if resource_type == 'excel':
            # Excel COM application object
            try:
                resource_obj.Quit()
            except:
                pass
            finally:
                del resource_obj
                
        elif resource_type == 'workbook':
            # Excel workbook COM object
            try:
                resource_obj.Close(SaveChanges=False)
            except:
                pass
            finally:
                del resource_obj
                
        elif resource_type == 'openpyxl_workbook':
            # openpyxl workbook object
            try:
                resource_obj.close()
            except:
                pass
                
        elif resource_type == 'file_handle':
            # Open file handle
            try:
                resource_obj.close()
            except:
                pass
                
        gc.collect()  # Force Python garbage collection
        return True
        
    except Exception as e:
        logger.warning(f"Error closing resource {handle_id}: {e}")
        return False

def add_temp_file(file_path: str):
    """
    Add a file to the temporary files list for later cleanup.
    
    Args:
        file_path: Path to the temporary file
    """
    if file_path not in _temp_files:
        _temp_files.append(file_path)
        
def cleanup_temp_files():
    """Clean up all tracked temporary files."""
    if not _temp_files:
        return
        
    logger.info(f"Cleaning up {len(_temp_files)} temporary files")
    for temp_file in list(_temp_files):
        try:
            if os.path.exists(temp_file):
                # Try to forcibly unlock the file first on Windows
                unlock_file(temp_file)
                
                # Now try to delete it
                os.unlink(temp_file)
                logger.debug(f"Removed temp file: {temp_file}")
                _temp_files.remove(temp_file)
        except Exception as e:
            logger.warning(f"Failed to clean up temp file {temp_file}: {e}")

def unlock_file(file_path: str) -> bool:
    """
    Attempt to unlock a file by closing any processes that have it open (Windows specific).
    
    Args:
        file_path: Path to the file to unlock
        
    Returns:
        bool: Whether unlock was successful
    """
    try:
        # First close Excel instances that might have the file open
        close_excel_instances()
        
        # On Windows we can try to forcibly close handles
        if 'win32' in sys.platform or 'win64' in sys.platform:
            file_path = os.path.abspath(file_path)
            
            # Try closing COM connections to the file
            pythoncom.CoFreeUnusedLibraries()
            gc.collect()  # Force Python garbage collection
            
            # Wait a moment for handles to release
            time.sleep(0.5)
            
            # Use psutil to check which processes have the file open
            import psutil
            target_path = os.path.normcase(os.path.normpath(file_path))
            for proc in psutil.process_iter(['pid', 'name', 'open_files']):
                try:
                    open_files = proc.info.get('open_files', [])
                    if open_files:
                        for file_info in open_files:
                            if os.path.normcase(os.path.normpath(file_info.path)) == target_path:
                                logger.info(f"Found process {proc.info['name']} (PID {proc.info['pid']}) with open handle to {file_path}")
                                proc.kill()
                                return True
                except (psutil.AccessDenied, psutil.NoSuchProcess, Exception):
                    continue
                    
            return True
    except Exception as e:
        logger.warning(f"Error unlocking file {file_path}: {e}")
        return False

def close_excel_instances() -> List[int]:
    """
    Close all running Excel instances to prevent file locking.
    
    Returns:
        List[int]: List of terminated process IDs
    """
    terminated_pids = []
    
    # First find all Excel processes
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
                pid = proc.info['pid']
                logger.info(f"Attempting to terminate Excel process {pid}")
                
                try:
                    process = psutil.Process(pid)
                    process.terminate()
                    try:
                        process.wait(timeout=5)
                        terminated_pids.append(pid)
                    except psutil.TimeoutExpired:
                        # If termination times out, try to kill
                        process.kill()
                        terminated_pids.append(pid)
                except (psutil.NoSuchProcess, psutil.AccessDenied, Exception) as e:
                    logger.warning(f"Failed to terminate Excel process (PID {pid}): {e}")
                    
        except Exception:
            pass
            
    # Force COM cleanup
    pythoncom.CoFreeUnusedLibraries()
    gc.collect()
    
    # Wait for processes to fully terminate
    time.sleep(1)
    
    return terminated_pids

def get_temp_file_path(prefix: str = "excel", suffix: str = ".xlsx") -> str:
    """
    Generate a temporary file path with proper Windows compatibility.
    
    Args:
        prefix: Prefix for temporary file
        suffix: Suffix for temporary file
        
    Returns:
        str: Temporary file path
    """
    # Use tempfile.mkstemp to get unique filename
    fd, temp_path = tempfile.mkstemp(suffix=suffix, prefix=f"{prefix}_", dir=None)
    
    # Close the file descriptor immediately to avoid handle leaks
    os.close(fd)
    
    # Track for later cleanup
    add_temp_file(temp_path)
    
    return temp_path

@contextlib.contextmanager
def safe_file_operation(file_path: str):
    """
    Context manager for safely operating on a file with error handling.
    
    Args:
        file_path: Path to the file
        
    Yields:
        str: The file path
    """
    try:
        # Ensure the file isn't locked
        unlock_file(file_path)
        
        # Yield the path for operations
        yield file_path
        
    except Exception as e:
        logger.error(f"Error during file operation on {file_path}: {e}")
        raise
    finally:
        # Force resource cleanup
        pythoncom.CoFreeUnusedLibraries()
        gc.collect()

def safe_copy_file(source_path: str, dest_path: str, retries: int = 3, delay: float = 1.0) -> bool:
    """
    Safely copy a file with retries and proper Windows file handle management.
    
    Args:
        source_path: Source file path
        dest_path: Destination file path
        retries: Number of retry attempts
        delay: Delay between retries in seconds
        
    Returns:
        bool: Whether the copy was successful
    """
    # Create destination directory if it doesn't exist
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    
    # Normalize paths
    source_path = os.path.abspath(source_path)
    dest_path = os.path.abspath(dest_path)
    
    logger.info(f"Copying file from {source_path} to {dest_path}")
    
    for attempt in range(retries):
        try:
            # Close any processes that might have the files open
            unlock_file(source_path)
            if os.path.exists(dest_path):
                unlock_file(dest_path)
                
            # Create a temporary file first, then move to final location
            temp_dest = get_temp_file_path(prefix="copy_temp")
            
            # Perform the copy to temp location
            shutil.copy2(source_path, temp_dest)
            
            # Verify the copy succeeded
            if not os.path.exists(temp_dest) or os.path.getsize(temp_dest) == 0:
                raise IOError(f"Copy verification failed for {temp_dest}")
                
            # If dest exists, try to remove it
            if os.path.exists(dest_path):
                try:
                    os.unlink(dest_path)
                except PermissionError:
                    logger.warning(f"Could not remove existing destination file {dest_path}")
                    unlock_file(dest_path)
                    os.unlink(dest_path)
                    
            # Move temp file to final destination
            shutil.move(temp_dest, dest_path)
            
            logger.info(f"Successfully copied file to {dest_path}")
            return True
            
        except Exception as e:
            logger.warning(f"Copy attempt {attempt + 1}/{retries} failed: {e}")
            if attempt < retries - 1:
                time.sleep(delay)
                # Increase delay for next retry
                delay *= 1.5
            else:
                logger.error(f"Failed to copy file after {retries} attempts: {e}")
                return False

@contextlib.contextmanager
def safe_open_workbook(file_path: str, read_only: bool = True):
    """
    Safely open an Excel workbook with proper resource cleanup.
    
    Args:
        file_path: Path to Excel file
        read_only: Whether to open in read-only mode
        
    Yields:
        openpyxl.Workbook: The open workbook
    """
    workbook = None
    handle_id = None
    
    try:
        # Ensure file isn't locked
        unlock_file(file_path)
        
        # Open the workbook
        workbook = openpyxl.load_workbook(file_path, read_only=read_only)
        
        # Track the resource
        handle_id = track_resource('openpyxl_workbook', workbook, file_path)
        
        # Yield for operations
        yield workbook
        
    except Exception as e:
        logger.error(f"Error opening workbook {file_path}: {e}")
        raise
    finally:
        # Clean up resources
        if handle_id:
            close_resource(handle_id)
        elif workbook:
            try:
                workbook.close()
            except:
                pass

def safe_save_workbook(workbook, file_path: str, retries: int = 3, delay: float = 1.0) -> bool:
    """
    Safely save an Excel workbook with proper error handling and retries.
    
    Args:
        workbook: openpyxl Workbook object
        file_path: Path to save to
        retries: Number of retry attempts
        delay: Delay between retries in seconds
        
    Returns:
        bool: Whether the save was successful
    """
    # Create a backup before saving
    if os.path.exists(file_path):
        backup_path = create_backup_file(file_path)
        logger.info(f"Created backup at: {backup_path}")
    
    # Always save to a temporary file first
    temp_file = get_temp_file_path(prefix="save_temp")
    
    for attempt in range(retries):
        try:
            logger.info(f"Save attempt {attempt + 1}/{retries} to temp file {temp_file}")
            
            # Save to temp file
            workbook.save(temp_file)
            
            # Verify the save succeeded
            if not os.path.exists(temp_file) or os.path.getsize(temp_file) == 0:
                raise IOError(f"Save verification failed for {temp_file}")
            
            # Test that the saved file can be opened
            with safe_open_workbook(temp_file, read_only=True) as test_wb:
                if not test_wb:
                    raise ValueError("Failed to verify saved workbook integrity")
            
            # Move temp file to final destination
            if os.path.exists(file_path):
                unlock_file(file_path)
                try:
                    os.unlink(file_path)
                except PermissionError:
                    logger.warning(f"Permission error removing {file_path}, retrying...")
                    time.sleep(0.5)
                    os.unlink(file_path)
            
            # Move the temp file to the target location
            shutil.move(temp_file, file_path)
            
            logger.info(f"Successfully saved workbook to {file_path}")
            return True
            
        except Exception as e:
            logger.warning(f"Save attempt {attempt + 1}/{retries} failed: {e}")
            if attempt < retries - 1:
                time.sleep(delay)
                # Increase delay for next retry
                delay *= 1.5
                # Force resource cleanup
                pythoncom.CoFreeUnusedLibraries()
                gc.collect()
            else:
                logger.error(f"Failed to save workbook after {retries} attempts: {traceback.format_exc()}")
                return False

def create_backup_file(file_path: str) -> str:
    """
    Create a backup copy of a file.
    
    Args:
        file_path: Path to file to backup
        
    Returns:
        str: Path to backup file
    """
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    backup_dir = os.path.join(os.path.dirname(file_path), "backups")
    os.makedirs(backup_dir, exist_ok=True)
    
    file_name = os.path.basename(file_path)
    base_name, ext = os.path.splitext(file_name)
    backup_path = os.path.join(backup_dir, f"{base_name}_backup_{timestamp}{ext}")
    
    # Use safe copy to ensure proper resource management
    safe_copy_file(file_path, backup_path)
    return backup_path

@contextlib.contextmanager
def excel_com_session(visible: bool = False):
    """
    Context manager for safely using an Excel COM session.
    
    Args:
        visible: Whether Excel should be visible
        
    Yields:
        object: The Excel application COM object
    """
    excel = None
    handle_id = None
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Start Excel with proper error handling
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        
        # Track the Excel instance
        handle_id = track_resource('excel', excel)
        
        # Yield for operations
        yield excel
        
    except Exception as e:
        logger.error(f"Error in Excel COM session: {e}")
        raise
    finally:
        # Clean up resources
        if handle_id:
            close_resource(handle_id)
        elif excel:
            try:
                excel.Quit()
            except:
                pass
            del excel
            
        gc.collect()
        pythoncom.CoUninitialize()

def open_com_workbook(excel, file_path: str, read_only: bool = True):
    """
    Safely open a workbook using Excel COM with proper resource tracking.
    
    Args:
        excel: Excel application COM object
        file_path: Path to Excel file
        read_only: Whether to open in read-only mode
        
    Returns:
        Tuple[object, str]: Workbook COM object and its handle ID
    """
    wb = None
    try:
        # Prepare absolute path - important for Windows
        file_path = os.path.abspath(file_path)
        
        # Make sure file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
            
        # Ensure file isn't locked
        unlock_file(file_path)
        
        # Open the workbook with appropriate options
        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=0,  # Don't update links
            ReadOnly=read_only,
            IgnoreReadOnlyRecommended=True,
            CorruptLoad=2  # xlRepairFile (better corruption handling)
        )
        
        # Track the workbook
        handle_id = track_resource('workbook', wb, file_path)
        return wb, handle_id
        
    except Exception as e:
        logger.error(f"Error opening workbook with COM: {e}")
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
            del wb
        raise

def process_excel_file_safely(source_file: str, dest_file: str, processor_func=None) -> bool:
    """
    Process an Excel file with comprehensive error handling and resource management.
    
    Args:
        source_file: Source Excel file path
        dest_file: Destination Excel file path
        processor_func: Function to process the workbook
        
    Returns:
        bool: Whether processing was successful
    """
    # Track the start time
    start_time = time.time()
    
    # Normalize paths
    source_file = os.path.abspath(source_file)
    dest_file = os.path.abspath(dest_file)
    
    logger.info(f"Starting safe Excel processing: {source_file} → {dest_file}")
    
    # Register cleanup handler if not already done
    setup_resource_tracking()
    
    # Close any Excel instances first
    close_excel_instances()
    
    # First make a safe copy of the file
    temp_copy = get_temp_file_path(prefix="safe_copy")
    if not safe_copy_file(source_file, temp_copy):
        logger.error("Failed to create safe copy of the source file")
        return False
        
    try:
        # Create destination directory if it doesn't exist
        os.makedirs(os.path.dirname(dest_file), exist_ok=True)
        
        # Try processing with openpyxl first
        logger.info("Attempting to process with openpyxl")
        success = False
        
        try:
            with safe_open_workbook(temp_copy, read_only=False) as wb:
                if processor_func:
                    processor_func(wb)
                    
                # Save the processed workbook
                success = safe_save_workbook(wb, dest_file)
                
        except Exception as e:
            logger.warning(f"openpyxl processing failed: {e}")
            success = False
            
        # If openpyxl fails, try with COM
        if not success:
            logger.info("Falling back to COM processing")
            try:
                with excel_com_session() as excel:
                    wb, handle_id = open_com_workbook(excel, temp_copy, read_only=False)
                    
                    try:
                        # Run processing function if provided
                        if processor_func:
                            processor_func(wb)
                            
                        # Save to destination
                        wb.SaveAs(
                            dest_file,
                            FileFormat=51,  # xlOpenXMLWorkbook (.xlsx)
                            CreateBackup=False
                        )
                        
                        success = True
                        logger.info(f"Successfully saved processed file to {dest_file}")
                        
                    finally:
                        # Close workbook
                        close_resource(handle_id)
                        
            except Exception as e:
                logger.error(f"COM processing failed: {e}")
                success = False
                
        # Verify the final file
        if success and os.path.exists(dest_file) and os.path.getsize(dest_file) > 0:
            # Log processing time
            elapsed = time.time() - start_time
            logger.info(f"Excel processing completed successfully in {elapsed:.2f} seconds")
            return True
        else:
            logger.error("Excel processing failed - output file is missing or empty")
            return False
            
    except Exception as e:
        logger.error(f"Error during Excel processing: {traceback.format_exc()}")
        return False
    finally:
        # Clean up all resources
        cleanup_all_resources()

# Initialize module when imported
setup_resource_tracking()