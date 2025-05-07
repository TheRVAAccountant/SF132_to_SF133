"""
Excel file recovery utility for handling file access errors on Windows.
This module provides robust error recovery mechanisms when files are locked or corrupted.
"""

import os
import sys
import time
import logging
import tempfile
import shutil
import gc
import traceback
import contextlib
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

# Import required modules
try:
    import win32api
    import win32con
    import win32file
    import pywintypes
    WINDOWS_MODULES_AVAILABLE = True
except ImportError:
    WINDOWS_MODULES_AVAILABLE = False

# Import other internal modules
from win_path_handler import normalize_windows_path, get_short_path_name, fix_excel_path
from excel_handler_win import force_file_unlock, robust_excel_session, repair_excel_file_access
from file_operations_win import (
    safe_file_operation, 
    get_temp_file_path,
    safe_copy_file,
    close_excel_instances
)

# Initialize logger
logger = logging.getLogger(__name__)

class ExcelFileRecovery:
    """
    Provides comprehensive Excel file recovery mechanisms for Windows.
    """
    
    def __init__(self, max_recovery_attempts: int = 3):
        """
        Initialize the recovery handler.
        
        Args:
            max_recovery_attempts: Maximum number of recovery attempts
        """
        self.max_recovery_attempts = max_recovery_attempts
        self.temp_files = []
        
    def __del__(self):
        """Clean up any temporary files when the instance is destroyed."""
        self.cleanup_temp_files()
        
    def cleanup_temp_files(self):
        """Clean up all temporary files."""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception as e:
                logger.debug(f"Failed to clean up temp file {temp_file}: {e}")
    
    def recover_excel_file(self, file_path: str) -> Tuple[bool, str]:
        """
        Attempt to recover an Excel file from access errors.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Tuple[bool, str]: Success status and path to recovered file (or error message)
        """
        file_path = normalize_windows_path(file_path)
        logger.info(f"Attempting to recover Excel file: {file_path}")
        
        # First check if the file exists
        if not os.path.exists(file_path):
            return False, f"File does not exist: {file_path}"
        
        # Create a recovery directory
        recovery_dir = os.path.join(os.path.dirname(file_path), "recovery")
        os.makedirs(recovery_dir, exist_ok=True)
        
        # Create a recovery file path
        base_name = os.path.basename(file_path)
        name, ext = os.path.splitext(base_name)
        recovery_path = os.path.join(recovery_dir, f"{name}_recovered{ext}")
        
        # Try using multiple recovery methods
        for attempt in range(self.max_recovery_attempts):
            try:
                logger.info(f"Recovery attempt {attempt+1}/{self.max_recovery_attempts}")
                
                # Method 1: Try to create a copy with file system operations
                if attempt == 0:
                    if self._recover_with_file_copy(file_path, recovery_path):
                        return True, recovery_path
                
                # Method 2: Try using COM automation to open and save
                elif attempt == 1:
                    if self._recover_with_excel_com(file_path, recovery_path):
                        return True, recovery_path
                
                # Method 3: Try using more aggressive recovery techniques
                else:
                    if self._recover_with_aggressive_techniques(file_path, recovery_path):
                        return True, recovery_path
                
            except Exception as e:
                logger.warning(f"Recovery attempt {attempt+1} failed: {e}")
                
                # Wait before next attempt
                time.sleep(2)
        
        return False, "All recovery attempts failed"
    
    def _recover_with_file_copy(self, source_path: str, dest_path: str) -> bool:
        """
        Recover by creating a copy of the file with proper error handling.
        
        Args:
            source_path: Source file path
            dest_path: Destination file path
            
        Returns:
            bool: Whether recovery was successful
        """
        logger.info("Attempting recovery with file copy")
        
        # First try to unlock the file
        force_file_unlock(source_path)
        
        # Try to create a copy with robust error handling
        if safe_copy_file(source_path, dest_path, retries=3):
            # Verify the copy can be opened
            if self._validate_excel_file(dest_path):
                logger.info(f"Successfully recovered file to {dest_path}")
                return True
            else:
                logger.warning(f"File copy succeeded but validation failed: {dest_path}")
                return False
        
        logger.warning("File copy recovery failed")
        return False
    
    def _recover_with_excel_com(self, source_path: str, dest_path: str) -> bool:
        """
        Recover by using Excel COM automation to open and save the file.
        
        Args:
            source_path: Source file path
            dest_path: Destination file path
            
        Returns:
            bool: Whether recovery was successful
        """
        logger.info("Attempting recovery with Excel COM automation")
        
        # Close all Excel instances first
        close_excel_instances()
        
        # Create a temporary path for intermediate processing
        temp_path = get_temp_file_path(prefix="recovery")
        self.temp_files.append(temp_path)
        
        try:
            # Try a robust Excel session to repair and save
            with robust_excel_session() as excel:
                try:
                    # Open with repair option
                    wb = excel.Workbooks.Open(
                        source_path,
                        UpdateLinks=0,
                        ReadOnly=True,
                        CorruptLoad=2  # xlRepairFile
                    )
                    
                    # Save to temporary location
                    wb.SaveAs(
                        temp_path,
                        FileFormat=51,  # xlOpenXMLWorkbook
                        CreateBackup=False
                    )
                    
                    # Close without saving changes
                    wb.Close(SaveChanges=False)
                    
                    # Force cleanup
                    del wb
                    gc.collect()
                    
                    # If temp file was created successfully, move it to destination
                    if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
                        # Remove destination if it exists
                        if os.path.exists(dest_path):
                            os.unlink(dest_path)
                            
                        # Move temp file to destination
                        shutil.move(temp_path, dest_path)
                        
                        # Verify the file can be opened
                        if self._validate_excel_file(dest_path):
                            logger.info(f"Successfully recovered file with Excel COM to {dest_path}")
                            return True
                    
                    return False
                    
                except Exception as e:
                    logger.warning(f"Excel COM recovery failed: {e}")
                    return False
                    
        except Exception as e:
            logger.warning(f"Excel COM session failed: {e}")
            return False
    
    def _recover_with_aggressive_techniques(self, source_path: str, dest_path: str) -> bool:
        """
        Recover using aggressive techniques for badly corrupted files.
        
        Args:
            source_path: Source file path
            dest_path: Destination file path
            
        Returns:
            bool: Whether recovery was successful
        """
        logger.info("Attempting recovery with aggressive techniques")
        
        try:
            # First ensure the Windows file system isn't holding locks
            repair_excel_file_access(source_path)
            
            # Try to extract data with pandas as a last resort
            import pandas as pd
            
            try:
                # Read all sheets with pandas
                excel_file = pd.ExcelFile(source_path)
                sheet_names = excel_file.sheet_names
                
                # Create a new Excel writer
                with pd.ExcelWriter(dest_path, engine='openpyxl') as writer:
                    # Copy each sheet's data
                    for sheet_name in sheet_names:
                        try:
                            df = pd.read_excel(source_path, sheet_name=sheet_name)
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                        except Exception as sheet_e:
                            logger.warning(f"Could not recover sheet {sheet_name}: {sheet_e}")
                
                # Verify the file was created successfully
                if os.path.exists(dest_path) and os.path.getsize(dest_path) > 0:
                    if self._validate_excel_file(dest_path):
                        logger.info(f"Successfully recovered file data to {dest_path}")
                        return True
            
            except Exception as e:
                logger.warning(f"Pandas recovery failed: {e}")
            
            # If pandas method failed, try the Windows-specific method if available
            if WINDOWS_MODULES_AVAILABLE:
                try:
                    # Create a binary copy of the file using low-level Windows APIs
                    src_handle = win32file.CreateFile(
                        source_path,
                        win32con.GENERIC_READ,
                        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                        None,
                        win32con.OPEN_EXISTING,
                        0,
                        None
                    )
                    
                    # Create the destination file
                    dest_handle = win32file.CreateFile(
                        dest_path,
                        win32con.GENERIC_WRITE,
                        0,  # No sharing
                        None,
                        win32con.CREATE_ALWAYS,
                        0,
                        None
                    )
                    
                    # Read and write in chunks
                    while True:
                        error, data = win32file.ReadFile(src_handle, 65536)
                        if not data:
                            break
                        win32file.WriteFile(dest_handle, data)
                    
                    # Close handles
                    win32file.CloseHandle(src_handle)
                    win32file.CloseHandle(dest_handle)
                    
                    # Verify the file was created successfully
                    if os.path.exists(dest_path) and os.path.getsize(dest_path) > 0:
                        if self._validate_excel_file(dest_path):
                            logger.info(f"Successfully recovered file with Windows API to {dest_path}")
                            return True
                    
                except Exception as e:
                    logger.warning(f"Windows API recovery failed: {e}")
            
            return False
            
        except Exception as e:
            logger.warning(f"Aggressive recovery failed: {e}")
            return False
    
    def _validate_excel_file(self, file_path: str) -> bool:
        """
        Validate that an Excel file can be opened.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            bool: Whether the file is valid
        """
        try:
            # Try to open with openpyxl
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True)
            wb.close()
            return True
        except Exception as e:
            logger.warning(f"Excel validation failed: {e}")
            return False

def fix_excel_file_in_use_error(file_path: str, output_dir: str = None) -> Tuple[bool, str]:
    """
    Fix the "file in use" error specifically for Excel files on Windows.
    
    Args:
        file_path: Path to the Excel file
        output_dir: Optional output directory
        
    Returns:
        Tuple[bool, str]: Success status and path to fixed file (or error message)
    """
    # Normalize path
    file_path = normalize_windows_path(file_path)
    
    # Determine output path
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.basename(file_path)
        output_path = os.path.join(output_dir, base_name)
    else:
        dirname = os.path.dirname(file_path)
        base_name = os.path.basename(file_path)
        name, ext = os.path.splitext(base_name)
        output_path = os.path.join(dirname, f"{name}_fixed{ext}")
    
    logger.info(f"Attempting to fix file-in-use error for {file_path}")
    
    # Step 1: Close any Excel instances
    close_excel_instances()
    
    # Step 2: Force unlock the file
    force_file_unlock(file_path)
    
    # Step 3: Try to create a clean copy
    if safe_copy_file(file_path, output_path):
        logger.info(f"Successfully created clean copy at {output_path}")
        return True, output_path
    
    # Step 4: If copy failed, try recovery methods
    recovery = ExcelFileRecovery()
    success, recovery_path = recovery.recover_excel_file(file_path)
    
    if success:
        # Move recovered file to output path
        try:
            if os.path.exists(output_path):
                os.unlink(output_path)
            shutil.move(recovery_path, output_path)
            return True, output_path
        except Exception as e:
            logger.error(f"Error moving recovered file: {e}")
            return success, recovery_path
    
    return False, "Failed to fix file-in-use error"

# Utility function for use in the excel_processor.py
def process_with_recovery(process_func, file_path: str, *args, **kwargs) -> Tuple[bool, Any]:
    """
    Execute a processing function with comprehensive error recovery.
    
    Args:
        process_func: Function to process the file
        file_path: Path to the Excel file
        *args: Additional arguments for process_func
        **kwargs: Additional keyword arguments for process_func
        
    Returns:
        Tuple[bool, Any]: Success status and return value from process_func (or error message)
    """
    # Normalize path
    file_path = normalize_windows_path(file_path)
    
    try:
        # Try normal processing first
        result = process_func(file_path, *args, **kwargs)
        return True, result
    except Exception as original_error:
        logger.warning(f"Original processing failed: {original_error}")
        
        # Check if this is a file access error
        if "in use" in str(original_error).lower() or "being used" in str(original_error).lower() or \
           "cannot access" in str(original_error).lower() or "permission denied" in str(original_error).lower():
            
            logger.info("Detected file access error, attempting recovery...")
            
            # Try to fix the file-in-use error
            success, fixed_path = fix_excel_file_in_use_error(file_path)
            
            if success:
                try:
                    # Try processing with the fixed file
                    result = process_func(fixed_path, *args, **kwargs)
                    return True, result
                except Exception as recovery_error:
                    logger.error(f"Recovery processing failed: {recovery_error}")
            
        # If we got here, recovery failed
        logger.error(f"Processing failed and recovery was unsuccessful: {original_error}")
        return False, str(original_error)

# Module initialization
def handle_excel_process_error(error_msg: str, file_path: str) -> Tuple[bool, str]:
    """
    Handle common Excel processing errors with automatic recovery.
    
    Args:
        error_msg: The error message
        file_path: Path to the Excel file
        
    Returns:
        Tuple[bool, str]: Success status and recovery path or error message
    """
    # Normalize file path
    file_path = normalize_windows_path(file_path)
    
    # Check for known error patterns
    if "process cannot access the file" in error_msg.lower():
        # Classic file in use error
        logger.info("Detected 'process cannot access the file' error")
        return fix_excel_file_in_use_error(file_path)
        
    elif "being used by another process" in error_msg.lower():
        # Another variant of file in use
        logger.info("Detected 'being used by another process' error")
        return fix_excel_file_in_use_error(file_path)
        
    elif "failed to terminate excel process" in error_msg.lower():
        # Excel process termination issue
        logger.info("Detected Excel process termination issue")
        
        # More aggressive Excel process cleanup
        close_excel_instances()
        time.sleep(1)  # Give OS time to release resources
        
        return fix_excel_file_in_use_error(file_path)
        
    elif "com copy failed" in error_msg.lower() or "com validation failed" in error_msg.lower():
        # COM automation issues
        logger.info("Detected COM automation issue")
        
        # Recovery with Excel file recovery utility
        recovery = ExcelFileRecovery()
        return recovery.recover_excel_file(file_path)
        
    # Generic error handling for other cases
    logger.info(f"Handling generic Excel error: {error_msg}")
    
    # Try repair with file access utility
    if repair_excel_file_access(file_path):
        # If repair worked, return the original path
        return True, file_path
        
    # If all else fails, try full recovery
    recovery = ExcelFileRecovery()
    return recovery.recover_excel_file(file_path)