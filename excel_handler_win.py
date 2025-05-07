"""
Windows-specific Excel handling module with file resource management.
This module is designed to fix file access issues on Windows systems.
"""

import os
import sys
import time
import logging
import pythoncom
import win32com.client
import win32api
import win32con
import win32process
import ctypes
import gc
from ctypes import wintypes
from contextlib import contextmanager
from typing import Optional, List, Dict, Any, Tuple

# Import the base file operations module
from file_operations_win import (
    safe_file_operation, 
    unlock_file,
    close_excel_instances,
    get_temp_file_path,
    setup_resource_tracking,
    cleanup_all_resources
)

# Initialize logger
logger = logging.getLogger(__name__)

# Windows API constants for file operations
FILE_SHARE_READ = 0x00000001
FILE_SHARE_WRITE = 0x00000002
FILE_SHARE_DELETE = 0x00000004
OPEN_EXISTING = 3
GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000

# Windows handle types
INVALID_HANDLE_VALUE = -1

# Load Windows DLLs for advanced file operations
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
kernel32.FindFirstFileW.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(wintypes.WIN32_FIND_DATAW)]
kernel32.FindFirstFileW.restype = wintypes.HANDLE

def is_file_locked(filepath: str) -> bool:
    """
    Check if a file is locked by another process on Windows.
    
    Args:
        filepath: Path to the file to check
        
    Returns:
        bool: True if the file is locked, False otherwise
    """
    if not os.path.exists(filepath):
        return False
        
    try:
        # Try to open the file with sharing permissions
        handle = kernel32.CreateFileW(
            str(filepath),
            GENERIC_READ | GENERIC_WRITE, 
            0,  # No sharing
            None,
            OPEN_EXISTING,
            0,
            None
        )
        
        if handle == INVALID_HANDLE_VALUE:
            error = ctypes.get_last_error()
            # 32 = ERROR_SHARING_VIOLATION (file in use)
            # 33 = ERROR_LOCK_VIOLATION (file locked)
            if error in (32, 33):
                return True
        else:
            # Close the handle if we got one
            kernel32.CloseHandle(handle)
            return False
    except Exception as e:
        logger.warning(f"Error checking if file is locked: {e}")
        return True  # Assume locked on error

def force_file_unlock(filepath: str) -> bool:
    """
    Force unlock a file on Windows using system-level mechanisms.
    
    Args:
        filepath: Path to the file to unlock
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not os.path.exists(filepath):
        return True
        
    try:
        # First try normal unlock methods
        if unlock_file(filepath):
            # Check if that worked
            if not is_file_locked(filepath):
                return True
                
        # Try more aggressive methods (Windows only)
        filepath = os.path.abspath(filepath)
        logger.info(f"Attempting aggressive unlock of {filepath}")
        
        # Use Windows system file cache flush
        try:
            sys_path = os.path.normpath(filepath)
            kernel32.SetFileAttributesW(sys_path, win32con.FILE_ATTRIBUTE_NORMAL)
        except:
            pass
            
        # Force flush the system cache
        try:
            import win32file
            win32file.FlushFileBuffers(win32file._get_osfhandle(sys_path))
        except:
            pass
            
        # Release COM objects that might have the file open
        pythoncom.CoFreeUnusedLibraries()
        gc.collect()
        time.sleep(0.5)  # Small delay to let OS process things
        
        # Check again
        return not is_file_locked(filepath)
    except Exception as e:
        logger.warning(f"Error force unlocking file: {e}")
        return False

def reset_excel_automation() -> bool:
    """
    Reset Excel automation on Windows to fix COM issues.
    
    Returns:
        bool: True if successful
    """
    # First terminate all Excel processes
    close_excel_instances()
    
    try:
        # Unload COM libraries
        pythoncom.CoUninitialize()
        time.sleep(0.5)
        pythoncom.CoInitialize()
        
        # Reset Office application cache
        pythoncom.CoFreeUnusedLibraries()
        
        # Force garbage collection
        gc.collect()
        time.sleep(1)
        
        # Additional Windows-specific registry reset
        try:
            # Only import Windows registry modules when needed
            import winreg
            
            # Reset Excel automation keys (safe to try even if keys don't exist)
            keys_to_reset = [
                r'Software\Microsoft\Office\16.0\Excel\Resiliency',
                r'Software\Microsoft\Office\Excel\Addins'
            ]
            
            for key_path in keys_to_reset:
                try:
                    registry_key = winreg.OpenKey(
                        winreg.HKEY_CURRENT_USER,
                        key_path,
                        0, 
                        winreg.KEY_WRITE
                    )
                    winreg.CloseKey(registry_key)
                except:
                    pass
        except:
            pass
        
        return True
    except Exception as e:
        logger.warning(f"Error resetting Excel automation: {e}")
        return True  # Return True anyway since this is just an optimization

@contextmanager
def robust_excel_session(visible: bool = False):
    """
    Context manager for Excel automation with enhanced robustness for Windows.
    
    Args:
        visible: Whether Excel should be visible
        
    Yields:
        object: Excel application COM object
    """
    # Initialize session tracking
    excel = None
    
    try:
        # Reset Excel automation state
        reset_excel_automation()
        
        # Initialize COM with appropriate threading model
        pythoncom.CoInitialize()
        
        # Create Excel application with enhanced error handling
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        
        # Disable specific features that might cause issues
        excel.FeatureInstall = 0  # msoFeatureInstallNone
        
        # Set calculation to manual for performance
        excel.Calculation = -4135  # xlCalculationManual
        
        # Yield the Excel application
        yield excel
        
    except Exception as e:
        logger.error(f"Error in Excel session: {e}")
        raise
    finally:
        # Clean up Excel instance
        if excel:
            try:
                excel.Quit()
            except:
                pass
            
            # Force cleanup of COM object
            del excel
            gc.collect()
            
        # Release COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def fix_excel_file(source_path: str, dest_path: str = None) -> str:
    """
    Fix a potentially corrupted Excel file by using the robust Excel session.
    
    Args:
        source_path: Path to the Excel file to fix
        dest_path: Optional destination path (if None, creates a fixed copy)
        
    Returns:
        str: Path to the fixed file
    """
    if dest_path is None:
        # Generate a path for the fixed file
        dirname, filename = os.path.split(source_path)
        base, ext = os.path.splitext(filename)
        dest_path = os.path.join(dirname, f"{base}_fixed{ext}")
    
    logger.info(f"Attempting to fix Excel file: {source_path} -> {dest_path}")
    
    # Ensure directories exist
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    
    # Create a temp file for intermediate processing
    temp_path = get_temp_file_path(prefix="excel_fix")
    
    try:
        # Force unlock the source file
        force_file_unlock(source_path)
        
        # Use robust Excel session to open and repair
        with robust_excel_session() as excel:
            # Open workbook with recovery options
            wb = excel.Workbooks.Open(
                os.path.abspath(source_path),
                UpdateLinks=0,
                ReadOnly=True,
                CorruptLoad=2  # xlRepairFile
            )
            
            # Save to temp location with clean options
            wb.SaveAs(
                os.path.abspath(temp_path),
                FileFormat=51,  # xlOpenXMLWorkbook
                ConflictResolution=2  # xlLocalSessionChanges
            )
            
            # Close properly
            wb.Close(SaveChanges=False)
            
            # Force cleanup
            del wb
            gc.collect()
        
        # Verify the temp file was created successfully
        if not os.path.exists(temp_path) or os.path.getsize(temp_path) == 0:
            raise IOError(f"Failed to create valid fixed file at {temp_path}")
        
        # If the destination exists, force unlock and remove it
        if os.path.exists(dest_path):
            force_file_unlock(dest_path)
            os.unlink(dest_path)
        
        # Move the temp file to the destination
        os.rename(temp_path, dest_path)
        
        logger.info(f"Successfully fixed Excel file: {dest_path}")
        return dest_path
        
    except Exception as e:
        logger.error(f"Error fixing Excel file: {e}")
        raise
    finally:
        # Clean up temp file if it still exists
        if os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except:
                pass

def check_excel_temp_files() -> bool:
    """
    Check and clean up any Excel temporary files that might be causing issues.
    
    Returns:
        bool: True if cleaning was successful
    """
    try:
        # Clean up Excel temp files (Windows specific)
        temp_dir = os.environ.get('TEMP')
        if not temp_dir:
            return False
            
        # Look for Excel temporary files
        excel_temps = []
        
        # Pattern 1: Excel temporary files
        for pattern in ['~$*.xls*', 'Excel*.tmp']:
            search_path = os.path.join(temp_dir, pattern)
            excel_temps.extend(win32api.FindFiles(search_path))
        
        # Remove found temporary files
        for file_info in excel_temps:
            try:
                file_path = os.path.join(temp_dir, file_info[8])
                force_file_unlock(file_path)
                os.unlink(file_path)
                logger.info(f"Removed Excel temporary file: {file_path}")
            except Exception as e:
                logger.warning(f"Failed to remove temporary file: {e}")
        
        return True
    except Exception as e:
        logger.warning(f"Error cleaning Excel temporary files: {e}")
        return False

def repair_excel_file_access(file_path: str) -> bool:
    """
    Comprehensive solution to fix Windows file access issues with Excel files.
    
    Args:
        file_path: Path to Excel file with access issues
        
    Returns:
        bool: True if repairs were successful
    """
    logger.info(f"Performing comprehensive file access repair for: {file_path}")
    
    # Step 1: Set up resource tracking
    setup_resource_tracking()
    
    # Step 2: Close all Excel instances
    close_excel_instances()
    
    # Step 3: Reset Excel automation
    reset_excel_automation()
    
    # Step 4: Clean up temporary files
    check_excel_temp_files()
    
    # Step 5: Force unlock the file
    unlocked = force_file_unlock(file_path)
    
    if not unlocked:
        logger.warning(f"Could not unlock file: {file_path}")
    
    # Step 6: Check if the file exists and is accessible
    if os.path.exists(file_path):
        try:
            # Test file access with proper error handling
            with open(file_path, 'rb') as f:
                # Just read a small bit to verify access
                f.read(10)
                
            logger.info(f"Successfully repaired file access for: {file_path}")
            return True
        except Exception as e:
            logger.error(f"File still inaccessible after repair attempt: {e}")
            return False
    else:
        logger.error(f"File does not exist: {file_path}")
        return False