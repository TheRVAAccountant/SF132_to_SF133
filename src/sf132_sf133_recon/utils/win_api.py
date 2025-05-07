"""
Windows-specific low-level API functionality.

This module provides direct access to Windows APIs through ctypes,
primarily for file handling and process management.
"""

import os
import sys
import logging
import ctypes
from typing import Optional, List, Dict, Any, Tuple, Union, Set
from ctypes import wintypes
import time

# Initialize logger
logger = logging.getLogger(__name__)

# Only load these on Windows
if sys.platform.startswith('win'):
    try:
        # Windows-specific imports
        import win32api
        import win32con
        import winreg
        
        # Define Windows constants
        INVALID_HANDLE_VALUE = -1
        FILE_SHARE_READ = 0x00000001
        FILE_SHARE_WRITE = 0x00000002
        FILE_SHARE_DELETE = 0x00000004
        OPEN_EXISTING = 3
        FILE_ATTRIBUTE_NORMAL = 0x80
        GENERIC_READ = 0x80000000
        FILE_FLAG_BACKUP_SEMANTICS = 0x02000000
        
        # Get kernel32.dll
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        
        # Define required functions
        kernel32.CreateFileW.argtypes = [
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.LPVOID,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.HANDLE
        ]
        kernel32.CreateFileW.restype = wintypes.HANDLE
        
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel32.CloseHandle.restype = wintypes.BOOL
        
        # Track resources for cleanup
        _tracked_resources = set()
        
    except ImportError as e:
        if sys.platform.startswith('win'):
            # Critical on Windows
            logger.critical(f"Failed to load Windows API modules: {e}")
            print(f"CRITICAL ERROR: Windows API modules missing: {e}")
            print("Please install required dependencies: pip install pywin32")
            raise ImportError(f"Windows API modules required: {e}") from e
        else:
            # Just log on non-Windows
            logger.warning(f"Windows API modules not available on this platform: {e}")
            
else:
    # Not on Windows, so log a warning
    logger.warning("Not running on Windows - Windows API features unavailable")

def is_windows() -> bool:
    """
    Check if running on Windows.
    
    Returns:
        bool: True if on Windows
    """
    return sys.platform.startswith('win')

def setup_resource_tracking():
    """Initialize resource tracking for Windows handles."""
    global _tracked_resources
    _tracked_resources = set()

def track_resource(resource_handle):
    """
    Track a Windows resource handle for later cleanup.
    
    Args:
        resource_handle: Windows resource handle
    """
    if resource_handle and resource_handle != INVALID_HANDLE_VALUE:
        _tracked_resources.add(resource_handle)

def close_resource(resource_handle):
    """
    Close a tracked Windows resource handle.
    
    Args:
        resource_handle: Windows resource handle
    
    Returns:
        bool: True if successful
    """
    if resource_handle and resource_handle != INVALID_HANDLE_VALUE:
        try:
            if resource_handle in _tracked_resources:
                _tracked_resources.remove(resource_handle)
            success = kernel32.CloseHandle(resource_handle)
            return success
        except Exception as e:
            logger.error(f"Error closing handle {resource_handle}: {e}")
    return False

def cleanup_resources():
    """Close all tracked Windows resource handles."""
    global _tracked_resources
    for handle in list(_tracked_resources):
        close_resource(handle)
    _tracked_resources.clear()

def is_file_locked_win(file_path: str) -> bool:
    """
    Check if a file is locked using Windows API.
    
    Args:
        file_path: Path to the file
        
    Returns:
        bool: True if the file is locked
    """
    if not is_windows():
        return False
        
    if not os.path.exists(file_path):
        return False
        
    # Convert to Unicode for Windows API
    file_path_w = os.path.abspath(file_path)
    
    # Try to open the file with minimal sharing
    handle = kernel32.CreateFileW(
        file_path_w,
        GENERIC_READ,
        FILE_SHARE_READ,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None
    )
    
    if handle == INVALID_HANDLE_VALUE:
        error = ctypes.get_last_error()
        # Error code 32 = ERROR_SHARING_VIOLATION (file in use)
        # Error code 33 = ERROR_LOCK_VIOLATION (file locked)
        return error in (32, 33)
    else:
        # We could open it, so it's not locked
        kernel32.CloseHandle(handle)
        return False

def force_file_unlock(file_path: str) -> bool:
    """
    Attempt to forcefully unlock a file using multiple methods.
    
    Args:
        file_path: Path to the file
        
    Returns:
        bool: True if unlocking was successful
    """
    if not is_windows():
        logger.warning("force_file_unlock only works on Windows")
        return False
        
    if not os.path.exists(file_path):
        return True  # Nothing to unlock
        
    logger.info(f"Attempting to force unlock file: {file_path}")
    
    # Method 1: Try to close any Excel processes
    from ..modules.excel_handler import close_excel_instances
    close_excel_instances()
    
    # Wait to see if that helped
    time.sleep(1)
    
    # Check if file is now unlocked
    if not is_file_locked_win(file_path):
        logger.info("File unlocked successfully by closing Excel instances")
        return True
    
    # Method 2: Try to open with FILE_SHARE_DELETE
    try:
        # Try opening with more permissive flags
        file_path_w = os.path.abspath(file_path)
        handle = kernel32.CreateFileW(
            file_path_w,
            GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL | FILE_FLAG_BACKUP_SEMANTICS,
            None
        )
        
        if handle != INVALID_HANDLE_VALUE:
            kernel32.CloseHandle(handle)
            logger.info("File unlocked successfully using permissive file flags")
            return True
    except Exception as e:
        logger.warning(f"Error trying to unlock file with permissive flags: {e}")
    
    # Method 3: Try to use a temporary file with move-after-reboot
    try:
        temp_path = file_path + ".unlockreboot"
        
        # Try to use MoveFileEx with MOVEFILE_DELAY_UNTIL_REBOOT
        # This requires admin privileges, so it might fail
        try:
            win32api.MoveFileEx(file_path, None, win32con.MOVEFILE_DELAY_UNTIL_REBOOT)
            logger.info("File scheduled for deletion on reboot")
        except Exception as e:
            logger.warning(f"Failed to schedule file for deletion on reboot: {e}")
    
    except Exception as e:
        logger.warning(f"Error trying to schedule file unlock: {e}")
    
    logger.warning(f"Could not unlock file: {file_path}")
    return False

def reset_excel_automation() -> bool:
    """
    Reset Excel COM automation registry settings.
    
    Returns:
        bool: True if successful
    """
    if not is_windows():
        logger.warning("reset_excel_automation only works on Windows")
        return False
    
    try:
        # Reset Excel Automation registry settings
        reg_paths = [
            r'SOFTWARE\Microsoft\Office\Excel\Addins',
            r'SOFTWARE\Classes\Excel.Application',
            r'SOFTWARE\Classes\Excel.Sheet.12'
        ]
        
        for reg_path in reg_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, 
                                    winreg.KEY_WRITE) as key:
                    pass  # Just attempt to open for writing to test access
            except Exception:
                logger.debug(f"Could not access registry key: {reg_path}")
        
        logger.info("Reset Excel automation settings")
        return True
    except Exception as e:
        logger.error(f"Error resetting Excel automation: {e}")
        return False

def check_excel_temp_files() -> List[str]:
    """
    Check for Excel temporary files that might be causing lock issues.
    
    Returns:
        List[str]: List of Excel temporary files
    """
    if not is_windows():
        return []
        
    temp_files = []
    try:
        # Check common Excel temp file locations
        temp_dir = os.environ.get('TEMP', '')
        if temp_dir and os.path.exists(temp_dir):
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    if file.startswith('~$') or file.endswith('.tmp'):
                        temp_files.append(os.path.join(root, file))
    except Exception as e:
        logger.error(f"Error checking Excel temp files: {e}")
    
    return temp_files

def cleanup_excel_temp_files() -> int:
    """
    Clean up Excel temporary files.
    
    Returns:
        int: Number of files cleaned up
    """
    if not is_windows():
        return 0
        
    temp_files = check_excel_temp_files()
    cleaned = 0
    
    for file in temp_files:
        try:
            if not is_file_locked_win(file):
                os.unlink(file)
                cleaned += 1
        except Exception:
            pass
    
    if cleaned > 0:
        logger.info(f"Cleaned up {cleaned} Excel temporary files")
        
    return cleaned