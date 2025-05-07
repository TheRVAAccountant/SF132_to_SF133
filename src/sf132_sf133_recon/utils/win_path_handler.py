"""
Windows path handling utilities.

This module provides functions for handling Windows-specific path issues,
particularly for Excel file operations.
"""

import os
import sys
import re
import logging
from pathlib import Path
from typing import Optional, Union, List

# Initialize logger
logger = logging.getLogger(__name__)

def is_windows() -> bool:
    """
    Check if running on Windows platform.
    
    Returns:
        bool: True if running on Windows
    """
    return sys.platform.startswith('win')

def normalize_windows_path(path: str) -> str:
    """
    Normalize a Windows path for proper handling across platforms.
    
    Args:
        path: The file path to normalize
        
    Returns:
        str: Normalized path
    """
    if not path:
        return path
        
    # Convert to raw string to handle escape sequences properly
    path = str(path)
    
    # Handle forward slashes first
    path = path.replace('/', '\\')
    
    # Handle double backslashes (might be from string escaping)
    path = path.replace('\\\\', '\\')
    
    # Normalize path using os.path functions
    path = os.path.normpath(path)
    
    # If running on Windows, ensure proper drive letter casing
    if is_windows():
        if re.match(r'^[a-zA-Z]:\\', path):
            # Capitalize drive letter for consistency
            path = path[0].upper() + path[1:]
    
    return path

def get_short_path_name(path: str) -> str:
    """
    Get the Windows short path (8.3 format) for a long path.
    Useful for handling paths with spaces or special characters.
    
    Args:
        path: Long file path
        
    Returns:
        str: Short file path
    """
    # Only relevant on Windows
    if not is_windows():
        return path
        
    try:
        # Normalize the path first
        path = normalize_windows_path(path)
        
        if not os.path.exists(path):
            return path  # Return original if path doesn't exist
        
        try:
            import win32api
            short_path = win32api.GetShortPathName(path)
            return short_path
        except ImportError:
            logger.warning("win32api module not available, cannot get short path name")
            return path
            
    except Exception as e:
        logger.warning(f"Error getting short path name: {e}")
        return path  # Return original on error

def get_long_path_name(path: str) -> str:
    """
    Get the long path name for a Windows short path.
    
    Args:
        path: Short file path (8.3 format)
        
    Returns:
        str: Long file path
    """
    # Only relevant on Windows
    if not is_windows():
        return path
        
    try:
        # Normalize the path first
        path = normalize_windows_path(path)
        
        try:
            import win32api
            long_path = win32api.GetLongPathName(path)
            return long_path
        except ImportError:
            logger.warning("win32api module not available, cannot get long path name")
            return path
            
    except Exception as e:
        logger.warning(f"Error getting long path name: {e}")
        return path  # Return original on error

def handle_long_paths(path: str) -> str:
    """
    Handle Windows long paths (>260 characters) using the \\?\ prefix.
    
    Args:
        path: File path that might exceed Windows path limits
        
    Returns:
        str: Path with appropriate prefix for long path support
    """
    # Only relevant on Windows
    if not is_windows():
        return path
        
    # Normalize the path first
    path = normalize_windows_path(path)
    
    # Add the long path prefix if needed
    if len(path) >= 260 and not path.startswith('\\\\?\\'):
        # Get absolute path
        path = os.path.abspath(path)
        
        # Add the prefix
        path = '\\\\?\\' + path
    
    return path

def ensure_unc_path(path: str) -> str:
    """
    Ensure network paths are properly formatted as UNC paths on Windows.
    
    Args:
        path: File path that might be a network path
        
    Returns:
        str: Properly formatted UNC path
    """
    # Only relevant on Windows
    if not is_windows():
        return path
        
    # Return if not a network path
    if not path.startswith(r'\\'):
        return path
        
    # Normalize path separators
    path = path.replace('/', '\\')
    
    # Check if this is already a properly formatted UNC path
    if path.startswith(r'\\?\UNC\'):
        return path
        
    # Format as UNC path for long path support
    if path.startswith(r'\\'):
        # Remove the initial \\ and add \\?\UNC\
        path = r'\\?\UNC\' + path[2:]
    
    return path

def get_windows_shortpath(path: str) -> str:
    """
    Get the Windows short path (8.3 format) for a long path name.
    
    Args:
        path: Long file path
        
    Returns:
        str: Windows short path
    """
    # Only relevant on Windows
    if not is_windows():
        return path
        
    if not os.path.exists(path):
        return path
        
    try:
        import win32api
        return win32api.GetShortPathName(path)
    except ImportError:
        logger.warning("win32api not available, cannot get short path")
        return path
    except Exception as e:
        logger.warning(f"Error getting short path: {e}")
        return path

def get_windows_special_folder(folder_name: str) -> str:
    """
    Get the path to a Windows special folder.
    
    Args:
        folder_name: Name of the special folder (e.g., 'TEMP', 'DESKTOP')
        
    Returns:
        str: Path to the special folder
    """
    # Only relevant on Windows
    if not is_windows():
        return ""
        
    try:
        import win32com.client
        import win32con
        
        # Map of folder names to CSIDL constants
        folder_map = {
            'TEMP': 0x2A,  # CSIDL_LOCAL_APPDATA\Temp
            'APPDATA': 0x1A,  # CSIDL_APPDATA
            'LOCALAPPDATA': 0x1C,  # CSIDL_LOCAL_APPDATA
            'DESKTOP': 0x10,  # CSIDL_DESKTOPDIRECTORY
            'DOCUMENTS': 0x05,  # CSIDL_PERSONAL
            'PROFILE': 0x28,  # CSIDL_PROFILE
        }
        
        if folder_name.upper() == 'TEMP':
            # For TEMP, get %TEMP% directly
            return os.environ.get('TEMP', '')
        
        if folder_name.upper() in folder_map:
            shell = win32com.client.Dispatch("WScript.Shell")
            folder_id = folder_map[folder_name.upper()]
            folder = shell.SpecialFolders(folder_id)
            return folder
    except ImportError:
        logger.warning("win32com not available, cannot get special folder")
    except Exception as e:
        logger.warning(f"Error getting special folder: {e}")
    
    return ""

def fix_excel_path(path: str) -> str:
    """
    Fix a path for use with Excel, handling all Windows-specific issues.
    
    Args:
        path: Path to fix
        
    Returns:
        str: Fixed path
    """
    # 1. Normalize the path
    path = normalize_windows_path(path)
    
    # Windows-specific handling
    if is_windows():
        # 2. Handle long paths if needed
        if len(path) >= 260:
            path = handle_long_paths(path)
        
        # 3. Convert to short path if it contains problematic characters
        if ' ' in path or any(c in path for c in '()[]{}!@#$%^&=;,\'`~'):
            try:
                if os.path.exists(path):
                    path = get_short_path_name(path)
            except:
                pass
        
        # 4. For Excel compatibility, format drive letters properly
        if re.match(r'^[a-zA-Z]:\\', path):
            path = path[0].upper() + path[1:]
    
    return path

def ensure_directory_exists(path: str) -> bool:
    """
    Ensure a directory exists, creating it if necessary.
    Handles path normalization for Windows.
    
    Args:
        path: Directory path
        
    Returns:
        bool: True if directory exists or was created
    """
    # Normalize the path first
    path = normalize_windows_path(path)
    
    try:
        # Create directory with parents
        os.makedirs(path, exist_ok=True)
        return True
    except Exception as e:
        logger.error(f"Error creating directory {path}: {e}")
        return False