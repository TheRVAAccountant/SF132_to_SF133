"""
Windows path handling utility for Excel processing.
This module provides functions to ensure robust file path handling on Windows systems.
"""

import os
import sys
import re
import logging
from pathlib import Path
from typing import Optional, Union, List

# Initialize logger
logger = logging.getLogger(__name__)

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
    if sys.platform.startswith('win'):
        if re.match(r'^[a-zA-Z]:\\', path):
            # Capitalize drive letter for consistency
            path = path[0].upper() + path[1:]
    
    return path

def ensure_unc_path(path: str) -> str:
    """
    Ensure a path is in proper UNC format for Windows network drives.
    
    Args:
        path: Path to convert
        
    Returns:
        str: Path in proper UNC format
    """
    # First normalize the path
    path = normalize_windows_path(path)
    
    # Check if already in UNC format
    if path.startswith('\\\\'):
        return path
        
    # Check if path is a network mapped drive
    if re.match(r'^[a-zA-Z]:\\', path) and sys.platform.startswith('win'):
        try:
            import win32wnet
            try:
                # Try to convert mapped drive to UNC
                unc_path = win32wnet.WNetGetUniversalName(path, 1)
                return unc_path
            except Exception:
                # Not a mapped drive or can't be converted
                pass
        except ImportError:
            pass
    
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
    if not sys.platform.startswith('win'):
        return path
        
    try:
        # Normalize the path first
        path = normalize_windows_path(path)
        
        if not os.path.exists(path):
            return path  # Return original if path doesn't exist
        
        import win32api
        short_path = win32api.GetShortPathName(path)
        return short_path
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
    if not sys.platform.startswith('win'):
        return path
        
    try:
        # Normalize the path first
        path = normalize_windows_path(path)
        
        import win32api
        long_path = win32api.GetLongPathName(path)
        return long_path
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
    if not sys.platform.startswith('win'):
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

def get_resolved_path(path: str) -> str:
    """
    Get a fully resolved path that works reliably on Windows.
    Handles path normalization, short/long path conversion,
    and long path prefixing.
    
    Args:
        path: File path to resolve
        
    Returns:
        str: Fully resolved path
    """
    # Normalize the path first
    path = normalize_windows_path(path)
    
    # If path doesn't exist, just return normalized path
    if not os.path.exists(path):
        return path
    
    # Handle Windows-specific conversions
    if sys.platform.startswith('win'):
        # 1. Get absolute path
        path = os.path.abspath(path)
        
        # 2. Try to get the canonical path to resolve any 8.3 names
        try:
            path = get_long_path_name(path)
        except:
            pass
            
        # 3. Add long path prefix if needed
        if len(path) >= 260:
            path = handle_long_paths(path)
    else:
        # On non-Windows, just use absolute path
        path = os.path.abspath(path)
    
    return path

def build_windows_compatible_path(*parts: str) -> str:
    """
    Build a Windows-compatible path from parts.
    
    Args:
        *parts: Path parts to join
        
    Returns:
        str: Windows-compatible path
    """
    # Join parts using os.path.join
    path = os.path.join(*parts)
    
    # Normalize the joined path
    path = normalize_windows_path(path)
    
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

def is_valid_windows_path(path: str) -> bool:
    """
    Check if a path is valid on Windows.
    
    Args:
        path: Path to check
        
    Returns:
        bool: True if path is valid
    """
    # Check for invalid characters on Windows
    invalid_chars = r'[<>:"|?*]'
    
    # Check if path has invalid characters
    if re.search(invalid_chars, path):
        return False
    
    # Check for reserved names
    reserved_names = [
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    ]
    
    # Split the path and check each part
    parts = path.split('\\')
    for part in parts:
        if not part:
            continue
            
        # Check for reserved names (case-insensitive)
        name_part = part.split('.')[0].upper()
        if name_part in reserved_names:
            return False
    
    return True

def convert_path_for_excel(path: str) -> str:
    """
    Convert a path for reliable use with Excel.
    
    Args:
        path: Path to convert
        
    Returns:
        str: Excel-compatible path
    """
    # Normalize the path first
    path = normalize_windows_path(path)
    
    # For UNC paths, ensure proper format
    if path.startswith('\\\\'):
        return path
    
    # For local paths, ensure drive letter is capitalized
    if re.match(r'^[a-zA-Z]:\\', path):
        path = path[0].upper() + path[1:]
    
    # If running on Windows, get a short path for paths with spaces
    if sys.platform.startswith('win') and ' ' in path:
        try:
            path = get_short_path_name(path)
        except:
            pass
    
    return path

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
    
    # 2. Handle long paths if needed
    if len(path) >= 260:
        path = handle_long_paths(path)
    
    # 3. Convert to short path if it contains problematic characters
    if sys.platform.startswith('win') and (
        ' ' in path or 
        any(c in path for c in '()[]{}!@#$%^&=;,\'`~')
    ):
        try:
            if os.path.exists(path):
                path = get_short_path_name(path)
        except:
            pass
    
    # 4. For Excel compatibility, format UNC paths properly
    if re.match(r'^[a-zA-Z]:\\', path):
        path = path[0].upper() + path[1:]
    
    return path