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