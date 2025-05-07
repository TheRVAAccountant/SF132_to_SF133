"""
Windows integration module for SF132 to SF133 reconciliation tool.

This module connects all Windows-specific features and ensures they work together
properly. It's automatically imported when running on Windows.
"""

import sys
import os
import logging
import functools
import inspect
from typing import Any, Callable, Optional, Union, Dict, List, Tuple, TypeVar

# Set up logger
logger = logging.getLogger(__name__)

# Type variables for function decorators
F = TypeVar('F', bound=Callable[..., Any])

# Only perform Windows-specific handling on Windows
IS_WINDOWS = sys.platform.startswith('win')

if IS_WINDOWS:
    try:
        # Import all Windows-specific modules
        import win32api
        import win32com.client
        import pythoncom
        import winreg

        # Import our Windows-specific modules
        from ..utils.win_api import (
            reset_excel_automation,
            cleanup_excel_temp_files,
            cleanup_resources,
            is_file_locked_win,
            force_file_unlock
        )
        
        from ..utils.win_path_handler import (
            normalize_windows_path,
            handle_long_paths,
            ensure_unc_path,
            get_windows_shortpath
        )
        
        from .excel_session import (
            robust_excel_session,
            repair_excel_workbook
        )
        
        from .excel_recovery import (
            repair_excel_file_access
        )
        
        WINDOWS_MODULES_LOADED = True
        logger.info("All Windows-specific modules loaded successfully")
        
    except ImportError as e:
        WINDOWS_MODULES_LOADED = False
        logger.critical(f"Failed to load Windows modules: {e}")
        print(f"CRITICAL ERROR: This application requires Windows-specific modules")
        print(f"Please install all required dependencies: {e}")
        raise ImportError(f"Windows modules required: {e}") from e
else:
    WINDOWS_MODULES_LOADED = False
    logger.warning("Not running on Windows - Windows integration not available")

def with_windows_path_handling(func: F) -> F:
    """
    Decorator to add Windows path handling to functions that deal with file paths.
    
    Args:
        func: Function to wrap
        
    Returns:
        Function with Windows path handling
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        if not IS_WINDOWS:
            return func(*args, **kwargs)
        
        # Check if the function has file path arguments
        sig = inspect.signature(func)
        param_names = [p.name for p in sig.parameters.values()]
        
        # Process positional arguments that might be file paths
        path_param_positions = []
        for param in param_names:
            if 'path' in param.lower() or 'file' in param.lower() or 'dir' in param.lower():
                position = param_names.index(param)
                path_param_positions.append(position)
        
        # Normalize path arguments
        new_args = list(args)
        for pos in path_param_positions:
            if pos < len(args) and isinstance(args[pos], str):
                new_args[pos] = normalize_windows_path(args[pos])
        
        # Process keyword arguments that might be file paths
        new_kwargs = dict(kwargs)
        for key, value in kwargs.items():
            if ('path' in key.lower() or 'file' in key.lower() or 'dir' in key.lower()) and isinstance(value, str):
                new_kwargs[key] = normalize_windows_path(value)
        
        # Call the original function with processed arguments
        return func(*new_args, **new_kwargs)
    
    return wrapper  # type: ignore

def with_excel_recovery(func: F) -> F:
    """
    Decorator to add Excel file recovery to functions that process Excel files.
    
    Args:
        func: Function to wrap
        
    Returns:
        Function with Excel recovery capability
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        if not IS_WINDOWS:
            return func(*args, **kwargs)
        
        try:
            return func(*args, **kwargs)
        except Exception as e:
            error_message = str(e).lower()
            
            # Check if this is a file access error
            is_access_error = (
                'access' in error_message or
                'permission' in error_message or
                'in use' in error_message or
                'sharing violation' in error_message or
                'excel' in error_message and 'failed' in error_message
            )
            
            if not is_access_error:
                # Not an access error, re-raise
                raise
            
            # Try to find the Excel file path in the arguments
            excel_path = None
            
            # Check positional arguments for file paths
            for arg in args:
                if isinstance(arg, str) and arg.lower().endswith(('.xlsx', '.xls')):
                    excel_path = arg
                    break
            
            # If not found, check keyword arguments
            if excel_path is None:
                for key, value in kwargs.items():
                    if ('path' in key.lower() or 'file' in key.lower()) and isinstance(value, str) and value.lower().endswith(('.xlsx', '.xls')):
                        excel_path = value
                        break
            
            if excel_path is None:
                # No Excel file found in arguments, re-raise
                raise
            
            # Try to recover the file
            logger.warning(f"Excel file access error, attempting recovery: {excel_path}")
            
            try:
                from .excel_recovery import repair_excel_file_access
                success, recovery_path = repair_excel_file_access(excel_path)
                
                if success:
                    logger.info(f"Successfully recovered file: {recovery_path}")
                    
                    # Replace the file path in the arguments
                    if excel_path in args:
                        new_args = list(args)
                        for i, arg in enumerate(new_args):
                            if arg == excel_path:
                                new_args[i] = recovery_path
                        args = tuple(new_args)
                    else:
                        # Must be in kwargs
                        for key, value in kwargs.items():
                            if value == excel_path:
                                kwargs[key] = recovery_path
                    
                    # Try again with the recovered file
                    return func(*args, **kwargs)
                else:
                    # Recovery failed, re-raise the original error
                    raise
            except ImportError:
                # Recovery module not available, re-raise the original error
                raise
            except Exception as recovery_error:
                # Recovery failed with a different error, log and re-raise the original
                logger.error(f"Recovery failed: {recovery_error}")
                raise
    
    return wrapper  # type: ignore

def apply_windows_patches():
    """Apply Windows-specific patches to key functions."""
    if not IS_WINDOWS or not WINDOWS_MODULES_LOADED:
        return
    
    # Import target modules
    try:
        from ..core import excel_processor
        from . import file_operations
        
        # Patch file operations
        file_operations.safe_file_copy = with_windows_path_handling(file_operations.safe_file_copy)
        file_operations.create_verified_copy = with_windows_path_handling(file_operations.create_verified_copy)
        
        # Patch Excel processor methods
        excel_processor.ExcelProcessor.process_file = with_excel_recovery(excel_processor.ExcelProcessor.process_file)
        
        logger.info("Applied Windows-specific patches to core functions")
    except ImportError as e:
        logger.warning(f"Could not apply Windows patches: {e}")

def initialize_windows_integration():
    """Initialize all Windows-specific integrations."""
    if not IS_WINDOWS:
        logger.warning("Not running on Windows, skipping Windows integration")
        return False
    
    if not WINDOWS_MODULES_LOADED:
        logger.critical("Windows modules not loaded, cannot initialize Windows integration")
        return False
    
    try:
        # Reset Excel automation in case of previous issues
        reset_excel_automation()
        
        # Clean up any Excel temp files
        cleanup_excel_temp_files()
        
        # Initialize resource tracking
        from ..utils.win_api import setup_resource_tracking
        setup_resource_tracking()
        
        # Apply Windows-specific patches
        apply_windows_patches()
        
        logger.info("Windows integration initialized successfully")
        return True
    except Exception as e:
        logger.error(f"Error initializing Windows integration: {e}")
        return False

# Initialize Windows integration when the module is imported
if IS_WINDOWS:
    initialize_windows_integration()