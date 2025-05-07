"""
Windows integration module for Excel processing application.
This module patches the necessary components to integrate Windows-specific
file handling and error recovery without modifying the original code.
"""

import os
import sys
import logging
import inspect
import functools
import importlib
from types import ModuleType
from typing import Any, Callable, Dict, Optional, Tuple, List, Union

# Initialize logger
logger = logging.getLogger(__name__)

# Track whether integration has been applied
_integration_applied = False

def is_windows() -> bool:
    """Check if running on Windows."""
    return sys.platform.startswith('win')

def import_module_if_available(module_name: str) -> Optional[ModuleType]:
    """
    Import a module if it's available, otherwise return None.
    
    Args:
        module_name: Name of the module to import
        
    Returns:
        Optional[ModuleType]: Imported module or None
    """
    try:
        return importlib.import_module(module_name)
    except ImportError:
        return None

def patch_method(cls: type, method_name: str, decorator: Callable) -> None:
    """
    Patch a method in a class with a decorator.
    
    Args:
        cls: Class to patch
        method_name: Method name to patch
        decorator: Decorator function to apply
    """
    if hasattr(cls, method_name):
        original_method = getattr(cls, method_name)
        setattr(cls, method_name, decorator(original_method))
        logger.debug(f"Patched {cls.__name__}.{method_name}")

def file_operation_decorator(func: Callable) -> Callable:
    """
    Decorator to add Windows-specific file handling to file operations.
    
    Args:
        func: Original function to decorate
        
    Returns:
        Callable: Decorated function
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # Only apply Windows handling on Windows
        if not is_windows():
            return func(*args, **kwargs)
            
        # Get file path argument - assume it's the first argument
        # that looks like a string path
        file_path = None
        for arg in args:
            if isinstance(arg, str) and (
                arg.endswith('.xlsx') or 
                arg.endswith('.xls') or
                '\\' in arg or
                '/' in arg
            ):
                file_path = arg
                break
                
        # Also check kwargs for file_path, path, or similar
        if not file_path:
            for key in ('file_path', 'filepath', 'path', 'source', 'filename'):
                if key in kwargs and isinstance(kwargs[key], str):
                    file_path = kwargs[key]
                    break
                    
        # If we found a file path, normalize it
        if file_path:
            try:
                # Import Windows path handler
                win_path_handler = import_module_if_available('win_path_handler')
                if win_path_handler:
                    normalized_path = win_path_handler.normalize_windows_path(file_path)
                    
                    # Update args or kwargs with normalized path
                    if file_path in args:
                        args_list = list(args)
                        for i, arg in enumerate(args_list):
                            if arg == file_path:
                                args_list[i] = normalized_path
                        args = tuple(args_list)
                    else:
                        for key in kwargs:
                            if kwargs[key] == file_path:
                                kwargs[key] = normalized_path
                                break
            except Exception as e:
                logger.warning(f"Error normalizing path: {e}")
        
        try:
            # Run the original function
            result = func(*args, **kwargs)
            return result
        except Exception as e:
            # Handle file access errors
            if file_path and is_windows():
                error_msg = str(e).lower()
                if any(pattern in error_msg for pattern in (
                    'process cannot access', 
                    'being used by another process',
                    'permission denied',
                    'file in use'
                )):
                    logger.info(f"Caught file access error: {e}")
                    
                    try:
                        # Import recovery module
                        recovery_module = import_module_if_available('excel_file_recovery')
                        if recovery_module:
                            logger.info("Attempting automatic recovery")
                            success, result_path = recovery_module.handle_excel_process_error(
                                str(e), file_path
                            )
                            
                            if success:
                                logger.info(f"Recovery successful, retrying with: {result_path}")
                                
                                # Replace the file path in args or kwargs
                                if file_path in args:
                                    args_list = list(args)
                                    for i, arg in enumerate(args_list):
                                        if arg == file_path:
                                            args_list[i] = result_path
                                    args = tuple(args_list)
                                else:
                                    for key in kwargs:
                                        if kwargs[key] == file_path:
                                            kwargs[key] = result_path
                                            break
                                            
                                # Retry the function call
                                return func(*args, **kwargs)
                    except Exception as recovery_error:
                        logger.error(f"Recovery attempt failed: {recovery_error}")
            
            # If we couldn't recover, re-raise the original exception
            raise
            
    return wrapper

def excel_processor_decorator(func: Callable) -> Callable:
    """
    Decorator specifically for Excel processing operations.
    
    Args:
        func: Original function to decorate
        
    Returns:
        Callable: Decorated function
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # Only apply on Windows
        if not is_windows():
            return func(*args, **kwargs)
            
        # Before processing, ensure any lingering Excel instances are closed
        try:
            file_ops_module = import_module_if_available('file_operations_win')
            if file_ops_module:
                file_ops_module.close_excel_instances()
        except Exception as e:
            logger.warning(f"Error closing Excel instances: {e}")
            
        try:
            # Run the original function
            return func(*args, **kwargs)
        except Exception as e:
            # Try processing with recovery
            try:
                recovery_module = import_module_if_available('excel_file_recovery')
                if recovery_module:
                    logger.info(f"Attempting to process with recovery due to: {e}")
                    success, result = recovery_module.process_with_recovery(func, *args, **kwargs)
                    if success:
                        return result
            except Exception as recovery_error:
                logger.error(f"Recovery processing failed: {recovery_error}")
                
            # If recovery failed, re-raise the original exception
            raise
            
    return wrapper

def cleanup_decorator(func: Callable) -> Callable:
    """
    Decorator to ensure proper resource cleanup after function execution.
    
    Args:
        func: Original function to decorate
        
    Returns:
        Callable: Decorated function
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            # Run the original function
            return func(*args, **kwargs)
        finally:
            # Clean up resources
            if is_windows():
                try:
                    file_ops_module = import_module_if_available('file_operations_win')
                    if file_ops_module:
                        file_ops_module.cleanup_all_resources()
                except Exception as e:
                    logger.warning(f"Error during resource cleanup: {e}")
                    
    return wrapper

def patch_excel_processor() -> None:
    """Patch the ExcelProcessor class with Windows-specific enhancements."""
    try:
        # Import the ExcelProcessor class
        from excel_processor import ExcelProcessor
        
        # File operation methods to patch
        file_methods = [
            '_create_verified_copy',
            '_save_with_excel_com',
            '_direct_file_copy',
            '_create_clean_copy',
            '_generate_new_filename',
        ]
        
        # Processing methods to patch
        process_methods = [
            'process_file',
            '_process_with_fresh_workbook',
            '_process_with_legacy_method',
            '_process_workbook',
        ]
        
        # Cleanup methods to patch
        cleanup_methods = [
            '__del__',
            '_cleanup_temp_files',
        ]
        
        # Apply patches
        for method_name in file_methods:
            patch_method(ExcelProcessor, method_name, file_operation_decorator)
            
        for method_name in process_methods:
            patch_method(ExcelProcessor, method_name, excel_processor_decorator)
            
        for method_name in cleanup_methods:
            patch_method(ExcelProcessor, method_name, cleanup_decorator)
            
        logger.info("Successfully patched ExcelProcessor for Windows compatibility")
        
    except Exception as e:
        logger.error(f"Error patching ExcelProcessor: {e}")

def patch_application_class() -> None:
    """Patch the ExcelProcessingApplication class with Windows-specific enhancements."""
    try:
        # Import the application class
        from main import ExcelProcessingApplication
        
        # Methods to patch
        methods = [
            'process_file',
            'run_cli',
        ]
        
        # Apply patches
        for method_name in methods:
            patch_method(ExcelProcessingApplication, method_name, excel_processor_decorator)
            
        logger.info("Successfully patched ExcelProcessingApplication for Windows compatibility")
        
    except Exception as e:
        logger.error(f"Error patching ExcelProcessingApplication: {e}")

def monkey_patch_os_functions() -> None:
    """Monkey patch os module functions for better Windows compatibility."""
    if not is_windows():
        return
        
    # Save original functions
    original_path_exists = os.path.exists
    original_makedirs = os.makedirs
    
    # Import Windows path handler
    win_path_handler = import_module_if_available('win_path_handler')
    if not win_path_handler:
        return
        
    # Path normalization wrapper
    def normalized_path_exists(path):
        if isinstance(path, str):
            path = win_path_handler.normalize_windows_path(path)
        return original_path_exists(path)
        
    def normalized_makedirs(path, mode=0o777, exist_ok=False):
        if isinstance(path, str):
            path = win_path_handler.normalize_windows_path(path)
        return original_makedirs(path, mode, exist_ok)
        
    # Apply monkey patches
    os.path.exists = normalized_path_exists
    os.makedirs = normalized_makedirs
    
    logger.info("Applied os function monkey patches for Windows compatibility")

def setup_windows_integration() -> None:
    """Set up Windows-specific integrations if running on Windows."""
    global _integration_applied
    
    if _integration_applied:
        return
        
    if not is_windows():
        logger.info("Not running on Windows, skipping Windows integration")
        return
        
    logger.info("Setting up Windows integration")
    
    try:
        # Import required modules
        win_path_handler = import_module_if_available('win_path_handler')
        file_ops_module = import_module_if_available('file_operations_win')
        excel_handler = import_module_if_available('excel_handler_win')
        
        if not all([win_path_handler, file_ops_module, excel_handler]):
            logger.warning("Some Windows integration modules not available")
            return
            
        # Set up resource tracking
        file_ops_module.setup_resource_tracking()
        
        # Apply patches
        patch_excel_processor()
        patch_application_class()
        monkey_patch_os_functions()
        
        # Pre-clean any Excel processes that might be running
        file_ops_module.close_excel_instances()
        
        # Mark integration as applied
        _integration_applied = True
        
        logger.info("Windows integration setup complete")
        
    except Exception as e:
        logger.error(f"Error setting up Windows integration: {e}")

# Automatic integration when module is imported
if is_windows():
    setup_windows_integration()