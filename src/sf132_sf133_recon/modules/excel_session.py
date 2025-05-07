"""
Robust Excel session management for Windows.

This module provides advanced Excel session management with Windows-specific
error handling and recovery mechanisms.
"""

import os
import sys
import logging
import time
import contextlib
import gc
from typing import Optional, Dict, Any, Callable, Union

# Initialize logger
logger = logging.getLogger(__name__)

# Only import Windows-specific modules on Windows
if sys.platform.startswith('win'):
    try:
        import pythoncom
        import win32com.client
        import win32api
        import win32con
        
        from ..utils.win_api import (
            reset_excel_automation,
            cleanup_excel_temp_files,
            cleanup_resources
        )
        
        # Excel constants
        XL_VISIBLE = True  # Set to False for production, True for debugging
        XL_ALERTS_ENABLED = False
        
        # Excel calculation modes
        XL_CALCULATION_MANUAL = -4135
        XL_CALCULATION_AUTOMATIC = -4105
        
        # Excel file formats
        XL_OPEN_XML_WORKBOOK = 51
        
        WIN32COM_AVAILABLE = True
    except ImportError as e:
        if sys.platform.startswith('win'):
            logger.critical(f"Failed to load COM automation modules: {e}")
            print(f"CRITICAL ERROR: COM automation modules missing: {e}")
            print("Please install required dependencies: pip install pywin32")
            raise ImportError(f"COM automation modules required: {e}") from e
        WIN32COM_AVAILABLE = False
else:
    WIN32COM_AVAILABLE = False
    logger.warning("Not running on Windows - Excel COM automation unavailable")

def is_com_initialized() -> bool:
    """
    Check if COM is initialized on the current thread.
    
    Returns:
        bool: True if COM is initialized
    """
    if not sys.platform.startswith('win'):
        return False
        
    if not WIN32COM_AVAILABLE:
        return False
        
    try:
        # Try to create a simple COM object
        test_obj = win32com.client.Dispatch("WScript.Shell")
        del test_obj
        return True
    except:
        return False

@contextlib.contextmanager
def excel_com_session(visible: bool = False, disable_alerts: bool = True):
    """
    Context manager for Excel COM automation session.
    
    Args:
        visible: Whether Excel should be visible
        disable_alerts: Whether to disable Excel alerts
        
    Yields:
        tuple: (Excel Application object, session info dict)
    """
    if not sys.platform.startswith('win'):
        raise RuntimeError("Excel COM automation only available on Windows")
        
    if not WIN32COM_AVAILABLE:
        raise ImportError("Required COM automation modules not available")
    
    excel = None
    init_com = not is_com_initialized()
    
    try:
        # Initialize COM
        if init_com:
            pythoncom.CoInitialize()
        
        # Start Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = not disable_alerts
        
        # Session info for caller
        session_info = {
            "initialized_com": init_com,
            "start_time": time.time()
        }
        
        # Yield both Excel and session info
        yield excel, session_info
        
    except Exception as e:
        logger.error(f"Error in Excel COM session: {e}")
        raise
    
    finally:
        # Clean up
        if excel:
            try:
                excel.Quit()
            except:
                pass
            excel = None
        
        # Force garbage collection
        gc.collect()
        
        # Uninitialize COM if we initialized it
        if init_com:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

@contextlib.contextmanager
def robust_excel_session(visible: bool = False, disable_alerts: bool = True, 
                        retry_count: int = 2):
    """
    Enhanced context manager for Excel COM automation with robust error handling.
    
    Args:
        visible: Whether Excel should be visible
        disable_alerts: Whether to disable Excel alerts
        retry_count: Number of retries on failure
        
    Yields:
        tuple: (Excel Application object, session info dict)
    """
    if not sys.platform.startswith('win'):
        raise RuntimeError("Excel COM automation only available on Windows")
        
    if not WIN32COM_AVAILABLE:
        raise ImportError("Required COM automation modules not available")
    
    excel = None
    init_com = not is_com_initialized()
    attempts = 0
    
    # Clean up any Excel temp files first
    cleanup_excel_temp_files()
    
    while attempts <= retry_count:
        try:
            # Initialize COM
            if init_com:
                pythoncom.CoInitialize()
            
            # Reset Excel automation if needed
            if attempts > 0:
                logger.info(f"Excel session retry {attempts}/{retry_count}")
                reset_excel_automation()
                
                # Clean up resources between attempts
                cleanup_resources()
                gc.collect()
                time.sleep(1)  # Give the system time to release resources
                
            # Start Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = visible
            excel.DisplayAlerts = not disable_alerts
            
            # Extra settings for robustness
            excel.EnableEvents = False  # Disable Excel events
            excel.AskToUpdateLinks = False
            excel.Calculation = XL_CALCULATION_MANUAL  # Manual calculation
            
            # Session info for caller
            session_info = {
                "initialized_com": init_com,
                "start_time": time.time(),
                "attempt": attempts + 1
            }
            
            # Yield both Excel and session info
            yield excel, session_info
            
            # If we get here, everything worked
            break
            
        except Exception as e:
            attempts += 1
            logger.error(f"Error in Excel COM session (attempt {attempts}): {e}")
            
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
                excel = None
            
            # Force garbage collection
            gc.collect()
            
            # Uninitialize COM if we initialized it
            if init_com:
                try:
                    pythoncom.CoUninitialize()
                    init_com = False  # Already uninitialized
                except:
                    pass
            
            # Re-raise on last attempt
            if attempts > retry_count:
                raise
            
            # Wait before retrying
            time.sleep(2)
    
    # Final cleanup
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
            excel = None
        
        # Force garbage collection
        gc.collect()
        
        # Uninitialize COM if we initialized it
        if init_com:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

def repair_excel_workbook(file_path: str, output_path: Optional[str] = None) -> bool:
    """
    Repair a potentially corrupted Excel workbook using COM.
    
    Args:
        file_path: Path to the Excel file
        output_path: Output path (defaults to overwriting the input file)
        
    Returns:
        bool: True if repair was successful
    """
    if not sys.platform.startswith('win'):
        logger.warning("Excel COM repair only available on Windows")
        return False
        
    if not WIN32COM_AVAILABLE:
        logger.warning("Required COM automation modules not available")
        return False
    
    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        return False
    
    # Default output path to input path if not specified
    if not output_path:
        output_path = file_path
    
    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    success = False
    
    # Use the robust Excel session for repair
    try:
        with robust_excel_session(visible=False, retry_count=3) as (excel, _):
            # Open with repair mode
            wb = excel.Workbooks.Open(
                file_path,
                UpdateLinks=0,  # Don't update links
                ReadOnly=False,
                CorruptLoad=2  # xlRepairFile - repair mode
            )
            
            # Calculate once to ensure data is current
            wb.Calculate()
            
            # Save to output path
            wb.SaveAs(
                output_path,
                FileFormat=XL_OPEN_XML_WORKBOOK,
                CreateBackup=False
            )
            
            # Close workbook
            wb.Close(SaveChanges=False)
            
            # Check if the file was created successfully
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                logger.info(f"Successfully repaired Excel file: {output_path}")
                success = True
            else:
                logger.error(f"Failed to save repaired file: {output_path}")
                
    except Exception as e:
        logger.error(f"Error repairing Excel file: {e}")
    
    return success