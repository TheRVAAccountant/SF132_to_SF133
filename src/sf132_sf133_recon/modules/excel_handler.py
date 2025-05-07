"""
Excel process and session handling module.

This module provides functions for managing Excel processes
and sessions, with special handling for Windows platforms.
"""

import os
import sys
import time
import logging
import contextlib
from typing import List, Dict, Optional, Any, Union, Callable

logger = logging.getLogger(__name__)

def is_windows() -> bool:
    """
    Check if running on Windows platform.
    
    Returns:
        bool: True if running on Windows
    """
    return sys.platform.startswith('win')

def close_excel_instances() -> List[int]:
    """
    Terminate all existing Excel processes to prevent file locking.
    
    Returns:
        List[int]: List of terminated process IDs
    """
    terminated_pids = []
    
    try:
        # Using psutil for cross-platform process management
        import psutil
        
        # First find all Excel processes
        excel_pids = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_name = proc.info['name']
                # Check for Excel process name on different platforms
                if proc_name and (
                    'EXCEL.EXE' in proc_name.upper() or  # Windows
                    'MICROSOFT EXCEL' in proc_name.upper()  # Mac
                ):
                    excel_pids.append(proc.info['pid'])
            except Exception as e:
                logger.debug(f"Error checking process: {e}")
        
        if excel_pids:
            logger.info(f"Found {len(excel_pids)} Excel processes to close")
            
        # Then terminate each process with proper cleanup
        for pid in excel_pids:
            try:
                proc = psutil.Process(pid)
                proc.terminate()
                proc.wait(timeout=5)
                terminated_pids.append(pid)
            except Exception as e:
                logger.warning(f"Failed to terminate Excel process (PID {pid}): {e}")
                # Try more aggressive termination
                try:
                    if psutil.pid_exists(pid):
                        if is_windows():
                            os.kill(pid, 9)  # SIGKILL on Windows
                        else:
                            os.kill(pid, 9)  # SIGKILL on Unix
                        terminated_pids.append(pid)
                except Exception as kill_e:
                    logger.warning(f"Failed to kill Excel process (PID {pid}): {kill_e}")
        
        # Verify all processes are terminated
        remaining = []
        for pid in excel_pids:
            if psutil.pid_exists(pid):
                remaining.append(pid)
                
        if remaining:
            logger.warning(f"Could not terminate {len(remaining)} Excel processes: {remaining}")
        else:
            logger.info("All Excel processes successfully closed")
            
    except ImportError:
        logger.warning("psutil module not available, cannot close Excel instances")
    
    return terminated_pids

def force_garbage_collection():
    """Force Python garbage collection to release resources."""
    import gc
    gc.collect()
    # Allow a moment for the OS to release resources
    time.sleep(0.5)

@contextlib.contextmanager
def excel_process_guard():
    """
    Context manager that ensures Excel processes are properly terminated.
    """
    try:
        # Close any existing Excel processes before starting
        close_excel_instances()
        yield
    finally:
        # Make sure to clean up any Excel processes when done
        close_excel_instances()
        force_garbage_collection()

def unlock_excel_file(file_path: str) -> bool:
    """
    Attempt to unlock an Excel file that might be locked by a process.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        bool: True if successful or file wasn't locked
    """
    if not os.path.exists(file_path):
        return True  # Nothing to unlock
        
    # First check if the file is actually locked
    try:
        # Try to open the file in write mode
        with open(file_path, 'rb+') as f:
            pass  # If we can open it, it's not locked
        return True
    except (IOError, PermissionError):
        # File is locked, attempt to unlock
        logger.info(f"File appears to be locked: {file_path}")
        
        # Close Excel processes that might have the file open
        close_excel_instances()
        
        # Wait a moment for the OS to release the file
        time.sleep(1)
        
        # Check if the file is now unlocked
        try:
            with open(file_path, 'rb+') as f:
                pass
            logger.info(f"Successfully unlocked file: {file_path}")
            return True
        except (IOError, PermissionError):
            logger.warning(f"Failed to unlock file: {file_path}")
            return False

def excel_com_available() -> bool:
    """
    Check if Excel COM automation is available.
    
    Returns:
        bool: True if Excel COM automation is available
    """
    if not is_windows():
        return False
        
    try:
        import pythoncom
        import win32com.client
        return True
    except ImportError:
        return False