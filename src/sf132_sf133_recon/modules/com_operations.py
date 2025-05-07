"""
COM-based Excel operations module.

This module provides functions for working with Excel files via the COM interface,
primarily for Windows systems.
"""

import os
import sys
import time
import gc
import logging
import shutil
from typing import Optional, Callable, Any, Dict, List, Tuple
import contextlib

# Type alias
PathLike = str
ProgressCallback = Callable[[float, str], None]
StatusCallback = Callable[[str], None]

# Initialize logger
logger = logging.getLogger(__name__)

# Check if Windows COM modules are available
try:
    import pythoncom
    import win32com.client
    WINDOWS_COM_AVAILABLE = True
except ImportError:
    WINDOWS_COM_AVAILABLE = False

def is_windows_platform() -> bool:
    """
    Check if running on Windows platform.
    
    Returns:
        bool: True if running on Windows
    """
    return sys.platform.startswith('win')

def create_clean_copy(original_file: PathLike, new_file: PathLike) -> bool:
    """
    Create a clean copy of the original file using COM if possible.
    
    Args:
        original_file: Original file path
        new_file: New file path
        
    Returns:
        bool: Whether the operation was successful
    """
    if not WINDOWS_COM_AVAILABLE or not is_windows_platform():
        logger.warning("COM operations not available on this platform")
        return False
        
    try:
        # First make sure the output directory exists
        os.makedirs(os.path.dirname(new_file), exist_ok=True)
        
        # Create a temporary file for intermediate processing
        from .file_operations import get_temp_file_path
        temp_file = get_temp_file_path("init_copy")
        
        # Try opening the original file to verify its integrity
        logger.info("Verifying original file integrity...")
        
        # Do a standard file copy first to get a basic working file
        logger.info("Creating initial file copy...")
        shutil.copy2(original_file, temp_file)
        
        # Now use Excel COM to save the file properly
        logger.info("Creating clean Excel copy...")
        return save_with_excel_com(temp_file, new_file, preserve_all=True)
        
    except Exception as e:
        logger.error(f"Error creating clean copy: {str(e)}")
        return False

def save_with_excel_com(source_file: PathLike, dest_file: PathLike, preserve_all: bool = False) -> bool:
    """
    Use Excel COM to save a file, with robust fallback mechanisms.
    
    Args:
        source_file: Source file path
        dest_file: Destination file path
        preserve_all: Whether to preserve all Excel features
        
    Returns:
        bool: Whether the save was successful
    """
    if not WINDOWS_COM_AVAILABLE or not is_windows_platform():
        logger.warning("COM operations not available on this platform")
        return False
        
    # Check if file exists before attempting COM
    if not os.path.exists(source_file):
        logger.error(f"Source file does not exist: {source_file}")
        return False
        
    # Make sure Excel is not running to avoid conflicts
    from .file_operations import close_excel_instances
    close_excel_instances()
    time.sleep(1)  # Give OS time to fully release resources
    
    excel = None
    wb = None
    success = False
    retry_count = 0
    max_retries = 2
    
    while retry_count <= max_retries and not success:
        try:
            # Initialize COM with proper threading model
            pythoncom.CoInitialize()
            
            # Start Excel with proper security context
            logger.info(f"Starting Excel (attempt {retry_count+1}/{max_retries+1})...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False  # Disable Excel events
            
            # Set calculation mode to manual to prevent recalculation issues
            excel.Calculation = -4135  # xlCalculationManual
            
            # Open workbook with safe options
            logger.info(f"Opening file with Excel COM: {source_file}")
            
            # Use explicit constants for better clarity
            XL_UPDATE_LINKS_NEVER = 0
            XL_CORRUPT_LOAD_NORMAL = 0
            
            # Open with more robust error handling
            wb = excel.Workbooks.Open(
                source_file,
                UpdateLinks=XL_UPDATE_LINKS_NEVER,  
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                CorruptLoad=XL_CORRUPT_LOAD_NORMAL
            )
            
            # Clean up external connections if not preserving everything
            if not preserve_all:
                clean_external_connections(wb, excel)
            
            # Make sure the workbook is recalculated
            wb.Calculate()
            
            # Save to the destination with correct format
            logger.info(f"Saving clean copy to: {dest_file}")
            
            # Use explicit constant for better clarity
            XL_XLSX = 51  # Excel.XlFileFormat.xlOpenXMLWorkbook
            
            # Add repair functionality to save process
            try:
                # First save to a temporary location
                from .file_operations import get_temp_file_path
                temp_save = get_temp_file_path("com_save")
                
                wb.SaveAs(
                    temp_save,
                    FileFormat=XL_XLSX,
                    CreateBackup=False
                )
                
                # Now let Excel repair the file and save it again
                # Close the current workbook
                wb.Close(SaveChanges=False)
                
                # Reopen with repair option
                wb = excel.Workbooks.Open(
                    temp_save,
                    UpdateLinks=0,
                    CorruptLoad=2  # xlRepairFile - repair mode
                )
                
                # Save to final destination
                wb.SaveAs(
                    dest_file,
                    FileFormat=XL_XLSX,
                    CreateBackup=False
                )
                
            except Exception:
                # Fall back to direct save if repair approach fails
                logger.warning("Repair approach failed, using direct save")
                wb.SaveAs(
                    dest_file,
                    FileFormat=XL_XLSX,
                    CreateBackup=False
                )
            
            # Close properly
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            # Verify the file was created successfully
            if os.path.exists(dest_file) and os.path.getsize(dest_file) > 0:
                success = True
                logger.info(f"Excel COM save operation successful")
            else:
                raise IOError(f"Failed to save file to {dest_file}")
            
        except Exception as e:
            retry_count += 1
            logger.warning(f"COM attempt {retry_count}/{max_retries+1} failed: {str(e)}")
            
            if retry_count <= max_retries:
                logger.info(f"Retrying COM operation (attempt {retry_count+1}/{max_retries+1})...")
                time.sleep(2)  # Wait before retrying
                # Make sure Excel is fully closed
                close_excel_instances()
            else:
                logger.error(f"All COM attempts failed: {str(e)}")
        finally:
            # Thorough cleanup of COM resources
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
                wb = None
            
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
                excel = None
            
            # Force garbage collection to release COM objects
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            
            # Give OS time to release resources
            time.sleep(0.5)
            
    return success

def clean_external_connections(wb, excel) -> None:
    """
    Clean up external data connections and links in the workbook.
    
    Args:
        wb: Excel workbook COM object
        excel: Excel application COM object
    """
    try:
        logger.info("Cleaning external data connections...")
        
        # Import the specialized cleaner
        from ..modules.excel_cleaner import ExcelDataCleaner
        
        # Initialize the specialized external data cleaner
        data_cleaner = ExcelDataCleaner()
        
        # Use the specialized sheet cleaning method for each sheet
        for i in range(1, wb.Sheets.Count + 1):
            try:
                sheet = wb.Sheets(i)
                data_cleaner.clean_sheet_data(sheet)
            except Exception as e:
                logger.warning(f"Error cleaning sheet {i}: {e}")
        
        # Clean general workbook connections
        try:
            # Break external links
            for link_type in [1, 2, 5, 6]:  # Different types of links
                try:
                    logger.info(f"Breaking links of type {link_type}...")
                    wb.BreakLink(Name="", Type=link_type)
                except:
                    pass
                
            # Remove connections
            if hasattr(wb, 'Connections') and wb.Connections.Count > 0:
                logger.info(f"Removing {wb.Connections.Count} external connections...")
                
                conn_names = []
                for i in range(1, wb.Connections.Count + 1):
                    try:
                        conn_names.append(wb.Connections(i).Name)
                    except:
                        pass
                
                for name in conn_names:
                    try:
                        wb.Connections(name).Delete()
                    except Exception as e:
                        logger.warning(f"Failed to remove connection {name}: {e}")
            
            # Commit connection changes
            try:
                if hasattr(wb, 'Connections'):
                    wb.Connections.CommitAll()
            except:
                pass
                
        except Exception as e:
            logger.warning(f"Error cleaning workbook connections: {e}")
            
    except Exception as e:
        logger.warning(f"Error during connection cleanup: {e}")

@contextlib.contextmanager
def excel_com_session(visible: bool = False):
    """
    Context manager for safely using an Excel COM session.
    
    Args:
        visible: Whether Excel should be visible
        
    Yields:
        object: Excel application COM object or None if not available
    """
    if not WINDOWS_COM_AVAILABLE or not is_windows_platform():
        yield None
        return
        
    excel = None
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Start Excel with proper error handling
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        
        # Yield for operations
        yield excel
        
    except Exception as e:
        logger.error(f"Error in Excel COM session: {e}")
        yield None
        
    finally:
        # Clean up resources
        if excel:
            try:
                excel.Quit()
            except:
                pass
            del excel
            
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def process_workbook_with_com(
    file_path: PathLike, 
    password: Optional[str], 
    config: Any, 
    update_progress: ProgressCallback,
    update_status: StatusCallback
) -> bool:
    """
    Process a workbook using COM automation.
    
    Args:
        file_path: Path to Excel file
        password: Sheet protection password
        config: Processing configuration
        update_progress: Callback for progress updates
        update_status: Callback for status updates
        
    Returns:
        bool: Whether processing was successful
    """
    if not WINDOWS_COM_AVAILABLE or not is_windows_platform():
        update_status("COM automation not available on this platform")
        return False
        
    file_path = os.path.abspath(file_path)
    
    update_status(f"Checking if file exists: {file_path}")
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File does not exist: {file_path}")
    
    from .file_operations import close_excel_instances
    close_excel_instances()
    time.sleep(2)
    
    excel = None
    wb = None
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        update_status("Starting Excel application...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        # Disable events and calculations for better performance
        excel.EnableEvents = False
        excel.Calculation = -4135  # xlCalculationManual
        
        update_status(f"Opening workbook: {file_path}")
        wb = excel.Workbooks.Open(
            file_path,
            UpdateLinks=0,
            ReadOnly=False,
            CorruptLoad=2  # xlRepairFile - better error handling
        )    
        
        update_status("Selecting worksheet...")
        ws = wb.Worksheets(config.sheet_name)
            
        # Process using COM operations
        process_rows_with_pywin32(
            ws,
            config,
            update_progress,
            update_status
        ) 
        
        # Calculate before saving
        wb.Calculate()
        
        update_status("Saving workbook...")
        wb.Save()    
        
        return True
        
    except Exception as e:
        update_status(f"Error in Excel processing: {str(e)}")
        logger.error(f"Detailed error: {str(e)}", exc_info=True)
        raise
    finally:
        # Clean up COM objects
        if wb:
            try:
                wb.Close(SaveChanges=True)
            except Exception as e:
                logger.error(f"Error closing workbook: {str(e)}")
        
        if excel:
            try:
                excel.Quit()
            except Exception as e:
                logger.error(f"Error quitting Excel: {str(e)}")
            
            del wb
            del excel
            
        # Force garbage collection to release COM objects
        gc.collect()
        pythoncom.CoUninitialize()

def process_rows_with_pywin32(
    ws, 
    config: Any,
    update_progress: ProgressCallback,
    update_status: StatusCallback
) -> None:
    """
    Process individual rows with pywin32.
    
    Args:
        ws: pywin32 worksheet object
        config: Processing configuration
        update_progress: Callback for progress updates
        update_status: Callback for status updates
    """
    try:
        header_row = config.header_row
        headers_to_find = config.headers_to_find
        
        # Find column indexes
        column_indexes = {}
        for col in range(1, 100):  # Reasonable limit
            header_value = ws.Cells(header_row, col).Value
            if header_value in headers_to_find:
                column_indexes[header_value] = col
        
        missing_headers = [h for h in headers_to_find if h not in column_indexes]
        if missing_headers:
            raise ValueError(f"Missing headers: {', '.join(missing_headers)}")
        
        # Find last column
        last_col = ws.UsedRange.Columns.Count
        
        # Find matching row (end of data)
        matching_row = 0
        header_color = ws.Cells(header_row, 1).Interior.Color
        
        for row in range(header_row + 1, 1000):  # Reasonable limit
            if ws.Cells(row, 1).Interior.Color == header_color:
                matching_row = row
                break
        
        if not matching_row:
            matching_row = ws.UsedRange.Rows.Count
        
        # Add DO Comments column
        update_status("Adding DO Comments column...")
        ws.Cells(header_row, last_col + 1).Value = "DO Comments"
        ws.Cells(header_row, last_col + 1).Interior.Color = 65535  # Yellow
        ws.Cells(header_row, last_col + 1).Font.Bold = True
        ws.Cells(header_row, last_col + 1).Font.Color = 255  # Red
        ws.Cells(header_row, last_col + 1).HorizontalAlignment = -4108  # Center
        ws.Columns(last_col + 1).ColumnWidth = 25
        
        # Process rows
        processed_count = 0
        
        for row in range(header_row + 1, matching_row):
            try:
                difference_value = ws.Cells(row, column_indexes["Difference"]).Value
                include_cfo_value = ws.Cells(row, column_indexes["Include in CFO Cert Letter"]).Value
                explanation_value = ws.Cells(row, column_indexes["Explanation"]).Value
                
                if difference_value not in (None, ""):
                    if include_cfo_value == "N" and explanation_value not in (None, 0, ""):
                        ws.Cells(row, last_col + 1).Value = "Explanation Reasonable"
                        processed_count += 1
                    elif include_cfo_value == "Y" and explanation_value not in (None, ""):
                        ws.Cells(row, last_col + 1).Value = "Explanation Reasonable; Include in CFO Cert Letter"
                        processed_count += 1
                    elif explanation_value in (None, "", 0) and difference_value != 0:
                        ws.Cells(row, last_col + 1).Value = "Explanation Required"
                        # Highlight cells that need attention
                        ws.Cells(row, last_col + 1).Interior.Color = 0xFF9999  # Light red
                        processed_count += 1
                        
            except Exception as e:
                update_status(f"Error processing row {row}: {str(e)}")
        
        update_status(f"Successfully processed {processed_count} rows")
        update_progress(100, "Processing complete")
        
    except Exception as e:
        logger.error(f"Error in pywin32 processing: {str(e)}", exc_info=True)
        raise