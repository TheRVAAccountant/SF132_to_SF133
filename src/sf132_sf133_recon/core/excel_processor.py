"""
Core Excel processor for SF132/SF133 reconciliation.

This module contains the primary ExcelProcessor class that handles the main
Excel processing operations for SF132 to SF133 reconciliation.
"""

import os
import time
import logging
import gc
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Dict, Tuple, Optional, Any, List, Union
from queue import Queue

# Try to import Excel-related modules - they might be unavailable on some platforms
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Border, Side, Alignment, Color
    from openpyxl.styles.colors import COLOR_INDEX
    from openpyxl.cell import MergedCell
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.exceptions import InvalidFileException
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# Import Windows-specific modules - these are essential for the application
import sys

# Check if running on Windows
IS_WINDOWS = sys.platform.startswith('win')

# Import Windows-specific modules
if IS_WINDOWS:
    try:
        import pythoncom
        import win32com.client
        import win32api
        WINDOWS_COM_AVAILABLE = True
    except ImportError as e:
        # On Windows, these imports are critical
        raise ImportError(f"Critical Windows modules missing: {e}. Please install 'pywin32' package.") from e
else:
    # When not on Windows, note that functionality will be limited
    WINDOWS_COM_AVAILABLE = False
    print("WARNING: Not running on Windows - Excel COM automation unavailable. Functionality will be limited.")

# Local configuration
DEFAULT_SHEET_NAME = "Sheet1"
DEFAULT_HEADER_ROW = 5
DEFAULT_HEADERS_TO_FIND = ["Difference", "Include in CFO Cert Letter", "Explanation"]
DEFAULT_OUTPUT_DIR = "output"

class ExcelProcessor:
    """
    Handles Excel file processing operations including file manipulation,
    formatting, and content analysis for SF132 to SF133 reconciliation.
    """
    
    def __init__(self, queue: Optional[Queue] = None):
        """
        Initialize the Excel processor.
        
        Args:
            queue: Queue for communication with GUI
        """
        self.queue = queue
        self._setup_logging()
        self._temp_files = []  # Track temp files for cleanup
        
        # Default configuration
        self.sheet_name = DEFAULT_SHEET_NAME
        self.header_row = DEFAULT_HEADER_ROW
        self.headers_to_find = DEFAULT_HEADERS_TO_FIND
        self.output_directory = DEFAULT_OUTPUT_DIR
        
        # Create output directory
        os.makedirs(self.output_directory, exist_ok=True)
        
    def _setup_logging(self):
        """Configure logging for the processor."""
        self.logger = logging.getLogger(__name__)
    
    def __del__(self):
        """Clean up resources when instance is destroyed."""
        self._cleanup_temp_files()
        
    def _cleanup_temp_files(self):
        """Clean up all temporary files."""
        for temp_file in self._temp_files:
            if os.path.exists(temp_file):
                try:
                    os.unlink(temp_file)
                    self.logger.debug(f"Cleaned up temp file: {temp_file}")
                except Exception as e:
                    self.logger.warning(f"Failed to clean up temp file {temp_file}: {e}")
    
    def _update_progress(self, value: float, message: str):
        """
        Update progress through the queue.
        
        Args:
            value: Progress value (0-100)
            message: Status message
        """
        if self.queue:
            self.queue.put(("progress", (value, message)))
        self.logger.info(message)
    
    def _update_status(self, message: str):
        """
        Send status update through the queue.
        
        Args:
            message: Status message
        """
        if self.queue:
            self.queue.put(("status", message))
        self.logger.info(message)
    
    def close_excel_instances(self):
        """Terminate all existing Excel processes to prevent file locking."""
        self._update_status("Ensuring all Excel instances are closed...")
        
        try:
            import psutil
            excel_pids = []
            
            # First find all Excel processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    proc_name = proc.info['name']
                    # Check for Excel process name on different platforms
                    if proc_name and (
                        'EXCEL.EXE' in proc_name.upper() or  # Windows
                        'MICROSOFT EXCEL' in proc_name.upper()  # Mac
                    ):
                        excel_pids.append(proc.info['pid'])
                except Exception:
                    pass
            
            if excel_pids:
                self._update_status(f"Found {len(excel_pids)} Excel processes to close")
                
            # Then terminate each process with proper cleanup
            for pid in excel_pids:
                try:
                    proc = psutil.Process(pid)
                    proc.terminate()
                    proc.wait(timeout=5)
                except Exception as e:
                    self.logger.warning(f"Failed to terminate Excel process (PID {pid}): {e}")
                    # Try more aggressive termination
                    try:
                        if psutil.pid_exists(pid):
                            os.kill(pid, 9)  # SIGKILL
                    except Exception:
                        pass
            
            # Verify all processes are terminated
            remaining = []
            for pid in excel_pids:
                if psutil.pid_exists(pid):
                    remaining.append(pid)
                    
            if remaining:
                self.logger.warning(f"Could not terminate {len(remaining)} Excel processes: {remaining}")
            else:
                self._update_status("All Excel processes successfully closed")
        except ImportError:
            self.logger.warning("psutil module not available, cannot close Excel instances")

    def process_file(self, original_file: str, password: str = None) -> bool:
        """
        Main processing function for Excel file.
        
        Args:
            original_file: Path to original Excel file
            password: Sheet protection password
        
        Returns:
            bool: True if processing was successful, False otherwise
        """
        try:
            # Convert to absolute path
            original_file = os.path.abspath(original_file)
            self._update_status(f"Processing file: {original_file}")
            
            self._validate_file(original_file)
            new_file = self._generate_new_filename(original_file)
            
            # Ensure Excel is fully closed before starting
            self.close_excel_instances()
            
            # First make a clean copy of the original file
            temp_copy = self._create_verified_copy(original_file)
            
            # ENHANCEMENT: Create a fresh, validated workbook rather than modifying the original
            self._update_progress(10, "Creating fresh workbook...")
            success = self._process_with_fresh_workbook(temp_copy, new_file, password)
            
            if success:
                self._update_progress(95, "Validating final workbook...")
                if self._validate_excel_file(new_file):
                    self._update_progress(100, "Processing complete")
                    self._update_status(f"Successfully created and processed: {new_file}")
                    if self.queue:
                        self.queue.put(("success", f"File processed successfully. Output saved to: {new_file}"))
                    return True
                else:
                    # If validation fails, try repair process
                    self._update_status("File validation failed, attempting repair...")
                    if self._repair_workbook(new_file):
                        self._update_status("File repaired successfully")
                        if self.queue:
                            self.queue.put(("success", f"File processed and repaired. Output saved to: {new_file}"))
                        return True
                    else:
                        raise ValueError("Failed to create a valid Excel file")
            else:
                # Fall back to legacy processing method
                self._update_status("Fresh workbook approach failed, falling back to original method...")
                return self._process_with_legacy_method(original_file, new_file, password)
                
        except Exception as e:
            self.logger.error("Processing failed", exc_info=True)
            if self.queue:
                self.queue.put(("error", str(e)))
            return False
        finally:
            # Final cleanup
            self._cleanup_temp_files()
    
    def _create_verified_copy(self, original_file: str) -> str:
        """
        Create a verified copy of the original file to prevent corruption.
        
        Args:
            original_file: Path to original Excel file
            
        Returns:
            str: Path to verified copy
        """
        temp_copy = self._get_temp_file_path("verified_copy")
        self._update_status(f"Creating verified copy at {temp_copy}...")
        
        # Use native Excel to create a clean copy (most reliable method)
        try:
            if WINDOWS_COM_AVAILABLE:
                # Start Excel
                pythoncom.CoInitialize()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                # Open workbook with recovery options
                wb = excel.Workbooks.Open(
                    original_file,
                    UpdateLinks=0,
                    ReadOnly=True,
                    CorruptLoad=2  # xlRepairFile (better corruption handling)
                )
                
                # Save as a new clean file
                wb.SaveAs(
                    temp_copy,
                    FileFormat=51,  # xlOpenXMLWorkbook
                    CreateBackup=False
                )
                
                # Clean close
                wb.Close(SaveChanges=False)
                excel.Quit()
                
                # Force cleanup
                del wb
                del excel
                gc.collect()
                pythoncom.CoUninitialize()
                
                # Verify file exists and has content
                if not os.path.exists(temp_copy) or os.path.getsize(temp_copy) == 0:
                    raise ValueError("Failed to create valid copy")
                    
                return temp_copy
            else:
                # If COM is not available, use regular copy
                shutil.copy2(original_file, temp_copy)
                return temp_copy
                
        except Exception as e:
            self.logger.warning(f"COM copy failed: {e}, falling back to direct copy")
            
            # Fallback to direct copy
            shutil.copy2(original_file, temp_copy)
            return temp_copy
    
    def _validate_excel_file(self, file_path: str) -> bool:
        """
        Validate an Excel file by trying to open it and check for errors.
        
        Args:
            file_path: Path to Excel file to validate
            
        Returns:
            bool: Whether the file is valid
        """
        try:
            # Method 1: Try opening with openpyxl
            if OPENPYXL_AVAILABLE:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                wb.close()
            
            # Method 2: Verify with Excel COM if possible
            if WINDOWS_COM_AVAILABLE:
                # Initialize COM
                pythoncom.CoInitialize()
                
                # Start Excel
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                try:
                    # Try to open the file
                    wb = excel.Workbooks.Open(
                        file_path,
                        UpdateLinks=0,
                        ReadOnly=True
                    )
                    
                    # Check if there are any error messages
                    has_errors = excel.ErrorCheckingStatus if hasattr(excel, 'ErrorCheckingStatus') else False
                    
                    # Close without saving
                    wb.Close(SaveChanges=False)
                    excel.Quit()
                    
                    # Cleanup COM objects
                    del wb
                    del excel
                    gc.collect()
                    pythoncom.CoUninitialize()
                    
                    if has_errors:
                        self.logger.warning(f"Excel detected errors in {file_path}")
                        return False
                        
                except Exception as e:
                    self.logger.warning(f"COM validation failed: {e}")
                    return False
                    
            # If we got here, file seems valid
            return True
            
        except Exception as e:
            self.logger.warning(f"File validation failed: {e}")
            return False
    
    def _process_with_fresh_workbook(self, source_file: str, output_file: str, password: str) -> bool:
        """
        Process by extracting data and creating a fresh workbook to avoid corruption.
        
        Args:
            source_file: Source Excel file
            output_file: Path for output file
            password: Sheet protection password
            
        Returns:
            bool: Whether processing succeeded
        """
        try:
            self._update_status("Loading source workbook data...")
            
            # Load source workbook
            if not OPENPYXL_AVAILABLE:
                return False
                
            source_wb = openpyxl.load_workbook(source_file, data_only=True)
            
            if self.sheet_name not in source_wb.sheetnames:
                self.logger.warning(f"Required sheet '{self.sheet_name}' not found")
                return False
                
            source_sheet = source_wb[self.sheet_name]
            
            # Create a new workbook
            self._update_status("Creating fresh workbook...")
            new_wb = openpyxl.Workbook()
            
            # Remove default sheet
            if "Sheet" in new_wb.sheetnames:
                new_wb.remove(new_wb["Sheet"])
                
            # Create new sheet with the right name
            new_sheet = new_wb.create_sheet(title=self.sheet_name)
            
            # Copy all data and formatting from source
            self._update_status("Copying data from source...")
            self._copy_sheet_data(source_sheet, new_sheet)
            
            # Process the fresh sheet
            self._update_progress(40, "Processing data...")
            
            # Process column visibility
            self._process_columns(new_sheet)
            
            # Find column indexes
            column_indexes = self._find_column_indexes(new_sheet)
            
            # Find header color
            rgb_color = self._process_header_formatting(new_sheet)
            
            # Find matching row
            matching_row = self._find_matching_row(new_sheet, rgb_color)
            
            # Add DO Comments column
            self._add_do_comments_column(new_sheet)
            
            # Process rows with comments
            self._process_rows_with_openpyxl(new_sheet, column_indexes, matching_row)
            
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # Save the workbook
            self._update_status("Saving processed workbook...")
            new_wb.save(output_file)
            
            # Cleanup
            source_wb.close()
            new_wb.close()
            
            # Verify the saved file
            if not os.path.exists(output_file):
                return False
                
            return True
            
        except Exception as e:
            self.logger.error(f"Error in fresh workbook processing: {e}", exc_info=True)
            return False
    
    def _copy_sheet_data(self, source_sheet, target_sheet):
        """
        Copy data and basic formatting from source sheet to target sheet.
        
        Args:
            source_sheet: Source worksheet
            target_sheet: Target worksheet
        """
        # Get dimensions of source sheet
        max_row = source_sheet.max_row
        max_col = source_sheet.max_column
        
        # Copy column dimensions and properties
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            if col_letter in source_sheet.column_dimensions:
                src_col_dim = source_sheet.column_dimensions[col_letter]
                tgt_col_dim = target_sheet.column_dimensions[col_letter]
                
                # Copy column width
                tgt_col_dim.width = src_col_dim.width if src_col_dim.width else 8.43  # Default width
                
                # Copy hidden status
                tgt_col_dim.hidden = False  # We want all columns visible
        
        # Copy row heights
        for row_idx in range(1, max_row + 1):
            if row_idx in source_sheet.row_dimensions:
                src_row_dim = source_sheet.row_dimensions[row_idx]
                tgt_row_dim = target_sheet.row_dimensions[row_idx]
                
                # Copy row height
                tgt_row_dim.height = src_row_dim.height if src_row_dim.height else 15  # Default height
        
        # Copy cell data and basic formatting
        for row_idx in range(1, max_row + 1):
            for col_idx in range(1, max_col + 1):
                # Get source cell
                src_cell = source_sheet.cell(row=row_idx, column=col_idx)
                
                # Skip merged cells (we'll handle them separately)
                if isinstance(src_cell, MergedCell):
                    continue
                    
                # Create target cell and copy value
                tgt_cell = target_sheet.cell(row=row_idx, column=col_idx, value=src_cell.value)
                
                # Copy basic formatting with proper error handling
                if hasattr(src_cell, 'has_style') and src_cell.has_style:
                    try:
                        # Font
                        tgt_cell.font = Font(
                            name=src_cell.font.name,
                            size=src_cell.font.size,
                            bold=src_cell.font.bold,
                            italic=src_cell.font.italic,
                            color=src_cell.font.color
                        )
                    except Exception as e:
                        self.logger.debug(f"Error copying font at {row_idx},{col_idx}: {e}")
                    
                    try:
                        # Fill - with proper Color object creation
                        if src_cell.fill and hasattr(src_cell.fill, 'start_color') and src_cell.fill.start_color:
                            fill_color = src_cell.fill.start_color.rgb or "FFFFFF"
                            # Create a proper Color object from the RGB string
                            color_obj = Color(rgb=fill_color)
                            tgt_cell.fill = PatternFill(
                                fill_type='solid',
                                start_color=color_obj
                            )
                    except Exception as e:
                        self.logger.debug(f"Error copying fill at {row_idx},{col_idx}: {e}")
                    
                    try:
                        # Border
                        if src_cell.border:
                            tgt_cell.border = Border(
                                left=src_cell.border.left,
                                right=src_cell.border.right,
                                top=src_cell.border.top,
                                bottom=src_cell.border.bottom
                            )
                    except Exception as e:
                        self.logger.debug(f"Error copying border at {row_idx},{col_idx}: {e}")
                    
                    try:
                        # Alignment
                        if src_cell.alignment:
                            tgt_cell.alignment = Alignment(
                                horizontal=src_cell.alignment.horizontal,
                                vertical=src_cell.alignment.vertical,
                                wrap_text=src_cell.alignment.wrap_text
                            )
                    except Exception as e:
                        self.logger.debug(f"Error copying alignment at {row_idx},{col_idx}: {e}")
    
    def _repair_workbook(self, file_path: str) -> bool:
        """
        Repair a potentially corrupted Excel file.
        
        Args:
            file_path: Path to Excel file to repair
            
        Returns:
            bool: Whether repair was successful
        """
        try:
            # Use the excel_recovery module if available
            try:
                from ..modules.excel_recovery import repair_excel_file_access
                success, result_path = repair_excel_file_access(file_path)
                if success and os.path.exists(result_path):
                    # Replace the original file with the repaired one
                    shutil.copy2(result_path, file_path)
                    return True
            except ImportError:
                self.logger.warning("Excel recovery module not available")
            
            # Method 1: Try pandas-based repair as fallback
            try:
                if PANDAS_AVAILABLE:
                    self._update_status("Attempting pandas-based repair...")
                    
                    # Read all sheets with pandas
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    
                    # Create a new Excel writer
                    repaired_path = self._get_temp_file_path("pandas_repaired")
                    writer = pd.ExcelWriter(repaired_path, engine='openpyxl')
                    
                    # Copy each sheet
                    for sheet_name in sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Save the writer
                    writer.close()
                    
                    # Replace original with repaired version
                    if os.path.exists(repaired_path) and os.path.getsize(repaired_path) > 0:
                        if os.path.exists(file_path):
                            os.unlink(file_path)
                        shutil.copy2(repaired_path, file_path)
                        return True
            except Exception as e:
                self.logger.warning(f"Pandas repair failed: {e}")
            
            # If all repair methods failed, return False
            return False
            
        except Exception as e:
            self.logger.error(f"Repair process failed: {e}")
            return False
    
    def _process_with_legacy_method(self, original_file: str, new_file: str, password: str = None) -> bool:
        """
        Process using the legacy method as a fallback.
        
        Args:
            original_file: Original file path
            new_file: New file path
            password: Sheet protection password
            
        Returns:
            bool: Whether processing was successful
        """
        try:
            # First try to process with pandas/openpyxl directly
            self._update_progress(5, "Attempting to process file with pandas/openpyxl...")
            try:
                success = self._process_with_libraries(original_file, new_file, password)
                if success:
                    self._update_progress(90, "Direct library processing successful")
                    self._update_status(f"Successfully created and processed: {new_file}")
                    if self.queue:
                        self.queue.put(("success", f"File processed successfully. Output saved to: {new_file}"))
                    return True
                else:
                    self._update_status("Direct library processing failed, falling back to basic copy...")
            except Exception as e:
                self.logger.warning(f"Direct library processing failed: {e}")
                self._update_status("Falling back to basic file copy...")
            
            # Very simple fallback - just make a copy
            try:
                self._update_progress(10, f"Creating new workbook at {new_file}...")
                shutil.copy2(original_file, new_file)
                
                self._update_progress(100, "Processing complete (basic copy only)")
                self._update_status(f"Successfully created copy: {new_file}")
                if self.queue:
                    self.queue.put(("success", f"File copied successfully. Output saved to: {new_file}"))
                return True
            except Exception as e:
                self.logger.error(f"Basic copy failed: {str(e)}", exc_info=True)
                if self.queue:
                    self.queue.put(("error", f"Processing failed: {str(e)}"))
                return False
                
        except Exception as e:
            self.logger.error("Processing failed", exc_info=True)
            if self.queue:
                self.queue.put(("error", str(e)))
            return False
    
    def _process_with_libraries(self, original_file: str, output_file: str, password: str) -> bool:
        """
        Process Excel file using Python libraries (openpyxl/pandas) instead of COM.
        
        Args:
            original_file: Path to original Excel file
            output_file: Path to output file
            password: Sheet protection password
            
        Returns:
            bool: True if processing was successful, False otherwise
        """
        if not OPENPYXL_AVAILABLE:
            self.logger.warning("openpyxl not available, cannot process with libraries")
            return False
            
        try:
            self._update_status(f"Loading file with pandas/openpyxl: {original_file}")
            
            # Create output directory if needed
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # Copy the file first for safety
            shutil.copy2(original_file, output_file)
            
            # Load workbook with openpyxl
            self._update_status("Loading workbook with openpyxl...")
            wb = openpyxl.load_workbook(output_file, data_only=True)
            
            if self.sheet_name not in wb.sheetnames:
                self.logger.warning(f"Required sheet '{self.sheet_name}' not found")
                return False
            
            sheet = wb[self.sheet_name]
            
            # Unprotect sheet if password provided
            if password:
                self._update_progress(20, "Unprotecting sheet...")
                self._unprotect_sheet(sheet, password)
            
            self._update_progress(30, "Processing columns...")
            self._process_columns(sheet)
            
            self._update_progress(60, "Finding column indexes...")
            column_indexes = self._find_column_indexes(sheet)
            
            self._update_progress(70, "Processing header formatting...")
            rgb_color = self._process_header_formatting(sheet)
            
            self._update_progress(80, "Finding matching rows...")
            matching_row = self._find_matching_row(sheet, rgb_color)
            
            self._update_progress(85, "Adding DO Comments column...")
            self._add_do_comments_column(sheet)
            
            self._update_progress(90, "Processing rows with comments...")
            self._process_rows_with_openpyxl(sheet, column_indexes, matching_row)
            
            # Save the workbook
            wb.save(output_file)
            wb.close()
            
            self._update_progress(100, "Processing complete")
            return True
                
        except Exception as e:
            self.logger.warning(f"Library-based processing failed: {str(e)}")
            return False
            
    def _process_columns(self, sheet) -> None:
        """
        Process and unhide all columns in worksheet.
        
        Args:
            sheet: Worksheet to process
        """
        for col in range(1, sheet.max_column + 1):
            col_letter = get_column_letter(col)
            sheet.column_dimensions[col_letter].hidden = False
    
    def _unprotect_sheet(self, sheet, password: str) -> None:
        """
        Unprotect worksheet with password.
        
        Args:
            sheet: Worksheet to unprotect
            password: Protection password
        """
        if password:
            try:
                if hasattr(sheet.protection, 'set_password'):
                    sheet.protection.set_password(password)
                sheet.protection.sheet = False
            except Exception as e:
                self.logger.warning(f"Failed to unprotect sheet: {e}")
    
    def _find_column_indexes(self, sheet) -> Dict[str, int]:
        """
        Find column indexes for required headers.
        
        Args:
            sheet: Worksheet to process
            
        Returns:
            Dict[str, int]: Mapping of header names to column indexes
        """
        column_indexes = {}
        
        # Set up default headers if not already specified
        headers_to_find = getattr(self, 'headers_to_find', DEFAULT_HEADERS_TO_FIND)
        header_row = getattr(self, 'header_row', DEFAULT_HEADER_ROW)
        
        # Find all headers
        for cell in sheet[header_row]:
            if cell.value in headers_to_find:
                column_indexes[cell.value] = cell.column
                
        # Check if we found all required headers
        if len(column_indexes) < len(headers_to_find):
            self.logger.warning(f"Not all required headers found. Found: {list(column_indexes.keys())}")
            
        return column_indexes
    
    def _process_header_formatting(self, sheet) -> str:
        """
        Process header cell formatting and return RGB color.
        
        Args:
            sheet: Worksheet to process
            
        Returns:
            str: RGB color value
        """
        header_row = getattr(self, 'header_row', DEFAULT_HEADER_ROW)
        header_cell = sheet.cell(row=header_row, column=1)
        
        # Default to white if no fill
        if not hasattr(header_cell, 'fill') or not header_cell.fill:
            return "FFFFFF"
            
        # Extract fill color
        fill_color = getattr(header_cell.fill.start_color, 'index', "FFFFFF")
        
        # Convert index to RGB if needed
        if isinstance(fill_color, int):
            fill_color = f"{fill_color:06X}"
            
        # Handle color index
        rgb_color = fill_color
        if hasattr(COLOR_INDEX, '__contains__') and int(fill_color, 16) in COLOR_INDEX:
            rgb_color = COLOR_INDEX[int(fill_color, 16)]
            
        # Clean up color format
        if isinstance(rgb_color, str) and len(rgb_color) == 8 and rgb_color.startswith('FF'):
            rgb_color = rgb_color[2:]
            
        return rgb_color
    
    def _find_matching_row(self, sheet, rgb_color: str) -> int:
        """
        Find first row matching header color.
        
        Args:
            sheet: Worksheet to process
            rgb_color: RGB color to match
            
        Returns:
            int: Matching row number
        """
        header_row = getattr(self, 'header_row', DEFAULT_HEADER_ROW)
        
        for row in sheet.iter_rows(min_row=header_row + 1):
            cell = row[0]
            
            # Skip cells without fill
            if not hasattr(cell, 'fill') or not cell.fill:
                continue
                
            # Get cell fill color
            cell_fill_color = getattr(cell.fill.start_color, 'index', None)
            
            # Skip cells without fill color
            if not cell_fill_color:
                continue
                
            # Convert to RGB
            if isinstance(cell_fill_color, int):
                cell_fill_color = f"{cell_fill_color:06X}"
                
            # Handle color index
            cell_rgb_color = cell_fill_color
            if hasattr(COLOR_INDEX, '__contains__') and int(cell_fill_color, 16) in COLOR_INDEX:
                cell_rgb_color = COLOR_INDEX[int(cell_fill_color, 16)]
                
            # Clean up color format
            if isinstance(cell_rgb_color, str) and len(cell_rgb_color) == 8 and cell_rgb_color.startswith('FF'):
                cell_rgb_color = cell_rgb_color[2:]
                
            # Compare colors
            if cell_rgb_color == rgb_color:
                return cell.row
                
        # If no matching row found, return max row
        return sheet.max_row
    
    def _add_do_comments_column(self, sheet) -> None:
        """
        Add and format DO Comments column.
        
        Args:
            sheet: Worksheet to process
        """
        last_col = sheet.max_column
        header_row = getattr(self, 'header_row', DEFAULT_HEADER_ROW)
        
        # Create the header cell
        new_header_cell = sheet.cell(row=header_row, column=last_col + 1)
        new_header_cell.value = "DO Comments"
        
        # Format the header cell
        new_header_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        new_header_cell.font = Font(color="FF0000", bold=True, size=11, name="Calibri")
        new_header_cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        new_header_cell.alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True
        )
        
        # Set column width
        col_letter = get_column_letter(last_col + 1)
        sheet.column_dimensions[col_letter].width = 25
    
    def _process_rows_with_openpyxl(self, sheet, column_indexes: Dict[str, int], matching_row: int) -> None:
        """
        Process individual rows with openpyxl.
        
        Args:
            sheet: openpyxl worksheet
            column_indexes: Column index mapping
            matching_row: Last row to process
        """
        last_col = sheet.max_column
        comment_col = last_col
        processed_count = 0
        header_row = getattr(self, 'header_row', DEFAULT_HEADER_ROW)
        
        # Check if we have required columns
        required_columns = ["Difference", "Include in CFO Cert Letter", "Explanation"]
        missing_columns = [col for col in required_columns if col not in column_indexes]
        
        if missing_columns:
            self.logger.warning(f"Missing required columns: {missing_columns}")
            return
            
        for row in range(header_row + 1, matching_row):
            try:
                difference_cell = sheet.cell(row=row, column=column_indexes["Difference"])
                include_cfo_cell = sheet.cell(row=row, column=column_indexes["Include in CFO Cert Letter"])
                explanation_cell = sheet.cell(row=row, column=column_indexes["Explanation"])
                
                difference_value = difference_cell.value
                include_cfo_value = include_cfo_cell.value
                explanation_value = explanation_cell.value
                
                # Add comment cell with appropriate formatting
                comment_cell = sheet.cell(row=row, column=comment_col)
                comment_cell.alignment = Alignment(wrap_text=True)
                
                if difference_value not in (None, ""):
                    if include_cfo_value == "N" and explanation_value not in (None, 0, ""):
                        comment_cell.value = "Explanation Reasonable"
                        processed_count += 1
                    elif include_cfo_value == "Y" and explanation_value not in (None, ""):
                        comment_cell.value = "Explanation Reasonable; Include in CFO Cert Letter"
                        processed_count += 1    
                    elif explanation_value in (None, "", 0) and difference_value != 0:
                        comment_cell.value = "Explanation Required"
                        # Add highlighting for cells that require attention
                        comment_cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                        processed_count += 1
                        
            except Exception as e:
                self._update_status(f"Error processing row {row}: {str(e)}")
        
        self._update_status(f"Successfully processed {processed_count} rows")
    
    def _validate_file(self, file_path: str) -> None:
        """
        Validate the input file path.
        
        Args:
            file_path: Path to Excel file
        
        Raises:
            ValueError: If file validation fails
        """
        if not os.path.exists(file_path):
            raise ValueError(f"File does not exist: {file_path}")
        if not file_path.lower().endswith('.xlsx'):
            raise ValueError("File must be an Excel (.xlsx) file")
    
    def _generate_new_filename(self, original_file: str) -> str:
        """
        Generate the new filename based on the original.
        
        Args:
            original_file: Original file path
        
        Returns:
            str: New file path
        """
        # Get original file basename without extension
        original_basename = Path(original_file).stem
        
        # Create output directory if it doesn't exist
        output_dir = Path(self.output_directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create a descriptive filename with timestamp
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        new_filename = f"{original_basename}_processed_{timestamp}.xlsx"
        
        # Construct full path
        new_path = output_dir / new_filename
        
        self._update_status(f"Generated output filename: {new_path}")
        return str(new_path.absolute())
    
    def _get_temp_file_path(self, prefix: str = "excel") -> str:
        """
        Generate a temporary file path.
        
        Args:
            prefix: Prefix for temporary file
            
        Returns:
            str: Temporary file path
        """
        temp_dir = tempfile.gettempdir()
        unique_id = str(uuid.uuid4())
        temp_file = os.path.join(temp_dir, f"{prefix}_temp_{unique_id}.xlsx")
        self._temp_files.append(temp_file)  # Track for cleanup
        return temp_file