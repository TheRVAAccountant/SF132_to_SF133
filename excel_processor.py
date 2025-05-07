import openpyxl
import win32com.client
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.cell import MergedCell
from openpyxl.utils import get_column_letter
import os
import time
import psutil
import logging
from typing import Dict, Tuple, Optional, Any, List
from dataclasses import dataclass
from queue import Queue
import shutil
from pathlib import Path
import tempfile
import uuid
import gc
import pythoncom
import contextlib
import subprocess
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException
from config import app_config
from excel_processor_config import ProcessingConfig, FileHandlingConfig, ExcelConfig

# Use the detailed configuration from excel_processor_config.py
DEFAULT_PROCESSING_CONFIG = ProcessingConfig()
DEFAULT_FILE_HANDLING_CONFIG = FileHandlingConfig()
DEFAULT_EXCEL_CONFIG = ExcelConfig()

class ExcelProcessor:
    """
    Handles Excel file processing operations including file manipulation,
    formatting, and content analysis.
    """
    
    def __init__(self, queue: Queue = None):
        """
        Initialize the Excel processor.
        
        Args:
            queue (Queue): Queue for communication with GUI
        """
        self.queue = queue
        self.config = DEFAULT_PROCESSING_CONFIG
        self.file_config = DEFAULT_FILE_HANDLING_CONFIG
        self.excel_config = DEFAULT_EXCEL_CONFIG
        self._setup_logging()
        self._temp_files = []  # Track temp files for cleanup
        
    def _setup_logging(self):
        """Configure logging for the processor."""
        self.logger = logging.getLogger(__name__)
    
    def __del__(self):
        """Clean up resources when instance is destroyed."""
        self._cleanup_temp_files()
        
    def _cleanup_temp_files(self):
        """Clean up all temporary files."""
        if app_config.cleanup_temp_files:
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
            value (float): Progress value (0-100)
            message (str): Status message
        """
        if self.queue:
            self.queue.put(("progress", (value, message)))
        self.logger.info(message)
    
    def _update_status(self, message: str):
        """
        Send status update through the queue.
        
        Args:
            message (str): Status message
        """
        if self.queue:
            self.queue.put(("status", message))
        self.logger.info(message)
    
    def close_excel_instances(self):
        """Terminate all existing Excel processes to prevent file locking."""
        self._update_status("Ensuring all Excel instances are closed...")
        excel_pids = []
        
        # First find all Excel processes
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
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

    def process_file(self, original_file: str, password: str = None) -> bool:
        """
        Main processing function for Excel file.
        
        Args:
            original_file (str): Path to original Excel file
            password (str): Sheet protection password
        
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
                    self._update_status("Direct library processing failed, falling back to COM methods...")
            except Exception as e:
                self.logger.warning(f"Direct library processing failed: {e}")
                self._update_status("Falling back to COM methods...")
            
            # Fall back to original COM processing approach
            try:
                # First make a copy of the original file - using more reliable file system operations
                self._update_progress(10, f"Creating new workbook at {new_file}...")
                if self.excel_config.use_com_for_copy:
                    # Try COM approach first, but fall back to direct copy if it fails
                    success = self._create_clean_copy(original_file, new_file)
                    if not success:
                        self._update_status("Falling back to direct file copy...")
                        self._direct_file_copy(original_file, new_file)
                else:
                    # Use direct file copy without COM (more reliable)
                    self._direct_file_copy(original_file, new_file)
                
                self._update_progress(20, "Processing Excel file...")
                self._process_workbook(new_file, password)
                
                # Final verification and notification
                if os.path.exists(new_file):
                    self._update_status(f"Successfully created and processed: {new_file}")
                    if self.queue:
                        self.queue.put(("success", f"File processed successfully. Output saved to: {new_file}"))
                    return True
                else:
                    raise FileNotFoundError(f"Expected output file not found: {new_file}")
            except Exception as e:
                self.logger.error(f"COM processing failed: {str(e)}")
                if self.queue:
                    self.queue.put(("error", f"Processing failed: {str(e)}"))
                return False
                
        except Exception as e:
            self.logger.error("Processing failed", exc_info=True)
            if self.queue:
                self.queue.put(("error", str(e)))
            return False
        finally:
            # Final cleanup
            self._cleanup_temp_files()
    
    def _validate_file(self, file_path: str) -> None:
        """
        Validate the input file path.
        
        Args:
            file_path (str): Path to Excel file
        
        Raises:
            ValueError: If file validation fails
        """
        if not os.path.exists(file_path):
            raise ValueError(f"File does not exist: {file_path}")
        if not file_path.lower().endswith('.xlsx'):
            raise ValueError("File must be an Excel (.xlsx) file")
    
    def _process_with_libraries(self, original_file: str, output_file: str, password: str) -> bool:
        """
        Process Excel file using Python libraries (openpyxl/pandas) instead of COM.
        
        Args:
            original_file (str): Path to original Excel file
            output_file (str): Path to output file
            password (str): Sheet protection password
            
        Returns:
            bool: True if processing was successful, False otherwise
        """
        try:
            self._update_status(f"Loading file with pandas/openpyxl: {original_file}")
            
            # First try with pandas
            try:
                # Create output directory if needed
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
                
                # Copy the file first for safety
                shutil.copy2(original_file, output_file)
                
                # Load workbook with openpyxl
                self._update_status("Loading workbook with openpyxl...")
                wb = openpyxl.load_workbook(output_file, data_only=True)
                
                if self.config.sheet_name not in wb.sheetnames:
                    self.logger.warning(f"Required sheet '{self.config.sheet_name}' not found")
                    return False
                
                sheet = wb[self.config.sheet_name]
                
                # Unprotect sheet if password provided
                if password:
                    self._update_progress(20, "Unprotecting sheet...")
                    self._unprotect_sheet(sheet, password)
                
                self._update_progress(30, "Processing columns...")
                self._process_columns(sheet)
                
                self._update_progress(50, "Processing merged cells...")
                self._process_merged_cells(sheet)
                
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
                self.logger.warning(f"Openpyxl processing failed: {str(e)}")
                return False
                
        except Exception as e:
            self.logger.warning(f"Library-based processing failed: {str(e)}")
            return False
        
    def _generate_new_filename(self, original_file: str) -> str:
        """
        Generate the new filename based on the original.
        
        Args:
            original_file (str): Original file path
        
        Returns:
            str: New file path
        """
        # Get original file basename without extension
        original_basename = Path(original_file).stem
        
        # Create output directory if it doesn't exist
        output_dir = Path(self.config.output_directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create a descriptive filename with timestamp
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        new_filename = f"{original_basename}_processed_{timestamp}.xlsx"
        
        # Construct full path
        new_path = output_dir / new_filename
        
        self._update_status(f"Generated output filename: {new_path}")
        return str(new_path.absolute())
    
    def _direct_file_copy(self, source_file: str, dest_file: str) -> bool:
        """
        Perform a direct file copy without using COM.
        This is more reliable when COM operations fail.
        
        Args:
            source_file (str): Source file path
            dest_file (str): Destination file path
            
        Returns:
            bool: Whether the copy was successful
        """
        try:
            self._update_status(f"Directly copying file from {source_file} to {dest_file}...")
            
            # Ensure the destination directory exists
            os.makedirs(os.path.dirname(dest_file), exist_ok=True)
            
            # Make a direct copy of the file
            shutil.copy2(source_file, dest_file)
            
            # Verify the copy succeeded
            if not os.path.exists(dest_file):
                raise IOError(f"Failed to copy file to {dest_file}")
            
            file_size = os.path.getsize(dest_file)
            if file_size == 0:
                raise IOError(f"Copied file has zero size: {dest_file}")
                
            self._update_status(f"Direct file copy successful ({file_size} bytes)")
            return True
            
        except Exception as e:
            self.logger.error(f"Error during direct file copy: {str(e)}", exc_info=True)
            return False

    def _verify_loaded_workbook(self, wb: openpyxl.Workbook, file_path: str) -> None:
        """
        Verify that a loaded workbook has the required sheet.
        
        Args:
            wb (openpyxl.Workbook): Workbook to verify
            file_path (str): Path to Excel file
            
        Raises:
            ValueError: If the workbook does not have the required sheet
        """
        if self.config.sheet_name not in wb.sheetnames:
            self.logger.error(f"Sheet '{self.config.sheet_name}' not found in {file_path}")
            self.logger.info(f"Available sheets: {wb.sheetnames}")
            raise ValueError(f"Required sheet '{self.config.sheet_name}' not found in workbook. "
                            f"Available sheets: {wb.sheetnames}")
        
        # Verify that the sheet has content
        sheet = wb[self.config.sheet_name]
        if sheet.max_row < 1 or sheet.max_column < 1:
            raise ValueError(f"Sheet '{self.config.sheet_name}' appears to be empty")

    def _create_clean_copy(self, original_file: str, new_file: str) -> bool:
        """
        Create a clean copy of the original file using COM if possible.
        
        Args:
            original_file (str): Original file path
            new_file (str): New file path
            
        Returns:
            bool: Whether the operation was successful
        """
        try:
            # First make sure the output directory exists
            os.makedirs(os.path.dirname(new_file), exist_ok=True)
            
            # Create a temporary file for intermediate processing
            temp_file = self._get_temp_file_path("init_copy")
            
            # Try opening the original file to verify its integrity
            self._update_status("Verifying original file integrity...")
            try:
                # Use openpyxl to verify file can be opened
                with self._open_and_close_workbook(original_file) as orig_wb:
                    if orig_wb is None:
                        self.logger.warning("Original file couldn't be verified with openpyxl, "
                                           "will attempt to fix through copy process")
            except Exception as e:
                self.logger.warning(f"Original file verification failed: {e}")
                
            # Do a standard file copy first to get a basic working file
            self._update_status("Creating initial file copy...")
            shutil.copy2(original_file, temp_file)
            
            # Now use Excel COM to save the file properly
            self._update_status("Creating clean Excel copy...")
            return self._save_with_excel_com(temp_file, new_file, preserve_all=True)
            
        except Exception as e:
            self._update_status(f"Error creating clean copy: {str(e)}")
            return False

    def _save_with_excel_com(self, source_file: str, dest_file: str, preserve_all: bool = False) -> bool:
        """
        Use Excel COM to save a file, with robust fallback mechanisms.
        
        Args:
            source_file (str): Source file path
            dest_file (str): Destination file path
            preserve_all (bool): Whether to preserve all Excel features
            
        Returns:
            bool: Whether the save was successful
        """
        if not self.excel_config.enable_com:
            self._update_status("COM operations disabled in configuration")
            return False
            
        # Check if file exists before attempting COM
        if not os.path.exists(source_file):
            self.logger.error(f"Source file does not exist: {source_file}")
            return False
            
        # Make sure Excel is not running to avoid conflicts
        self.close_excel_instances()
        time.sleep(1)  # Give OS time to fully release resources
        
        excel = None
        wb = None
        success = False
        retry_count = 0
        max_retries = self.excel_config.max_com_retries
        
        while retry_count <= max_retries and not success:
            try:
                # Add COM security settings
                try:
                    # Set default security for automation
                    import win32com
                    import win32com.client.gencache
                    win32com.client.gencache.EnsureDispatch("Excel.Application")
                except:
                    pass
                
                # Initialize COM with proper threading model
                pythoncom.CoInitialize()
                
                # Start Excel with proper security context
                self._update_status(f"Starting Excel (attempt {retry_count+1}/{max_retries+1})...")
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.AskToUpdateLinks = False
                excel.EnableEvents = False  # Disable Excel events
                
                # Open workbook with safe options
                self._update_status(f"Opening file with Excel COM: {source_file}")
                
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
                    self._clean_external_connections(wb, excel)
                
                # Save to the destination with correct format
                self._update_status(f"Saving clean copy to: {dest_file}")
                
                # Use explicit constant for better clarity
                XL_XLSX = 51  # Excel.XlFileFormat.xlOpenXMLWorkbook
                
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
                    self._update_status(f"Excel COM save operation successful")
                else:
                    raise IOError(f"Failed to save file to {dest_file}")
                
            except Exception as e:
                retry_count += 1
                self.logger.warning(f"COM attempt {retry_count}/{max_retries+1} failed: {str(e)}")
                
                if retry_count <= max_retries:
                    self._update_status(f"Retrying COM operation (attempt {retry_count+1}/{max_retries+1})...")
                    time.sleep(2)  # Wait before retrying
                    # Make sure Excel is fully closed
                    self.close_excel_instances()
                else:
                    self.logger.error(f"All COM attempts failed: {str(e)}")
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

    def _clean_external_connections(self, wb, excel) -> None:
        """
        Clean up external data connections and links in the workbook.
        
        Args:
            wb: Excel workbook COM object
            excel: Excel application COM object
        """
        try:
            # Break links to external sources
            self._update_status("Breaking links to external sources...")
            
            # Try to break external links
            try:
                wb.BreakLink(Name="", Type=1)  # Type 1 = xlExcelLinks
            except:
                pass
                
            # Remove external data ranges/connections
            try:
                # Check if there are connections to remove
                if hasattr(wb, 'Connections') and wb.Connections.Count > 0:
                    self._update_status(f"Removing {wb.Connections.Count} external connections...")
                    
                    # Get all connection names first (because removing changes the collection)
                    conn_names = []
                    for i in range(1, wb.Connections.Count + 1):
                        try:
                            conn_names.append(wb.Connections(i).Name)
                        except:
                            pass
                    
                    # Now remove each connection
                    for name in conn_names:
                        try:
                            wb.Connections(name).Delete()
                        except Exception as e:
                            self.logger.warning(f"Failed to remove connection {name}: {e}")
            except Exception as e:
                self.logger.warning(f"Error cleaning connections: {e}")
                
            # Clean up data ranges
            try:
                for sheet in wb.Sheets:
                    try:
                        # Check if the sheet has QueryTables
                        if hasattr(sheet, 'QueryTables') and sheet.QueryTables.Count > 0:
                            self._update_status(f"Removing query tables from sheet: {sheet.Name}")
                            
                            # Delete all QueryTables
                            for i in range(sheet.QueryTables.Count, 0, -1):
                                try:
                                    sheet.QueryTables(i).Delete()
                                except:
                                    pass
                    except:
                        pass
                    
                    # Also handle ListObjects (tables) with external data
                    try:
                        if hasattr(sheet, 'ListObjects') and sheet.ListObjects.Count > 0:
                            for i in range(sheet.ListObjects.Count, 0, -1):
                                try:
                                    table = sheet.ListObjects(i)
                                    if hasattr(table, 'QueryTable'):
                                        table.QueryTable.Delete()
                                except:
                                    pass
                    except:
                        pass
            except Exception as e:
                self.logger.warning(f"Error cleaning data ranges: {e}")
                
        except Exception as e:
            self.logger.warning(f"Error during connection cleanup: {e}")

    def _process_workbook(self, file_path: str, password: str) -> None:
        """
        Process the Excel workbook.
        
        Args:
            file_path (str): Path to Excel file
            password (str): Sheet protection password
        """
        try:
            # Load workbook with proper error handling
            self._update_status(f"Loading workbook: {file_path}")
            wb = self._load_workbook_safely(file_path)
            
            sheet = self._get_worksheet(wb)
            
            self._update_progress(20, "Unprotecting sheet...")
            self._unprotect_sheet(sheet, password)
            
            self._update_progress(30, "Processing columns...")
            self._process_columns(sheet)
            
            self._update_progress(50, "Processing merged cells...")
            self._process_merged_cells(sheet)
            
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
            
            # First save with openpyxl for the cell content changes
            self._update_status("Initial save with content changes...")
            self._save_workbook_safely(wb, file_path)
            
            # Then perform a special "clean" save to fix any potential corruption
            self._update_status("Cleaning and finalizing workbook...")
            self._final_clean_save(file_path)
            
            self._update_progress(100, "Processing complete")
        except Exception as e:
            self.logger.error(f"Error processing workbook: {str(e)}", exc_info=True)
            if self.queue:
                self.queue.put(("error", f"Workbook processing failed: {str(e)}"))
            raise
    
    def _final_clean_save(self, file_path: str) -> None:
        """
        Perform a final clean save using the most reliable approach.
        
        Args:
            file_path (str): Path to Excel file
        """
        # Create a backup of the current file state
        backup_path = self._create_backup_file(file_path)
        self._update_status(f"Created backup before final save: {backup_path}")
        
        # Get a new temporary filename for the cleaned version
        clean_file = self._get_temp_file_path("clean")
        
        # Try using Excel COM to save a clean version if enabled
        if self.excel_config.use_com_for_final_save:
            success = self._save_with_excel_com(file_path, clean_file, preserve_all=False)
        else:
            success = False
            
        if success and os.path.exists(clean_file):
            # Replace the original with the cleaned version
            if os.path.exists(file_path):
                os.unlink(file_path)  # Remove existing file
            shutil.move(clean_file, file_path)
            self._update_status("Final clean save completed successfully")
        else:
            # If COM approach failed, ensure we still have a working file
            self._update_status("Using openpyxl for final save...")
            
            try:
                # Load and save with openpyxl as fallback
                wb = openpyxl.load_workbook(file_path)
                wb.save(clean_file)
                wb.close()
                
                if os.path.exists(clean_file) and os.path.getsize(clean_file) > 0:
                    if os.path.exists(file_path):
                        os.unlink(file_path)
                    shutil.move(clean_file, file_path)
                    self._update_status("Final openpyxl save completed successfully")
                else:
                    raise IOError("Failed to save with openpyxl")
                    
            except Exception as e:
                self._update_status(f"Final save with openpyxl failed: {str(e)}")
                self._update_status("Keeping original processed file")
                if self.queue:
                    self.queue.put(("warning", 
                        "The enhanced clean-up process failed, but the file was successfully processed. "
                        "The file should still be usable."
                    ))
                
        # Final verification of output file
        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            self._update_status(f"Final file size: {os.path.getsize(file_path)} bytes")
        else:
            # If we somehow lost the file, restore from backup
            self._update_status("Output file missing or corrupted, restoring from backup...")
            shutil.copy2(backup_path, file_path)
            
    def _load_workbook_safely(self, file_path: str) -> openpyxl.Workbook:
        """
        Load workbook with enhanced error handling.
        
        Args:
            file_path (str): Path to Excel file
            
        Returns:
            openpyxl.Workbook: Loaded workbook
        """
        try:
            # Try loading with pandas first
            try:
                self._update_status("Attempting to load workbook with pandas...")
                df = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
                self._update_status(f"Successfully read with pandas: {len(df)} sheets")
            except Exception as e:
                self.logger.warning(f"Pandas loading failed, falling back to openpyxl: {str(e)}")
            
            # Make sure file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
                
            # Try to load with openpyxl's data_only mode
            self._update_status(f"Loading workbook with openpyxl: {file_path}")
            retry_count = 0
            max_retries = self.file_config.max_retries
            last_error = None
            
            while retry_count <= max_retries:
                try:
                    # Load with data_only=True to get values instead of formulas
                    wb = openpyxl.load_workbook(file_path, data_only=True)
                    self._update_status(f"Successfully loaded workbook with {len(wb.sheetnames)} sheets")
                    return wb
                except Exception as e:
                    retry_count += 1
                    last_error = e
                    self.logger.warning(f"Attempt {retry_count} to load workbook failed: {str(e)}")
                    
                    if retry_count <= max_retries:
                        # Wait before retrying
                        time.sleep(self.file_config.retry_delay)
                        # Make sure Excel is closed
                        self.close_excel_instances()
                    else:
                        break
                        
            # If all openpyxl attempts failed, try creating a safe copy first
            self._update_status("Creating safe copy for loading...")
            safe_copy = self._get_temp_file_path("safe_copy")
            
            # Copy the file
            shutil.copy2(file_path, safe_copy)
            
            # Try loading the safe copy
            try:
                wb = openpyxl.load_workbook(safe_copy, data_only=True)
                self._update_status("Successfully loaded workbook from safe copy")
                return wb
            except Exception as final_error:
                self.logger.error(f"Failed to load workbook after all attempts: {str(final_error)}")
                raise ValueError(f"Could not load Excel file after multiple attempts: {str(final_error)}")
            
        except FileNotFoundError:
            raise  # Re-raise file not found errors as is
        except Exception as e:
            self.logger.error(f"Failed to load workbook: {str(e)}", exc_info=True)
            raise ValueError(f"Could not load Excel file: {str(e)}")
    
    def _save_workbook_safely(self, wb: openpyxl.Workbook, file_path: str) -> None:
        """
        Save workbook with enhanced error handling and backup.
        
        Args:
            wb (openpyxl.Workbook): Workbook to save
            file_path (str): Path to save to
        """
        # Create a backup before saving
        backup_path = self._create_backup_file(file_path)
        self._update_status(f"Created backup at: {backup_path}")
        
        # Multi-stage save process to prevent corruption
        stage1_file = self._get_temp_file_path("stage1")
        
        try:
            # STAGE 1: Save to first temp file
            self._update_status("Stage 1: Initial save to temp file...")
            wb.save(stage1_file)
            
            # Verify stage 1 file was created and is valid
            if not os.path.exists(stage1_file):
                raise IOError("Failed to save to first-stage temporary file")
            
            # Test open the temp file to verify integrity
            with self._open_and_close_workbook(stage1_file) as test_wb:
                if test_wb is None:
                    raise ValueError("Failed to verify workbook structure")
            
            # Close the original workbook to release resources
            wb.close()
            
            # Give system time to release file handles
            time.sleep(0.5)
            gc.collect()  # Force garbage collection
            
            # Move to final destination
            if os.path.exists(file_path):
                os.unlink(file_path)  # Remove existing file
             
            # Copy instead of move to ensure file system caching doesn't cause issues
            shutil.copy2(stage1_file, file_path)
            
            # Verify the final copy succeeded
            if not os.path.exists(file_path):
                raise IOError(f"Failed to copy final file to destination: {file_path}")
                
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                raise IOError(f"Saved file has zero size: {file_path}")
                
            self._update_status(f"Saved file successfully ({file_size} bytes) to: {file_path}")
        except Exception as e:
            self.logger.error(f"Error during save: {str(e)}", exc_info=True)
            if os.path.exists(backup_path):
                self._update_status("Restoring from backup due to save error...")
                shutil.copy2(backup_path, file_path)
            raise ValueError(f"Failed to save workbook: {str(e)}")
            
        finally:
            # Clean up temp files
            if os.path.exists(stage1_file):
                try:
                    os.unlink(stage1_file)
                except:
                    # Add to cleanup list for later
                    self._temp_files.append(stage1_file)

    def _get_temp_file_path(self, prefix: str = "excel") -> str:
        """
        Generate a temporary file path.
        
        Args:
            prefix (str): Prefix for temporary file
            
        Returns:
            str: Temporary file path
        """
        temp_dir = self.file_config.get_temp_dir() or tempfile.gettempdir()
        unique_id = str(uuid.uuid4())
        temp_file = os.path.join(temp_dir, f"{prefix}_temp_{unique_id}.xlsx")
        self._temp_files.append(temp_file)  # Track for cleanup
        return temp_file
    
    def _create_backup_file(self, file_path: str) -> str:
        """
        Create a backup copy of the file.
        
        Args:
            file_path (str): Path to file to backup
            
        Returns:
            str: Path to backup file
        """
        backup_dir = Path(self.config.output_directory) / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        filename = Path(file_path).name
        backup_path = str(backup_dir / f"{Path(filename).stem}_backup_{timestamp}{Path(filename).suffix}")
        shutil.copy2(file_path, backup_path)
        return backup_path
    
    @contextlib.contextmanager
    def _open_and_close_workbook(self, file_path: str):
        """
        Context manager to safely open and close a workbook.
        
        Args:
            file_path (str): Path to Excel file
            
        Yields:
            openpyxl.Workbook or None: Opened workbook or None if failed
        """
        wb = None
        try:
            # Wait a moment for filesystem to settle
            time.sleep(0.2)
            wb = openpyxl.load_workbook(file_path, read_only=True)
            yield wb
        except Exception as e:
            self.logger.warning(f"Failed to open workbook {file_path}: {e}")
            yield None
        finally:
            if wb:
                try:
                    wb.close()
                except:
                    pass
    
    def _get_worksheet(self, workbook: openpyxl.Workbook) -> openpyxl.worksheet.worksheet.Worksheet:
        """
        Get the target worksheet from workbook.
        
        Args:
            workbook (Workbook): Openpyxl workbook object
            
        Returns:
            Worksheet: Target worksheet
            
        Raises:
            ValueError: If worksheet not found
        """
        if self.config.sheet_name not in workbook.sheetnames:
            raise ValueError(f"Sheet '{self.config.sheet_name}' not found")
        return workbook[self.config.sheet_name]
    
    def _unprotect_sheet(self, sheet: openpyxl.worksheet.worksheet.Worksheet, password: str) -> None:
        """
        Unprotect worksheet with password.
        
        Args:
            sheet (Worksheet): Worksheet to unprotect
            password (str): Protection password
        """
        if password:
            try:
                sheet.protection.set_password(password)
                sheet.protection.sheet = False
            except Exception as e:
                raise ValueError(f"Failed to unprotect sheet: {str(e)}")
    
    def _process_columns(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """
        Process and unhide all columns in worksheet.
        
        Args:
            sheet (Worksheet): Worksheet to process
        """
        for col in sheet.columns:
            cell = col[0]
            if not isinstance(cell, MergedCell):
                col_letter = get_column_letter(cell.column)
                sheet.column_dimensions[col_letter].hidden = False
    
    def _process_merged_cells(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """
        Unmerge all merged cells in worksheet.
        
        Args:
            sheet (Worksheet): Worksheet to process
        """
        merged_ranges = list(sheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            sheet.unmerge_cells(str(merged_range))
    
    def _find_column_indexes(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> Dict[str, int]:
        """
        Find column indexes for required headers.
        
        Args:
            sheet (Worksheet): Worksheet to process
            
        Returns:
            Dict[str, int]: Mapping of header names to column indexes
        """
        column_indexes = {}
        for cell in sheet[self.config.header_row]:
            if cell.value in self.config.headers_to_find:
                column_indexes[cell.value] = cell.column
        missing_headers = set(self.config.headers_to_find) - set(column_indexes.keys())
        if missing_headers:
            raise ValueError(f"Missing headers: {', '.join(missing_headers)}")
        return column_indexes
    
    def _process_header_formatting(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> str:
        """
        Process header cell formatting and return RGB color.
        
        Args:
            sheet (Worksheet): Worksheet to process
            
        Returns:
            str: RGB color value
        """
        header_cell = sheet.cell(row=self.config.header_row, column=1)
        fill_color = header_cell.fill.start_color.index
        
        if isinstance(fill_color, int):
            fill_color = f"{fill_color:06X}"
        if fill_color in COLOR_INDEX:
            rgb_color = COLOR_INDEX[int(fill_color, 16)]
        else:
            rgb_color = fill_color
        if not rgb_color.startswith('FF') and len(rgb_color) == 8:
            rgb_color = rgb_color[2:]
        return rgb_color
    
    def _find_matching_row(self, sheet: openpyxl.worksheet.worksheet.Worksheet, rgb_color: str) -> int:
        """
        Find first row matching header color.
        
        Args:
            sheet (Worksheet): Worksheet to process
            rgb_color (str): RGB color to match
            
        Returns:
            int: Matching row number
        """
        for row in sheet.iter_rows(min_row=self.config.header_row + 1):
            cell = row[0]
            cell_fill_color = cell.fill.start_color.index
            
            if isinstance(cell_fill_color, int):
                cell_fill_color = f"{cell_fill_color:06X}"
            if cell_fill_color in COLOR_INDEX:
                cell_rgb_color = COLOR_INDEX[int(cell_fill_color, 16)]
            else:
                cell_rgb_color = cell_fill_color
            if not cell_rgb_color.startswith('FF') and len(cell_rgb_color) == 8:
                cell_rgb_color = cell_rgb_color[2:]
            if cell_rgb_color == rgb_color:
                return cell.row
        return sheet.max_row
    
    def _add_do_comments_column(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """
        Add and format DO Comments column.
        
        Args:
            sheet (Worksheet): Worksheet to process
        """
        last_col = sheet.max_column
        new_header_cell = sheet.cell(row=self.config.header_row, column=last_col + 1)
        new_header_cell.value = "DO Comments"
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
        col_letter = get_column_letter(last_col + 1)
        sheet.column_dimensions[col_letter].width = 25
    
    def _process_rows_with_openpyxl(
        self,
        sheet: openpyxl.worksheet.worksheet.Worksheet,
        column_indexes: Dict[str, int],
        matching_row: int
    ) -> None:
        """
        Process individual rows with openpyxl.
        
        Args:
            sheet (Worksheet): openpyxl worksheet
            column_indexes (Dict[str, int]): Column index mapping
            matching_row (int): Last row to process
        """
        last_col = sheet.max_column
        comment_col = last_col
        processed_count = 0
        
        for row in range(self.config.header_row + 1, matching_row):
            try:
                difference_cell = sheet.cell(row=row, column=column_indexes["Difference"])
                include_cfo_cell = sheet.cell(row=row, column=column_indexes["Include in CFO Cert Letter"])
                explanation_cell = sheet.cell(row=row, column=column_indexes["Explanation"])
                
                difference_value = difference_cell.value
                include_cfo_value = include_cfo_cell.value
                explanation_value = explanation_cell.value
                
                if difference_value not in (None, ""):
                    if include_cfo_value == "N" and explanation_value not in (None, 0, ""):
                        sheet.cell(row=row, column=comment_col).value = "Explanation Reasonable"
                        processed_count += 1
                    elif include_cfo_value == "Y" and explanation_value not in (None, ""):
                        sheet.cell(row=row, column=comment_col).value = "Explanation Reasonable; Include in CFO Cert Letter"
                        processed_count += 1    
                    elif explanation_value in (None, "", 0) and difference_value != 0:
                        sheet.cell(row=row, column=comment_col).value = "Explanation Required"
                        processed_count += 1
                        
            except Exception as e:
                self._update_status(f"Error processing row {row}: {str(e)}")
        
        self._update_status(f"Successfully processed {processed_count} rows")
    
    def _process_with_pywin32(self, file_path: str, column_indexes: Dict[str, int], matching_row: int) -> None:
        """
        Process workbook with pywin32 for formula evaluation.
        
        Args:
            file_path (str): Path to Excel file
            column_indexes (Dict[str, int]): Column index mappings
            matching_row (int): Last row to process
        """
        file_path = os.path.abspath(file_path)
        
        self._update_status(f"Checking if file exists: {file_path}")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File does not exist: {file_path}")
        
        self.close_excel_instances()
        time.sleep(2)
        
        excel = None
        wb = None
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            self._update_status("Starting Excel application...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            self._update_status(f"Opening workbook: {file_path}")
            wb = excel.Workbooks.Open(file_path)    
            
            self._update_status("Selecting worksheet...")
            ws = wb.Worksheets(self.config.sheet_name)
                
            self._process_rows_with_pywin32(
                ws,
                column_indexes,
                matching_row
            ) 
            
            self._update_status("Saving workbook...")
            wb.Save()    
            
        except Exception as e:
            self._update_status(f"Error in Excel processing: {str(e)}")
            import traceback
            self._update_status(f"Detailed error: {traceback.format_exc()}")
            raise
        finally:
            # Clean up COM objects
            if wb:
                try:
                    wb.Close(SaveChanges=True)
                except Exception as e:
                    self._update_status(f"Error closing workbook: {str(e)}")
            
            if excel:
                try:
                    excel.Quit()
                except Exception as e:
                    self._update_status(f"Error quitting Excel: {str(e)}")
                
                del wb
                del excel
                
            # Force garbage collection to release COM objects
            gc.collect()
            pythoncom.CoUninitialize()
    
    def _process_rows_with_pywin32(self, ws, column_indexes: Dict[str, int], matching_row: int) -> None:
        """
        Process individual rows with pywin32.
        
        Args:
            ws: pywin32 worksheet object
            column_indexes (Dict[str, int]): Column index mapping
            matching_row (int): Last row to process
        """
        last_col = ws.UsedRange.Columns.Count
        processed_count = 0
        
        for row in range(self.config.header_row + 1, matching_row):
            try:
                difference_value = ws.Cells(row, column_indexes["Difference"]).Value
                include_cfo_value = ws.Cells(row, column_indexes["Include in CFO Cert Letter"]).Value
                explanation_value = ws.Cells(row, column_indexes["Explanation"]).Value
                
                if difference_value not in (None, ""):
                    if include_cfo_value == "N" and explanation_value not in (None, 0, ""):
                        ws.Cells(row, last_col).Value = "Explanation Reasonable"
                        processed_count += 1
                    elif include_cfo_value == "Y" and explanation_value not in (None, ""):
                        ws.Cells(row, last_col).Value = "Explanation Reasonable; Include in CFO Cert Letter"
                        processed_count += 1
                    elif explanation_value in (None, "", 0) and difference_value != 0:
                        ws.Cells(row, last_col).Value = "Explanation Required"
                        processed_count += 1
                        
            except Exception as e:
                self._update_status(f"Error processing row {row}: {str(e)}")
        
        self._update_status(f"Successfully processed {processed_count} rows")