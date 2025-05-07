"""
Core Excel processor for SF132/SF133 reconciliation.

This module contains the primary ExcelProcessor class that handles the main
Excel processing operations for SF132 to SF133 reconciliation.
"""

from pathlib import Path
import os
import time
import gc
import logging
from typing import Dict, Tuple, Optional, Any, List, Union
from dataclasses import dataclass
import queue
import shutil

import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment, Color
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

# Local imports
from ..modules.excel_handler import close_excel_instances
from ..modules.file_operations import (
    create_verified_copy,
    create_backup_file,
    get_temp_file_path,
    validate_excel_file
)
from ..modules.excel_cleaner import clean_excel_external_data
from ..config import (
    ProcessingConfig, 
    FileHandlingConfig, 
    ExcelConfig,
    app_config
)

# Type aliases
Queue = queue.Queue

logger = logging.getLogger(__name__)

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
        self.config = ProcessingConfig()
        self.file_config = FileHandlingConfig()
        self.excel_config = ExcelConfig()
        self._temp_files = []  # Track temp files for cleanup
        
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
                        logger.debug(f"Cleaned up temp file: {temp_file}")
                    except Exception as e:
                        logger.warning(f"Failed to clean up temp file {temp_file}: {e}")
    
    def _update_progress(self, value: float, message: str):
        """
        Update progress through the queue.
        
        Args:
            value: Progress value (0-100)
            message: Status message
        """
        if self.queue:
            self.queue.put(("progress", (value, message)))
        logger.info(message)
    
    def _update_status(self, message: str):
        """
        Send status update through the queue.
        
        Args:
            message: Status message
        """
        if self.queue:
            self.queue.put(("status", message))
        logger.info(message)
    
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
            close_excel_instances()
            
            # First make a clean copy of the original file
            temp_copy = create_verified_copy(original_file)
            self._temp_files.append(temp_copy)
            
            # ENHANCEMENT: Create a fresh, validated workbook rather than modifying the original
            self._update_progress(10, "Creating fresh workbook...")
            success = self._process_with_fresh_workbook(temp_copy, new_file, password)
            
            if success:
                self._update_progress(95, "Validating final workbook...")
                if validate_excel_file(new_file):
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
            logger.error("Processing failed", exc_info=True)
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
        output_dir = Path(self.config.output_directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Create a descriptive filename with timestamp
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        new_filename = f"{original_basename}_processed_{timestamp}.xlsx"
        
        # Construct full path
        new_path = output_dir / new_filename
        
        self._update_status(f"Generated output filename: {new_path}")
        return str(new_path.absolute())
    
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
            source_wb = openpyxl.load_workbook(source_file, data_only=True)
            
            if self.config.sheet_name not in source_wb.sheetnames:
                logger.warning(f"Required sheet '{self.config.sheet_name}' not found")
                return False
                
            source_sheet = source_wb[self.config.sheet_name]
            
            # Create a new workbook
            self._update_status("Creating fresh workbook...")
            new_wb = openpyxl.Workbook()
            
            # Remove default sheet
            if "Sheet" in new_wb.sheetnames:
                new_wb.remove(new_wb["Sheet"])
                
            # Create new sheet with the right name
            new_sheet = new_wb.create_sheet(title=self.config.sheet_name)
            
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
            logger.error(f"Error in fresh workbook processing: {e}", exc_info=True)
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
                if src_cell.has_style:
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
                        logger.debug(f"Error copying font at {row_idx},{col_idx}: {e}")
                    
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
                        logger.debug(f"Error copying fill at {row_idx},{col_idx}: {e}")
                    
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
                        logger.debug(f"Error copying border at {row_idx},{col_idx}: {e}")
                    
                    try:
                        # Alignment
                        if src_cell.alignment:
                            tgt_cell.alignment = Alignment(
                                horizontal=src_cell.alignment.horizontal,
                                vertical=src_cell.alignment.vertical,
                                wrap_text=src_cell.alignment.wrap_text
                            )
                    except Exception as e:
                        logger.debug(f"Error copying alignment at {row_idx},{col_idx}: {e}")
        
        # We'll deliberately NOT copy merged cells to avoid potential corruption
    
    def _repair_workbook(self, file_path: str) -> bool:
        """
        Repair a potentially corrupted Excel file.
        
        Args:
            file_path: Path to Excel file to repair
            
        Returns:
            bool: Whether repair was successful
        """
        try:
            self._update_status("Attempting to repair workbook...")
            
            # Create a backup before repair
            backup_path = create_backup_file(file_path, self.config.backup_directory)
            
            from ..modules.excel_recovery import repair_workbook
            repaired = repair_workbook(file_path)
            
            if not repaired:
                # If all repair methods failed, restore from backup
                self._update_status("Restoration from backup...")
                if os.path.exists(backup_path):
                    shutil.copy2(backup_path, file_path)
            
            return repaired
            
        except Exception as e:
            logger.error(f"Repair process failed: {e}")
            return False
    
    def _process_with_legacy_method(self, original_file: str, new_file: str, password: str) -> bool:
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
                    self._update_status("Direct library processing failed, falling back to COM methods...")
            except Exception as e:
                logger.warning(f"Direct library processing failed: {e}")
                self._update_status("Falling back to COM methods...")
            
            # Fall back to original COM processing approach
            try:
                from ..modules.com_operations import (
                    create_clean_copy,
                    process_workbook_with_com
                )
                
                # First make a copy of the original file - using more reliable file system operations
                self._update_progress(10, f"Creating new workbook at {new_file}...")
                if self.excel_config.use_com_for_copy:
                    # Try COM approach first, but fall back to direct copy if it fails
                    success = create_clean_copy(original_file, new_file)
                    if not success:
                        self._update_status("Falling back to direct file copy...")
                        self._direct_file_copy(original_file, new_file)
                else:
                    # Use direct file copy without COM (more reliable)
                    self._direct_file_copy(original_file, new_file)
                
                self._update_progress(20, "Processing Excel file...")
                process_workbook_with_com(new_file, password, self.config, self.update_progress, self.update_status)
                
                # Final verification and notification
                if os.path.exists(new_file):
                    self._update_status(f"Successfully created and processed: {new_file}")
                    if self.queue:
                        self.queue.put(("success", f"File processed successfully. Output saved to: {new_file}"))
                    return True
                else:
                    raise FileNotFoundError(f"Expected output file not found: {new_file}")
            except Exception as e:
                logger.error(f"COM processing failed: {str(e)}", exc_info=True)
                if self.queue:
                    self.queue.put(("error", f"Processing failed: {str(e)}"))
                return False
                
        except Exception as e:
            logger.error("Processing failed", exc_info=True)
            if self.queue:
                self.queue.put(("error", str(e)))
            return False

    def _direct_file_copy(self, source_file: str, dest_file: str) -> bool:
        """
        Perform a direct file copy without using COM.
        This is more reliable when COM operations fail.
        
        Args:
            source_file: Source file path
            dest_file: Destination file path
            
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
            logger.error(f"Error during direct file copy: {str(e)}", exc_info=True)
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
        try:
            self._update_status(f"Loading file with pandas/openpyxl: {original_file}")
            
            # Create output directory if needed
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # Copy the file first for safety
            shutil.copy2(original_file, output_file)
            
            # Load workbook with openpyxl
            self._update_status("Loading workbook with openpyxl...")
            wb = openpyxl.load_workbook(output_file, data_only=True)
            
            if self.config.sheet_name not in wb.sheetnames:
                logger.warning(f"Required sheet '{self.config.sheet_name}' not found")
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
            logger.warning(f"Library-based processing failed: {str(e)}")
            return False
    
    def _unprotect_sheet(self, sheet: openpyxl.worksheet.worksheet.Worksheet, password: str) -> None:
        """
        Unprotect worksheet with password.
        
        Args:
            sheet: Worksheet to unprotect
            password: Protection password
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
            sheet: Worksheet to process
        """
        for col in range(1, sheet.max_column + 1):
            col_letter = get_column_letter(col)
            sheet.column_dimensions[col_letter].hidden = False
    
    def _process_merged_cells(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """
        Unmerge all merged cells in worksheet.
        
        Args:
            sheet: Worksheet to process
        """
        merged_ranges = list(sheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            sheet.unmerge_cells(str(merged_range))
    
    def _find_column_indexes(self, sheet: openpyxl.worksheet.worksheet.Worksheet) -> Dict[str, int]:
        """
        Find column indexes for required headers.
        
        Args:
            sheet: Worksheet to process
            
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
            sheet: Worksheet to process
            
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
            sheet: Worksheet to process
            rgb_color: RGB color to match
            
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
            sheet: Worksheet to process
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
            sheet: openpyxl worksheet
            column_indexes: Column index mapping
            matching_row: Last row to process
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