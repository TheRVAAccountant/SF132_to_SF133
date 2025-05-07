"""
Specialized module for cleaning Excel external data references.
This module offers comprehensive solutions for removing external connections and data ranges
from Excel files using both COM automation and openpyxl approaches.
"""

import logging
import os
import win32com.client
import pythoncom
import openpyxl
from openpyxl.workbook.external_reference import ExternalReference
import time
from typing import List, Dict, Any, Optional, Set

class ExcelDataCleaner:
    """
    Specialized class for cleaning external data references in Excel files.
    """
    
    def __init__(self, logger=None):
        """
        Initialize the Excel data cleaner.
        
        Args:
            logger: Optional logger instance
        """
        self.logger = logger or logging.getLogger(__name__)
    
    def clean_external_references_openpyxl(self, workbook: openpyxl.Workbook) -> None:
        """
        Remove all external data references using openpyxl.
        
        Args:
            workbook: openpyxl Workbook object
        """
        try:
            # Clean external links/references
            if hasattr(workbook, '_external_links'):
                self.logger.info(f"Removing {len(workbook._external_links)} external links")
                workbook._external_links = []
            
            # Clean defined names with external references
            names_to_remove = []
            if hasattr(workbook, 'defined_names'):
                for name in workbook.defined_names:
                    if '[' in name.value or 'http' in name.value.lower() or 'file:' in name.value.lower():
                        names_to_remove.append(name)
                
                for name in names_to_remove:
                    self.logger.info(f"Removing defined name with external reference: {name.name}")
                    del workbook.defined_names[name.name]
            
            # Clean pivot caches (which can contain external data)
            if hasattr(workbook, '_pivots'):
                workbook._pivots = []
            
            # Clean up each worksheet
            for ws in workbook.worksheets:
                self._clean_worksheet_external_data(ws)
                
            # Remove external references from workbook.xml.rels
            if hasattr(workbook, 'rels'):
                rels_to_remove = []
                for rel_id, rel in workbook.rels.items():
                    if rel.type.endswith('externalLink') or 'connections' in rel.type:
                        rels_to_remove.append(rel_id)
                
                for rel_id in rels_to_remove:
                    self.logger.info(f"Removing external relationship: {rel_id}")
                    del workbook.rels[rel_id]
            
        except Exception as e:
            self.logger.warning(f"Error cleaning external references with openpyxl: {e}")
    
    def _clean_worksheet_external_data(self, worksheet) -> None:
        """
        Clean external data from a worksheet.
        
        Args:
            worksheet: openpyxl Worksheet object
        """
        try:
            # Remove data validations that might reference external sources
            if hasattr(worksheet, 'data_validations'):
                worksheet.data_validations = openpyxl.worksheet.datavalidation.DataValidationList()
            
            # Remove any cell formulas that reference external files
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f' and cell.value and ('[' in cell.value or 'http' in str(cell.value).lower()):
                        cell.value = None
                        cell.data_type = 'n'
        except Exception as e:
            self.logger.warning(f"Error cleaning worksheet external data: {e}")
    
    def clean_external_references_com(self, file_path: str) -> bool:
        """
        Use Excel COM to thoroughly remove all external data references.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            bool: True if successful, False otherwise
        """
        excel = None
        wb = None
        
        try:
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Start Excel with limited automation security
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False
            
            # Constants
            xlExcelLinks = 1
            xlOLELinks = 2
            xlPublishers = 5
            xlTextImport = 6
            
            # Open workbook with updatelinks=False
            self.logger.info(f"Opening file with Excel COM to clean external references: {file_path}")
            wb = excel.Workbooks.Open(
                file_path,
                UpdateLinks=0,
                ReadOnly=False
            )
            
            # Step 1: Break all external links by type
            for link_type in [xlExcelLinks, xlOLELinks, xlPublishers, xlTextImport]:
                try:
                    self.logger.info(f"Breaking links of type {link_type}")
                    wb.BreakLink(Name="", Type=link_type)
                except Exception:
                    pass
            
            # Step 2: Remove all connections
            if hasattr(wb, 'Connections') and wb.Connections.Count > 0:
                self.logger.info(f"Removing {wb.Connections.Count} connections")
                
                # Get connection names (collection changes during removal)
                conn_names = []
                for i in range(1, wb.Connections.Count + 1):
                    try:
                        conn_names.append(wb.Connections(i).Name)
                    except:
                        pass
                
                # Remove each connection by name
                for name in conn_names:
                    try:
                        wb.Connections(name).Delete()
                    except Exception as e:
                        self.logger.warning(f"Failed to remove connection {name}: {e}")
            
            # Step 3: Process all sheets thoroughly
            for i in range(1, wb.Sheets.Count + 1):
                try:
                    sheet = wb.Sheets(i)
                    self._clean_sheet_data(sheet)
                except Exception as e:
                    self.logger.warning(f"Error cleaning sheet {i}: {e}")
            
            # Step 4: Check and remove any remaining data connections
            try:
                if hasattr(wb, 'Connections'):
                    wb.Connections.CommitAll()
            except:
                pass
            
            # Step 5: Save with special options
            self.logger.info("Saving workbook with external references removed")
            wb.Save()
            
            # Optional: Save as new file with cleanup
            temp_path = file_path + ".cleaned.xlsx"
            wb.SaveAs(
                temp_path,
                ConflictResolution=2  # xlLocalSessionChanges
            )
            
            # Close properly
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            # If we created a temporary cleaned file, replace the original
            if os.path.exists(temp_path):
                if os.path.exists(file_path):
                    os.unlink(file_path)
                os.rename(temp_path, file_path)
            
            return True
            
        except Exception as e:
            self.logger.error(f"COM cleaning error: {str(e)}")
            return False
            
        finally:
            # Cleanup
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
            
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
            
            pythoncom.CoUninitialize()
    
    def _clean_sheet_data(self, sheet) -> None:
        """
        Clean external data from a worksheet using COM.
        
        Args:
            sheet: Excel COM worksheet object
        """
        try:
            # Clean QueryTables (external data ranges)
            if hasattr(sheet, 'QueryTables'):
                try:
                    qt_count = sheet.QueryTables.Count
                    if qt_count > 0:
                        self.logger.info(f"Removing {qt_count} QueryTables from {sheet.Name}")
                        # Have to remove in reverse order as collection changes
                        for i in range(qt_count, 0, -1):
                            try:
                                sheet.QueryTables(i).Delete()
                            except:
                                pass
                except:
                    pass
                    
            # Clean ListObjects with external data
            if hasattr(sheet, 'ListObjects'):
                try:
                    lo_count = sheet.ListObjects.Count
                    if lo_count > 0:
                        for i in range(lo_count, 0, -1):
                            try:
                                table = sheet.ListObjects(i)
                                if hasattr(table, 'QueryTable'):
                                    self.logger.info(f"Removing QueryTable from ListObject in {sheet.Name}")
                                    table.QueryTable.Delete()
                            except:
                                pass
                except:
                    pass
                    
            # Clean PivotTables
            if hasattr(sheet, 'PivotTables'):
                try:
                    pt_count = sheet.PivotTables().Count
                    if pt_count > 0:
                        self.logger.info(f"Handling {pt_count} PivotTables in {sheet.Name}")
                        for i in range(pt_count, 0, -1):
                            try:
                                pt = sheet.PivotTables(i)
                                # Don't delete the pivot table but reset its connection
                                if hasattr(pt, 'PivotCache') and hasattr(pt.PivotCache, 'Connection'):
                                    pt.PivotCache.Connection = ""
                            except:
                                pass
                except:
                    pass
            
            # Clean Data Connections in charts
            if hasattr(sheet, 'ChartObjects'):
                try:
                    chart_count = sheet.ChartObjects().Count
                    if chart_count > 0:
                        for i in range(chart_count, 0, -1):
                            try:
                                chart = sheet.ChartObjects(i).Chart
                                if hasattr(chart, 'ChartData'):
                                    chart.ChartData.Activate()
                                    for j in range(chart.ChartData.Connections.Count, 0, -1):
                                        try:
                                            chart.ChartData.Connections(j).Delete()
                                        except:
                                            pass
                            except:
                                pass
                except:
                    pass
        except Exception as e:
            self.logger.warning(f"Error during COM sheet cleaning: {e}")

def clean_excel_external_data(file_path: str, logger=None) -> bool:
    """
    Clean all external data references from an Excel file using the most reliable methods.
    
    Args:
        file_path: Path to Excel file
        logger: Optional logger
        
    Returns:
        bool: True if cleaning was successful
    """
    logger = logger or logging.getLogger(__name__)
    cleaner = ExcelDataCleaner(logger)
    
    # First try with COM which is most thorough
    success = cleaner.clean_external_references_com(file_path)
    
    # Additional openpyxl cleanup as backup
    try:
        # Load with openpyxl
        wb = openpyxl.load_workbook(file_path, keep_vba=True)
        
        # Clean with openpyxl
        cleaner.clean_external_references_openpyxl(wb)
        
        # Save back
        wb.save(file_path)
        wb.close()
        
        logger.info("Completed openpyxl external references cleanup")
        return True
    except Exception as e:
        logger.warning(f"Error during openpyxl cleanup: {e}")
        return success