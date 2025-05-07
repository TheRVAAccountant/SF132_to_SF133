"""
Unit tests for Excel processor functionality.
"""
import os
import unittest
from unittest.mock import MagicMock, patch
import tempfile
from queue import Queue
import shutil
from pathlib import Path

# Add parent directory to path to allow imports
import sys
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from excel_processor import ExcelProcessor
from excel_processor_config import ProcessingConfig

class TestExcelProcessor(unittest.TestCase):
    """Test cases for ExcelProcessor class."""
    
    def setUp(self):
        """Set up test environment."""
        self.queue = Queue()
        self.temp_dir = tempfile.mkdtemp()
        
        # Create a minimal config for testing
        self.config = ProcessingConfig(
            output_directory=self.temp_dir
        )
        
        # Create processor with mock configuration
        self.processor = ExcelProcessor(self.queue, processing_config=self.config)
        
        # Mock logging
        self.processor.logger = MagicMock()
    
    def tearDown(self):
        """Clean up after tests."""
        shutil.rmtree(self.temp_dir)
    
    def test_generate_new_filename(self):
        """Test new filename generation."""
        original_file = os.path.join(self.temp_dir, "test.xlsx")
        
        # Create empty file
        Path(original_file).touch()
        
        # Generate new filename
        new_file = self.processor._generate_new_filename(original_file)
        
        # Verify results
        self.assertTrue(os.path.dirname(new_file).startswith(self.temp_dir))
        self.assertTrue('test_processed_' in os.path.basename(new_file))
        self.assertTrue(new_file.endswith('.xlsx'))
    
    def test_create_backup_file(self):
        """Test backup file creation."""
        # Create test file
        file_path = os.path.join(self.temp_dir, "test_backup.xlsx")
        Path(file_path).touch()
        
        # Create backup
        backup_path = self.processor._create_backup_file(file_path)
        
        # Verify backup was created
        self.assertTrue(os.path.exists(backup_path))
        self.assertTrue('backup' in backup_path)
    
    @patch('win32com.client.Dispatch')
    @patch('pythoncom.CoInitialize')
    def test_verify_workbook_integrity(self, mock_coinit, mock_dispatch):
        """Test workbook integrity verification with mocked COM objects."""
        # Setup mocks
        mock_excel = MagicMock()
        mock_dispatch.return_value = mock_excel
        mock_excel.Workbooks.Open.return_value = MagicMock()
        
        # Create test file
        file_path = os.path.join(self.temp_dir, "verify_test.xlsx")
        Path(file_path).touch()
        
        # Call the method
        self.processor._verify_workbook_integrity(file_path)
        
        # Verify mocks were called
        mock_coinit.assert_called_once()
        mock_dispatch.assert_called_once_with("Excel.Application")
        mock_excel.Workbooks.Open.assert_called_once()

if __name__ == '__main__':
    unittest.main()
