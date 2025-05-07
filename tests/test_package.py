"""
Test the refactored package structure.

This test ensures that the new package structure is functional
and that both entry points (sf132_sf133_recon.py and the package's
main module) can be imported and used.
"""

import os
import sys
import unittest
from pathlib import Path

# Add the parent directory to the path for imports
parent_dir = Path(__file__).parent.parent
sys.path.insert(0, str(parent_dir))

class PackageStructureTests(unittest.TestCase):
    """Tests for the refactored package structure."""
    
    def test_entry_point_imports(self):
        """Test that the entry point script can be imported."""
        try:
            # Import the entry point script
            import sf132_sf133_recon
            self.assertTrue(True, "Entry point import succeeded")
        except ImportError as e:
            self.fail(f"Entry point import failed: {e}")
    
    def test_package_imports(self):
        """Test that the package modules can be imported."""
        try:
            # Import main package modules
            from src.sf132_sf133_recon import main
            from src.sf132_sf133_recon.core import excel_processor
            from src.sf132_sf133_recon.utils import logger
            from src.sf132_sf133_recon.modules import file_operations
            
            # Check that key classes and functions exist
            self.assertTrue(hasattr(main, "ExcelProcessingApplication"))
            self.assertTrue(hasattr(excel_processor, "ExcelProcessor"))
            self.assertTrue(hasattr(logger, "setup_logging"))
            self.assertTrue(hasattr(file_operations, "create_verified_copy"))
            
            self.assertTrue(True, "Package imports succeeded")
        except ImportError as e:
            self.fail(f"Package import failed: {e}")
    
    def test_compatibility_layer(self):
        """Test that the compatibility layer works."""
        try:
            # Import from the compatibility layer
            import compat
            
            # Check that key classes are available
            self.assertTrue(hasattr(compat, "ExcelProcessor"))
            self.assertTrue(hasattr(compat, "ExcelProcessingApplication"))
            
            self.assertTrue(True, "Compatibility imports succeeded")
        except ImportError as e:
            self.fail(f"Compatibility import failed: {e}")
    
    def test_import_main_module(self):
        """Test that the main module can be imported and functions can be called."""
        try:
            # Import the main module
            from src.sf132_sf133_recon.main import create_required_directories
            
            # Call a function to verify it works
            create_required_directories()
            
            # Check if directories were created
            self.assertTrue(os.path.exists("logs"))
            self.assertTrue(os.path.exists("output"))
            self.assertTrue(os.path.exists("output/backups"))
            
            self.assertTrue(True, "Main module functions work correctly")
        except Exception as e:
            self.fail(f"Main module function call failed: {e}")

if __name__ == "__main__":
    unittest.main()