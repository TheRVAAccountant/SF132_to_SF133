#!/usr/bin/env python3
"""
Entry point script for SF132 to SF133 Reconciliation Tool.

This script provides a simple launcher for the refactored package.
This application is designed for Windows and requires Windows-specific features.
"""

import sys
import os
import platform
import argparse

# Basic platform check - this application is designed for Windows
if not sys.platform.startswith('win'):
    print("WARNING: This application is designed specifically for Windows.")
    print(f"Current platform: {platform.platform()}")
    print("Some features may not work correctly on non-Windows platforms.")

# Add the current directory to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Verify Windows dependencies are available
if sys.platform.startswith('win'):
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("ERROR: Required Windows dependencies not found.")
        print("This application requires the following packages:")
        print("- pywin32")
        print("\nPlease install these packages with pip:")
        print("pip install pywin32")
        sys.exit(1)

def run():
    """
    Run the SF132 to SF133 Reconciliation Tool.
    
    This function parses command-line arguments and runs the application
    with the appropriate settings.
    
    Returns:
        int: Exit code (0 for success, 1 for failure)
    """
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description="SF132 to SF133 Excel Reconciliation Tool"
    )
    parser.add_argument(
        "file", nargs="?", help="Path to Excel file to process"
    )
    parser.add_argument(
        "password", nargs="?", help="Sheet protection password"
    )
    parser.add_argument(
        "--recover", "-r", action="store_true", 
        help="Run in recovery mode for file access issues"
    )
    
    args = parser.parse_args()
    
    # Find and import the main function dynamically to avoid linting errors
    main_function = None
    
    # Build a list of possible import paths
    import_paths = [
        # Local import for compatibility mode
        {'module': 'main', 'attr': 'main'},
        # Package import from src directory
        {'module': 'src.sf132_sf133_recon.main', 'attr': 'main'},
        # Package import for installed package
        {'module': 'sf132_sf133_recon.main', 'attr': 'main'}
    ]
    
    # Try each import path
    for path in import_paths:
        try:
            module_name = path['module']
            attr_name = path['attr']
            
            # Use __import__ to avoid static analysis errors
            module = __import__(module_name, fromlist=[attr_name])
            if hasattr(module, attr_name):
                main_function = getattr(module, attr_name)
                print(f"Successfully imported {attr_name} from {module_name}")
                break
        except ImportError:
            continue
    
    # Check if we found the main function
    if main_function is None:
        print("ERROR: Could not import main module from any location.")
        print("Make sure you are running this script from the correct directory.")
        sys.exit(1)
    
    # Run the main application
    # The argparse arguments will be available as sys.argv
    return main_function()

if __name__ == "__main__":
    sys.exit(run())