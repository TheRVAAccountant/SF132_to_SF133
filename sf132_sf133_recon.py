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
    
    # Import the main function in a way that works both when installed as a package
    # and when run directly from the repository
    main = None
    
    # Try different import paths
    import_paths = [
        # Regular import for when module is installed
        lambda: __import__('sf132_sf133_recon.main').main.main,
        # Import from src for development mode
        lambda: __import__('src.sf132_sf133_recon.main').sf132_sf133_recon.main.main,
        # Direct import for compatibility mode
        lambda: __import__('main').main
    ]
    
    # Try each import path until one works
    for import_func in import_paths:
        try:
            main = import_func()
            break
        except (ImportError, AttributeError):
            continue
    
    if main is None:
        print("ERROR: Could not import main module from any location.")
        print("Make sure you are running this script from the correct directory.")
        sys.exit(1)
    
    # Run the main application
    # The argparse arguments will be available as sys.argv
    return main()

if __name__ == "__main__":
    sys.exit(run())