# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## IMPORTANT NOTE FOR CODE ASSISTANCE

When working with this codebase, always prioritize the packaged scripts in the `src/sf132_sf133_recon/` directory. The root-level scripts are for legacy development and testing purposes only. All modifications, bug fixes, and enhancements should be made to the files in the `src/` directory, which contains the current and properly packaged version of the application.

## Project Overview

SF132_to_SF133 is a Python application for processing SF132 and SF133 reconciliation Excel files. It automates the analysis and comparison of financial data, with features including:

- Loading and processing reconciliation Excel files
- Unprotecting sheets with password authentication
- Analyzing differences and providing comments on reconciliation items
- Creating processed output files with timestamping
- Modern GUI interface with progress tracking
- Robust error handling and recovery mechanisms

## Environment Setup

### Requirements
- Python 3.9 or higher
- Windows OS (required for COM automation with Excel)

### Installation
```bash
# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Running the Application

### IMPORTANT: Use Packaged Scripts from src Directory
Always use the packaged scripts in the `src/sf132_sf133_recon/` directory when working with this codebase. The root-level scripts are for development only and may contain errors.

### GUI Mode
```bash
python -m src.sf132_sf133_recon.main
```

### CLI Mode
```bash
python -m src.sf132_sf133_recon.main [excel_file_path] [optional_password]
```

## Running Tests

```bash
# Run all tests
python -m unittest discover -s tests

# Run a specific test file
python -m unittest tests/test_excel_processor.py

# Run a specific test case
python -m unittest tests.test_excel_processor.TestExcelProcessor.test_generate_new_filename
```

## Architecture

### Key Components

1. **ExcelProcessingApplication** (`src/sf132_sf133_recon/main.py`)
   - Main application class that coordinates GUI, processing, and logging
   - Handles command-line and GUI modes
   - Sets up logging, error handling, and dependency checks

2. **ExcelProcessorGUI** (`src/sf132_sf133_recon/modules/gui.py`)
   - Provides the GUI interface using tkinter
   - Handles file selection, password input, progress tracking
   - Uses a queue-based communication system with the processor

3. **ExcelProcessor** (`src/sf132_sf133_recon/core/excel_processor.py`)
   - Core class that handles Excel file processing
   - Includes multiple fallback methods for processing files
   - Features robust error handling and repair mechanisms
   - Uses both openpyxl and win32com (COM) for Excel manipulation

4. **Configuration System**
   - `src/sf132_sf133_recon/config.py` - Application-wide settings
   - Environment variables - For deployment-specific configurations

5. **Logging and Error Handling**
   - `src/sf132_sf133_recon/utils/logger.py` - Sets up logging and error handling
   - Performance logging decorators
   - Exception capturing and reporting

### Data Flow

1. User selects an Excel file through GUI or CLI
2. Application validates the file and creates necessary directories
3. ExcelProcessor makes a verified copy of the original file
4. Processor attempts to process the file using various methods with fallbacks:
   - Fresh workbook approach
   - Library-based processing
   - COM-based processing
5. Processing includes:
   - Unprotecting sheets
   - Finding column indexes
   - Processing header formatting
   - Adding "DO Comments" column
   - Processing rows with comments
6. Results are saved to a timestamped output file
7. Progress and status updates are communicated back to the GUI

## Configuration

The application can be configured through:

- `app_config.json` - Application-wide settings (generated on first run)
- Environment variables:
  - `SF132_HEADER_ROW` - Header row number
  - `SF132_SHEET_NAME` - Target worksheet name
  - `SF132_OUTPUT_DIR` - Output directory
  - `SF132_MAX_RETRIES` - Maximum retry attempts
  - `SF132_VERIFY_SAVE` - Whether to verify after saving
  - `SF132_CLEAN_CONNECTIONS` - Whether to clean external connections

## Security Notes

- The application handles sheet protection passwords
- It cleans up external data connections for security
- Temp files are tracked and cleaned up