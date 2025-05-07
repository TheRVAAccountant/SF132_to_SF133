# SF132 to SF133 Excel Processor

A professional application for processing SF132 and SF133 reconciliation Excel files. This tool automates the analysis and comparison of financial data.

## Features

- Load and process SF132 to SF133 reconciliation Excel files
- Unprotect sheets with password authentication
- Analyze differences and provide comments on reconciliation items
- Create processed output files with timestamping
- Modern GUI interface with progress tracking
- Robust error handling and recovery mechanisms
- Windows-specific optimizations for handling file access errors

## Installation

### Prerequisites

- Python 3.9 or higher
- Windows OS (REQUIRED - this application is designed specifically for Windows)
- pywin32 package (for COM automation with Excel)
- openpyxl and pandas packages for Excel processing

### Setup

#### Option 1: Install as a Package

```bash
# Install from the current directory
pip install -e .

# Or install with Windows-specific dependencies
pip install -e .[windows]
```

#### Option 2: Development Setup

1. Clone this repository or download the source code
2. Create a virtual environment:
   ```
   python -m venv venv
   venv\Scripts\activate
   ```
3. Install required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

### GUI Mode

Run the application with the GUI interface:

```bash
# If installed as a package
sf132-sf133-recon

# Or run directly
python sf132_sf133_recon.py
```

### CLI Mode

Process a file directly from the command line:

```bash
# If installed as a package
sf132-sf133-recon path/to/file.xlsx [password]

# Or run directly
python sf132_sf133_recon.py path/to/file.xlsx [password]
```

### Recovery Mode

If you're experiencing file access errors, use recovery mode:

```bash
sf132-sf133-recon --recover path/to/file.xlsx [password]
```

## Project Structure

The codebase is organized into the following structure:

```
sf132_sf133_recon/
├── __init__.py         # Package initialization
├── main.py             # Main application entry point
├── config.py           # Configuration classes and settings
├── core/               # Core functionality
│   ├── __init__.py
│   └── excel_processor.py  # Main Excel processing logic
├── modules/            # Supporting modules
│   ├── __init__.py
│   ├── com_operations.py   # Windows COM automation
│   ├── excel_cleaner.py    # Excel data cleanup
│   ├── excel_handler.py    # Excel process management
│   ├── excel_recovery.py   # File recovery mechanisms
│   ├── file_operations.py  # File handling operations
│   └── gui.py              # GUI implementation
└── utils/              # Utility functions
    ├── __init__.py
    ├── logger.py           # Logging configuration
    └── win_path_handler.py # Windows path handling
```

## Configuration

The application can be configured through:

- `pyproject.toml` - Package configuration
- `app_config.json` - Application-wide settings (generated on first run)
- Environment variables - For deployment-specific configurations:
  - `SF132_HEADER_ROW` - Header row number
  - `SF132_SHEET_NAME` - Target worksheet name
  - `SF132_OUTPUT_DIR` - Output directory
  - `SF132_MAX_RETRIES` - Maximum retry attempts for operations
  - `SF132_VERIFY_SAVE` - Whether to verify after saving
  - `SF132_CLEAN_CONNECTIONS` - Whether to clean external connections

## Development

### Running Tests

Run the test suite with:

```bash
# Run all tests
python -m unittest discover -s tests

# Run specific tests
python -m unittest tests.test_excel_processor
```

### Building the Package

To build the package for distribution:

```bash
# Install build dependencies
pip install build

# Build the package
python -m build
```

## Compatibility Layer

For projects using the old module structure, a compatibility layer is provided:

```python
# Old imports will still work through the compatibility layer
from compat import ExcelProcessor, ExcelProcessingApplication

# But new imports are recommended
from sf132_sf133_recon.core.excel_processor import ExcelProcessor
from sf132_sf133_recon.main import ExcelProcessingApplication
```