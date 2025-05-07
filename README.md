# SF132 to SF133 Excel Processor

A professional application for processing SF132 and SF133 reconciliation Excel files. This tool automates the analysis and comparison of financial data.

## Features

- Load and process SF132 to SF133 reconciliation Excel files
- Unprotect sheets with password authentication
- Analyze differences and provide comments on reconciliation items
- Create processed output files with timestamping
- Modern GUI interface with progress tracking
- Robust error handling and recovery mechanisms

## Installation

### Prerequisites

- Python 3.9 or higher
- Windows OS (required for COM automation with Excel)

### Setup

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

1. Run the application:
   ```
   python main.py
   ```
2. Use the GUI to:
   - Select an input Excel file
   - Enter the sheet password
   - Process the file
   - View progress and results

## Configuration

The application can be configured through:

- `app_config.json` - Application-wide settings
- Environment variables - For deployment-specific configurations:
  - `SF132_HEADER_ROW` - Header row number
  - `SF132_SHEET_NAME` - Target worksheet name
  - `SF132_OUTPUT_DIR` - Output directory
  - `SF132_MAX_RETRIES` - Maximum retry attempts for operations

## Development

### Running Tests

Run the test suite with:
