# Installation Guide - Logistics Document Generator

This guide will help you install and use the Logistics Document Generator package on any computer.

## Quick Installation

### Method 1: Install from Source (Recommended for Development)

1. **Download/Clone the Project**
   ```bash
   # If you have git
   git clone <repository-url>
   cd logistics-document-generator
   
   # Or download and extract the ZIP file
   ```

2. **Install the Package**
   ```bash
   # Install in development mode (recommended for updates)
   pip install -e .
   
   # Or install normally
   pip install .
   ```

3. **Install with GUI Support (Optional)**
   ```bash
   pip install -e ".[gui]"
   ```

### Method 2: Install from Wheel File

1. **Build the Package**
   ```bash
   python -m build
   ```

2. **Install the Wheel**
   ```bash
   pip install dist/logistics_document_generator-2.1.0-py3-none-any.whl
   ```

### Method 3: Install from PyPI (When Available)

```bash
pip install logistics-document-generator
```

## Verification

After installation, verify everything works:

```bash
# Check installation
placard-generator --info

# Validate setup
placard-generator --validate

# Setup directories
placard-generator --setup
```

## Usage

### Command Line Interface

```bash
# Interactive mode
placard-generator

# Process specific shipments
placard-generator -s 1234567890 9876543210

# Process all shipments
placard-generator --all

# Show help
placard-generator --help
```

### GUI Interface (if installed with GUI support)

```bash
placard-gui
```

### Python API

```python
from logistics_generator import PlacardGenerator

# Create generator instance
generator = PlacardGenerator()

# Setup and load data
generator.setup_directories()
generator.initialize_log()
generator.load_and_prepare_data()

# Process a shipment
generator.process_shipment("1234567890")
```

## Directory Structure

After installation, the package will create these directories:

- **Windows**: `%APPDATA%/LogisticsGenerator/`
- **macOS**: `~/Library/Application Support/LogisticsGenerator/`
- **Linux**: `~/.local/share/LogisticsGenerator/`

Within this directory:
- `Data/` - Place your Excel files here
- `Template/` - Place your Word template here
- `Placards/` - Generated documents will be saved here
- `Logs/` - Log files will be created here

## Configuration

### Excel File Requirements

Place Excel files in the `Data/` directory with names starting with:
`WM-SPN-CUS105 Open Order Report`

### Template Requirements

Place your Word template in the `Template/` directory with the name:
`placard_template.docx`

The template should contain these placeholders:
- `{{Ship To}}`
- `{{Shipment Nbr}}`
- `{{PO}}`
- `{{DO #}}`
- `{{VAS}}`
- `{{Original Qty}}`
- `{{Label Type}}`
- `{{Order Type}}`
- `{{Pmt Term}}`
- `{{Start Ship}}`

## Troubleshooting

### Common Issues

1. **Template not found**
   ```bash
   placard-generator --setup
   # Then place template in the shown Template directory
   ```

2. **Excel file not found**
   - Ensure Excel file is in the Data directory
   - Check filename starts with `WM-SPN-CUS105 Open Order Report`

3. **Permission errors**
   - Make sure you have write permissions to the data directory
   - On Windows, try running as administrator

4. **Dependency issues**
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

### Getting Help

```bash
# Show package information
placard-generator --info

# Validate installation
placard-generator --validate

# Show help
placard-generator --help
```

## Uninstallation

```bash
pip uninstall logistics-document-generator
```

## Development Setup

For development work:

```bash
# Clone repository
git clone <repository-url>
cd logistics-document-generator

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install in development mode with all extras
pip install -e ".[dev,gui,security]"

# Run tests
pytest

# Code formatting
black logistics_generator/
```

## System Requirements

- Python 3.9 or higher
- Windows 10/11, macOS 10.14+, or Linux
- At least 512MB RAM
- 100MB disk space for package + data files

## License

This package is licensed under the MIT License. See LICENSE file for details. 