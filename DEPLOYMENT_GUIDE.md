# Deployment Guide - Logistics Document Generator Package

## Overview

Your placard generator has been successfully converted into a Python package that can be installed and used on any computer. The package includes:

- **Command-line interface** for terminal usage
- **GUI interface** for graphical usage
- **Python API** for programmatic usage
- **Secure file handling** and validation
- **Cross-platform compatibility** (Windows, macOS, Linux)

## Package Structure

```
logistics-document-generator/
├── logistics_generator/           # Main package
│   ├── __init__.py               # Package initialization
│   ├── core.py                   # Main generator logic
│   ├── security.py               # Security utilities
│   ├── config.py                 # Configuration
│   ├── cli.py                    # Command-line interface
│   ├── gui.py                    # GUI interface
│   ├── utils.py                  # Utility functions
│   └── templates/                # Template files
│       └── placard_template.docx # Default template
├── setup.py                      # Package setup (legacy)
├── pyproject.toml                # Modern package configuration
├── MANIFEST.in                   # Package manifest
├── requirements.txt              # Dependencies
├── LICENSE                       # MIT License
├── README.md                     # Documentation
├── INSTALL.md                    # Installation guide
└── dist/                         # Built packages
    ├── logistics_document_generator-2.1.0.tar.gz
    └── logistics_document_generator-2.1.0-py3-none-any.whl
```

## Installation Methods

### Method 1: Install from Built Package (Recommended)

On the target computer:

```bash
# Install from wheel file
pip install logistics_document_generator-2.1.0-py3-none-any.whl

# Or install from source distribution
pip install logistics_document_generator-2.1.0.tar.gz
```

### Method 2: Install from Source

```bash
# From the project directory
pip install .

# For development (editable install)
pip install -e .

# With all optional dependencies
pip install -e ".[gui,dev,security]"
```

### Method 3: Direct Installation (Advanced)

Copy the entire `logistics_generator` folder to the target computer and:

```bash
# Add to Python path
export PYTHONPATH="${PYTHONPATH}:/path/to/logistics_generator"

# Or install manually
python -m pip install --user -e /path/to/project
```

## Post-Installation Setup

### 1. Verify Installation

```bash
# Check if commands are available
placard-generator --version
placard-generator --info

# Validate installation
placard-generator --validate
```

### 2. Setup Directories

```bash
# Create required directories
placard-generator --setup
```

This will create:
- **Windows**: `%APPDATA%\LogisticsGenerator\`
- **macOS**: `~/Library/Application Support/LogisticsGenerator/`
- **Linux**: `~/.local/share/LogisticsGenerator/`

### 3. Add Data Files

After setup, place your files in the created directories:

```
LogisticsGenerator/
├── Data/                          # Excel files go here
│   └── WM-SPN-CUS105 Open Order Report_*.xlsx
├── Template/                      # Template files
│   └── placard_template.docx
├── Placards/                      # Generated output (auto-created)
└── Logs/                          # Log files (auto-created)
```

## Usage Examples

### Command Line Usage

```bash
# Interactive mode
placard-generator

# Process specific shipments
placard-generator -s 1234567890 9876543210

# Process all shipments
placard-generator --all

# Verbose output
placard-generator -v -s 1234567890

# Show help
placard-generator --help
```

### GUI Usage

```bash
# Launch GUI (if installed with GUI support)
placard-gui
```

### Python API Usage

```python
# Import the package
from logistics_generator import PlacardGenerator

# Create generator
generator = PlacardGenerator()

# Setup and initialize
generator.setup_directories()
generator.initialize_log()

# Load data
if generator.load_and_prepare_data():
    # Process a single shipment
    generator.process_shipment("1234567890")
    
    # Or get all shipments and process them
    shipments = generator.get_all_unique_shipments()
    for shipment in shipments:
        generator.process_shipment(shipment)
```

## Distribution Options

### Option 1: Share Built Packages

1. **Create distribution package**:
   ```bash
   # Already done - use files in dist/
   tar -czf logistics-generator-package.tar.gz dist/ INSTALL.md README.md
   ```

2. **Share with users**:
   - Send `logistics_document_generator-2.1.0-py3-none-any.whl`
   - Include `INSTALL.md` for installation instructions

### Option 2: Create Installer Script

Create `install.sh` (Linux/macOS) or `install.bat` (Windows):

```bash
#!/bin/bash
# install.sh
echo "Installing Logistics Document Generator..."
pip install logistics_document_generator-2.1.0-py3-none-any.whl
echo "Setting up directories..."
placard-generator --setup
echo "Installation complete!"
echo "Run 'placard-generator --help' for usage information"
```

### Option 3: Docker Container

Create `Dockerfile`:

```dockerfile
FROM python:3.11-slim

WORKDIR /app
COPY dist/logistics_document_generator-2.1.0-py3-none-any.whl .
RUN pip install logistics_document_generator-2.1.0-py3-none-any.whl

# Setup directories
RUN placard-generator --setup

# Create volume for data
VOLUME ["/data"]

# Default command
CMD ["placard-generator", "--help"]
```

Build and run:
```bash
docker build -t logistics-generator .
docker run -v /path/to/data:/data logistics-generator
```

## Deployment Checklist

### Pre-Deployment
- [ ] Test package installation on clean environment
- [ ] Verify all dependencies are included
- [ ] Test CLI commands work correctly
- [ ] Test GUI (if applicable)
- [ ] Verify template files are included
- [ ] Test with sample data

### During Deployment
- [ ] Install Python 3.9+ on target system
- [ ] Install the package using pip
- [ ] Run `placard-generator --validate`
- [ ] Setup directories with `placard-generator --setup`
- [ ] Copy data files to appropriate directories
- [ ] Test with a sample shipment

### Post-Deployment
- [ ] Verify output files are generated correctly
- [ ] Check log files for errors
- [ ] Test edge cases and error handling
- [ ] Provide user training if needed
- [ ] Document any system-specific configurations

## Troubleshooting

### Common Issues

1. **Command not found**
   ```bash
   # Check if package is installed
   pip list | grep logistics
   
   # Reinstall if needed
   pip install --force-reinstall logistics_document_generator-2.1.0-py3-none-any.whl
   ```

2. **Permission errors**
   ```bash
   # Install for user only
   pip install --user logistics_document_generator-2.1.0-py3-none-any.whl
   ```

3. **Missing dependencies**
   ```bash
   # Install with all dependencies
   pip install -r requirements.txt
   pip install logistics_document_generator-2.1.0-py3-none-any.whl
   ```

4. **Template not found**
   ```bash
   # Check setup
   placard-generator --validate
   
   # Re-run setup
   placard-generator --setup
   ```

### Support Commands

```bash
# Get system information
placard-generator --info

# Validate installation
placard-generator --validate

# Check directories
placard-generator --setup

# Verbose debugging
placard-generator -v --validate
```

## Security Considerations

The package includes built-in security features:
- Input validation and sanitization
- Path traversal protection
- File size and type validation
- Rate limiting
- Comprehensive logging

For production use:
- Keep the package updated
- Monitor log files for security events
- Restrict file system permissions
- Use virtual environments

## Updates and Maintenance

### Updating the Package

1. **Build new version**:
   ```bash
   # Update version in pyproject.toml
   python -m build
   ```

2. **Install update**:
   ```bash
   pip install --upgrade dist/logistics_document_generator-2.1.1-py3-none-any.whl
   ```

### Monitoring

- Check log files in the Logs directory
- Monitor system performance
- Validate output file quality
- Review security logs

## Contact and Support

For issues, questions, or feature requests:
- Check the documentation in README.md
- Run diagnostic commands
- Review log files
- Contact the development team

---

**Note**: This package is designed for secure, production use with comprehensive error handling and logging. Always test in a development environment before deploying to production. 