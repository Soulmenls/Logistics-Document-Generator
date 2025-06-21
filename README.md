# Logistics Document Generator

A high-performance, class-based Python application that generates professional multi-page shipping placards from Excel data using Word templates with advanced formatting preservation.

## Architecture Overview

Built using object-oriented design with the `PlacardGenerator` class that provides:

- **Memory-efficient data processing** with pandas vectorized operations
- **Advanced formatting preservation** across document generations
- **Robust error handling** with detailed validation and user feedback
- **Batch processing capabilities** for multiple shipments in one session

## Key Features

- **One-time data loading**: Reads Excel file once at startup using pandas for optimal performance
- **Smart input validation**: Validates shipment numbers (exactly 10 digits) and DO # formats (minimum 8 digits)
- **Multi-page document generation**: Creates separate pages for each DO # within a shipment with page breaks
- **Template-based generation**: Uses customizable Word templates with intelligent placeholder replacement
- **Advanced formatting preservation**: Maintains fonts, sizes, bold text, and spacing across all pages
- **Interactive CLI operation**: Process multiple shipment numbers with continue/exit options
- **Comprehensive error handling**: Graceful handling of missing files, invalid data, and permission errors

## Dependencies

The application requires these Python packages:

```bash
# Core dependencies
pandas          # Data manipulation and analysis
python-docx     # Word document generation and manipulation
openpyxl        # Excel file reading support
```

## Setup Instructions

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Directory Structure

The application automatically creates these directories:

```text
Logistics Document Generator/
├── Data/                    # Excel data files (input)
├── Template/               # Word template files
├── Placards/              # Generated documents (output, auto-created)
├── placard_generator.py   # Main application
└── requirements.txt       # Dependencies
```

### 3. Required Files

#### Excel Data File

**Location**: `Data/` folder
**Naming**: Must start with `"WM-SPN-CUS105 Open Order Report"`
**Formats**: `.xlsx` or `.xls` supported

**Required Columns** (case-sensitive):

- `Shipment Nbr` - 10-digit shipment identifier
- `DO #` - Delivery Order number (minimum 8 digits)
- `Label Type` - Shipment classification
- `Order Type` - Order classification
- `Pmt Term` - Payment terms
- `Start Ship` - Ship date (auto-formatted to MM/DD/YYYY)
- `VAS` - Value Added Service (Y/N, converted to "VAS"/"NOT VAS")
- `Ship To` - Destination information
- `PO` - Purchase Order numbers (multiple POs aggregated per DO #)
- `Original Qty` - Quantity values (auto-formatted with "Units" suffix)

#### Word Template

**Location**: `Template/placard_template.docx`
**Placeholders** (exact format required):

```text
{{Ship To}}        - Destination address/information
{{Shipment Nbr}}   - Shipment number (cleaned of decimals)
{{PO}}             - Purchase orders (newline-separated if multiple)
{{DO #}}           - Delivery order (formatted with leading zeros to 10 digits)
{{VAS}}            - Value added service status  
{{Original Qty}}   - Total quantity with "Units" suffix
{{Label Type}}     - Shipment label classification
{{Order Type}}     - Order type classification
{{Pmt Term}}       - Payment terms
{{Start Ship}}     - Formatted ship date (MM/DD/YYYY)
```

## Usage

### Running the Application

```bash
python placard_generator.py
```

### Operation Workflow

1. **Initialization**: Application loads and validates Excel data, reports data quality
2. **Input Processing**: Enter shipment numbers (comma-separated for batch processing)
3. **Data Validation**: Validates each shipment number format and existence
4. **Document Generation**: Creates multi-page documents with formatting preservation
5. **Output Confirmation**: Reports success/failure for each shipment processed
6. **Session Management**: Option to process additional shipments or exit

### Input Specifications

- **Shipment Numbers**: Exactly 10 digits, no letters/special characters
- **Batch Input**: Comma-separated values (e.g., `"1234567890, 9876543210"`)
- **Data Matching**: Handles float values in Excel (e.g., `9010157586.0` → `9010157586`)

## Technical Implementation

### Data Processing Pipeline

1. **File Discovery**: Locates Excel files using glob pattern matching
2. **Data Loading**: Pandas reads Excel with automatic type inference
3. **Data Cleaning**: Removes rows with empty Shipment Nbr
4. **Validation**: Filters out invalid DO # formats using regex validation
5. **Memory Storage**: Keeps validated data in memory for fast lookups

### Document Generation Process

For each shipment:

1. **Data Grouping**: Groups records by DO # using pandas groupby
2. **Data Aggregation**:
   - Combines unique PO numbers (newline-separated)
   - Sums Original Qty values per DO #
   - Preserves shipment-level metadata
3. **Template Processing**: Creates fresh document copy for each page
4. **Placeholder Replacement**: Advanced text replacement preserving formatting
5. **Multi-page Assembly**: Adds page breaks and copies formatted content
6. **File Output**: Saves with standardized naming convention

### Advanced Formatting Preservation

The application uses sophisticated formatting preservation techniques:

- **Run-level formatting**: Preserves font names, sizes, colors, and styles
- **Paragraph formatting**: Maintains alignment, spacing, and line spacing
- **Mixed formatting handling**: Preserves different formatting within single paragraphs
- **Cross-run placeholder replacement**: Handles placeholders spanning multiple formatted runs
- **Table formatting**: Preserves table styles and cell formatting
- **Header/footer preservation**: Maintains formatting in document headers and footers

### Data Transformation Details

#### Shipment-Level Data

(consistent across all pages):

- `Shipment Nbr`: Converted to clean integer (removes `.0` decimals)
- `Label Type`, `Order Type`, `Pmt Term`: Used as-is from first record
- `Start Ship`: Formatted as MM/DD/YYYY using pandas datetime parsing
- `VAS`: Converted to "VAS" (if Y) or "NOT VAS" (if N or empty)

#### DO-Level Data

(specific per page):

- `DO #`: Formatted with leading zeros to 10 digits total (e.g., `"66455734"` → `"0066455734"`)
- `Ship To`: From first record of DO # group
- `PO`: All unique POs for DO #, joined with newlines
- `Original Qty`: Sum of quantities with "Units" suffix (e.g., `"7782 Units"`)

## Error Handling & Validation

### File Validation

- **Missing Excel file**: Clear error with search pattern details
- **Missing template**: Specific path validation with helpful messages
- **Permission errors**: Graceful handling of read/write access issues

### Data Validation

- **Column verification**: Reports specific missing required columns
- **Shipment number format**: Regex validation for exact 10-digit requirement
- **DO # format**: Minimum 8-digit validation with detailed error messages
- **Data existence**: Validates shipment exists in filtered dataset

### Processing Resilience

- **Continue on error**: Processes remaining shipments after individual failures
- **Detailed logging**: Reports processing status for each DO # and shipment
- **Session summaries**: Provides counts of successful/failed operations

## Performance Optimizations

- **Single file read**: Excel loaded once at startup, not per shipment
- **Vectorized operations**: Pandas operations for efficient filtering and grouping
- **Memory-based processing**: All operations work with in-memory datasets
- **Batch processing**: Multiple shipments processed in single session
- **Template reuse**: Efficient document copying with formatting preservation

## Output Specifications

**File Location**: `Placards/` folder (auto-created)
**Naming Convention**: `Placard_[ShipmentNumber].docx`
**Content Structure**: Multi-page document with one page per DO #
**Formatting**: Complete template formatting preserved across all pages

**Data Formatting**:

- DO # with leading zeros (10 digits)
- Quantities with "Units" suffix
- Clean shipment numbers (no decimals)
- Formatted dates (MM/DD/YYYY)

## Troubleshooting Guide

### Setup Issues

#### No Excel file found

- Verify file is in `Data/` folder
- Check filename starts with `"WM-SPN-CUS105 Open Order Report"`
- Ensure file extension is `.xlsx` or `.xls`

#### Template file not found

- Confirm `placard_template.docx` exists in `Template/` folder
- Check exact spelling and capitalization
- Verify file is not corrupted or password-protected

### Data Issues

#### Missing required columns

- Excel must contain all 10 required columns (case-sensitive)
- Check for extra spaces or different casing in column names
- Verify columns are not hidden in Excel

#### Invalid shipment number format

- Shipment numbers must be exactly 10 digits
- Remove any letters, spaces, or special characters
- Check for leading/trailing spaces in input

#### No data found for shipment

- Verify shipment number exists in Excel data
- Check if shipment was filtered out during validation
- Ensure DO # values meet minimum 8-digit requirement

### Processing Issues

#### Error saving document

- Check write permissions to `Placards/` folder
- Ensure target file is not open in another application
- Verify sufficient disk space available

#### Error copying formatted content

- Template file may be corrupted or incompatible
- Try recreating template with simpler formatting
- Check for unsupported Word features in template

## Example Session Output

```text
=== Shipping Placard Generator ===
Loading data and preparing system...
Loading file: Data/WM-SPN-CUS105 Open Order Report 2024.xlsx
Loaded 1500 rows from Excel file
Removed 25 rows with empty Shipment Nbr
Removed 10 rows with invalid DO # format
Final dataset: 1465 rows ready for processing

Data loaded successfully! Ready to generate placards.

Enter one or more Shipment Numbers (comma-separated): 1234567890, 9876543210

Processing shipment: 1234567890
Found 15 records for shipment 1234567890
Processing 3 DO #s for shipment 1234567890
  Processing DO # 66455734 (1/3)
  Processing DO # 66455735 (2/3)  
  Processing DO # 66455736 (3/3)
SUCCESS: Created placard document: Placards/Placard_1234567890.docx

Processing shipment: 9876543210
Found 8 records for shipment 9876543210
Processing 2 DO #s for shipment 9876543210
  Processing DO # 77889900 (1/2)
  Processing DO # 77889901 (2/2)
SUCCESS: Created placard document: Placards/Placard_9876543210.docx

=== Processing Summary ===
Documents created: 2
Failed inputs: 0

Process more shipments? (y/n): n

=== Final Summary ===
Total documents created: 2
Total failed inputs: 0
Thank you for using the Shipping Placard Generator!
```

## Technical Requirements

- **Python**: 3.7+ (for type hints and f-string support)
- **Operating System**: Windows, macOS, Linux (cross-platform)
- **Memory**: Sufficient for Excel dataset size (typically < 100MB)
- **Storage**: Space for Excel files, templates, and generated documents
