# Logistics Document Generator

An interactive, high-performance Python script that generates multi-page shipping
placards on-demand using Excel data and Word templates.

## Features

- **One-time data loading**: Reads Excel file once at startup for optimal performance
- **Input validation**: Validates shipment numbers and DO # formats
- **Multi-page documents**: Creates separate pages for each DO # within a shipment
- **Template-based**: Uses customizable Word templates with placeholder replacement
- **Error handling**: Comprehensive error checking and user-friendly messages
- **Interactive operation**: Allows processing multiple shipment numbers in one session

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Directory Structure

Ensure your project has these folders:

- `Data/` - Place your Excel file here
- `Template/` - Place your Word template here  
- `Placards/` - Generated documents will be saved here (created automatically)

### 3. Required Files

#### Excel Data File

- Must be in the `Data/` folder
- Filename must start with "WM-SPN-CUS105 Open Order Report"
- Must contain these columns (case-sensitive):

  - Shipment Nbr
  - DO # (must be at least 8 digits)
  - Label Type
  - Order Type
  - Pmt Term
  - Start Ship
  - VAS
  - Ship To
  - PO
  - Original Qty

#### Word Template

- Must be named `placard_template.docx` in the `Template/` folder
- Include placeholders that will be replaced:

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

## Usage

### Running the Script

```bash
python placard_generator.py
```

### Operation Flow

1. **Startup**: Script loads and validates Excel data
2. **Input**: Enter shipment numbers (comma-separated)
3. **Processing**: Script generates placards for each valid shipment
4. **Output**: Documents saved to `Placards/` folder as `Placard_[ShipmentNbr].docx`
5. **Continue**: Option to process more shipments or exit

### Input Requirements

- **Shipment Numbers**: Must be exactly 10 digits
- **Multiple entries**: Separate with commas (e.g., "1234567890, 9876543210")

## How It Works

### Data Processing

1. Loads Excel file once at startup
2. Filters out rows with empty Shipment Nbr
3. Validates DO # format (at least 8 digits)
4. Keeps data in memory for fast lookups

### Document Generation

For each shipment:

1. Groups data by DO # (each DO # = one page)
2. Aggregates PO numbers within each DO #
3. Calculates total Original Qty per DO #
4. Creates multi-page document using template
5. Replaces all placeholders with actual data
6. Formats DO # with leading zeros (10 digits total)
7. Adds "Units" suffix to quantity values
8. Preserves template formatting across all pages

### Data Mapping

**Shipment-Level Data** (same on all pages):

- Shipment Nbr (clean integer format without decimals)
- Label Type, Order Type, Pmt Term
- Start Ship (formatted as MM/DD/YYYY)
- VAS ("VAS" if Y, "NOT VAS" if N)

**Page-Level Data** (specific to each DO #):

- DO # (formatted with leading zeros, e.g., "0066455734")
- Ship To
- PO (all unique POs for this DO #, newline-separated)
- Original Qty (sum for this DO # with "Units" suffix, e.g., "7782 Units")

## Error Handling

- **Missing files**: Clear error messages for missing Excel or template files
- **Invalid data**: Validates shipment numbers and DO # formats
- **Missing columns**: Reports which required columns are missing
- **File permissions**: Handles read/write permission errors
- **Continues processing**: Logs errors but continues with next shipment

## Performance Features

- **Single Excel read**: File loaded once, not per shipment
- **Vectorized operations**: Uses pandas for efficient data filtering
- **Memory-based processing**: All operations work with in-memory data
- **Batch processing**: Can process multiple shipments in one session

## Recent Improvements

### Data Formatting Enhancements

- **DO # Leading Zeros**: Automatically formats DO # with leading zeros (e.g., "66455734" → "0066455734")
- **Units Suffix**: Adds "Units" to all quantity values (e.g., "7782" → "7782 Units")
- **Clean Shipment Numbers**: Removes decimal points from shipment numbers (e.g., "9010157586.0" → "9010157586")

### Multi-Page Formatting Fix

- **Complete Formatting Preservation**: All pages now maintain the exact same formatting as the template
- **Font Consistency**: Font names, sizes, and styles preserved across all pages
- **Professional Appearance**: Bold titles, proper spacing, and formatting maintained throughout the document
- **Template Fidelity**: Each page looks identical to the original template with only data values changed

## Output

Generated documents are saved as:

- **Location**: `Placards/` folder
- **Filename**: `Placard_[ShipmentNumber].docx`
- **Content**: Multi-page document with one page per DO #
- **Formatting**: Preserves complete template formatting across all pages (fonts, sizes, bold, etc.)
- **Data Format**: DO # with leading zeros, quantities with "Units" suffix, clean shipment numbers

## Troubleshooting

### Common Issues

#### "No Excel file found"

- Check that file is in `Data/` folder
- Verify filename starts with "WM-SPN-CUS105 Open Order Report"
- Ensure file is .xlsx or .xls format

#### "Template file not found"

- Verify `placard_template.docx` exists in `Template/` folder
- Check spelling and capitalization

#### "Missing required columns"

- Excel file must contain all required columns (case-sensitive)
- Check column names match exactly

#### "Invalid shipment number format"

- Shipment numbers must be exactly 10 digits
- No letters, spaces, or special characters

#### "No data found for shipment"

- Shipment number doesn't exist in Excel data
- Check if data was filtered out during validation

### Getting Help

1. Check that all required files are in correct locations
2. Verify Excel data contains required columns
3. Ensure shipment numbers are valid format
4. Check file permissions for read/write access

## Example Session

```text
=== Shipping Placard Generator ===
Loading data and preparing system...
Loading file: Data/WM-SPN-CUS105 Open Order Report 2024.xlsx
Loaded 1500 rows from Excel file
Removed 25 rows with empty Shipment Nbr
Removed 10 rows with invalid DO # format
Final dataset: 1465 rows ready for processing

Data loaded successfully! Ready to generate placards.

Enter one or more Shipment Numbers (comma-separated): 1234567890

Processing shipment: 1234567890
Found 15 records for shipment 1234567890
Processing 3 DO #s for shipment 1234567890
  Processing DO # 66455734 (1/3)
  Processing DO # 66455735 (2/3)  
  Processing DO # 66455736 (3/3)
SUCCESS: Created placard document: Placards/Placard_1234567890.docx

=== Processing Summary ===
Documents created: 1
Failed inputs: 0

Process more shipments? (y/n): n

=== Final Summary ===
Total documents created: 1
Total failed inputs: 0
Thank you for using the Shipping Placard Generator!
