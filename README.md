# Logistics Document Generator

A professional Python application that generates multi-page shipping placards from Excel data using Word templates. Features enterprise-grade logging, bulk processing capabilities, and advanced document formatting preservation.

## 🚀 Key Features

- **Dual Processing Modes**: Manual entry for specific shipments or bulk processing for entire datasets
- **Advanced Template Engine**: Preserves complex Word formatting across all generated documents
- **Enterprise Logging**: Comprehensive CSV audit trail with session tracking and performance metrics
- **Real-time Interface**: Timestamped console output with professional user experience
- **Data Validation**: Robust input validation and error handling throughout processing pipeline
- **High Performance**: Memory-efficient processing with pandas vectorization for large datasets

## 📋 Technical Requirements

- **Python**: 3.12.11+ (recommended for optimal performance)
- **Dependencies**: pandas, python-docx, openpyxl
- **Platform**: Cross-platform (Windows, macOS, Linux)
- **Memory**: Sufficient for Excel dataset processing (typically <100MB)

## 🛠 Installation & Setup

### Quick Start with Conda (Recommended)

```bash
# Create and activate environment
conda env create -f environment.yml
conda activate logistics-doc-generator

# Run application
python placard_generator.py
```

### Alternative: pip Installation

```bash
pip install -r requirements.txt
python placard_generator.py
```

## 📁 Project Structure

```text
Logistics Document Generator/
├── Data/                   # Excel data files (input)
├── Template/               # Word template files
│   └── placard_template.docx
├── Placards/              # Generated documents (auto-created)
├── Logs/                  # CSV audit logs (auto-created)
├── placard_generator.py   # Main application
├── environment.yml        # Conda environment
└── requirements.txt       # Python dependencies
```

## 📊 Data Requirements

### Excel Input File

**Location**: `Data/` folder  
**Naming**: Must start with `"WM-SPN-CUS105 Open Order Report"`  
**Format**: `.xlsx` or `.xls`

**Required Columns**:

| Column | Description | Format |
|--------|-------------|---------|
| `Shipment Nbr` | Shipment identifier | Exactly 10 digits |
| `DO #` | Delivery Order number | Minimum 8 digits |
| `Label Type` | Shipment classification | Text |
| `Order Type` | Order classification | Text |
| `Pmt Term` | Payment terms | Text |
| `Start Ship` | Ship date | Any date format (auto-converted to MM/DD/YYYY) |
| `VAS` | Value Added Service | Y/N (converted to "VAS"/"NOT VAS") |
| `Ship To` | Destination information | Text |
| `PO` | Purchase Order numbers | Text (multiple POs aggregated) |
| `Original Qty` | Quantity values | Numeric (auto-formatted with "Units" suffix) |

### Word Template

**Location**: `Template/placard_template.docx`

**Required Placeholders**:

```text
{{Ship To}}        - Destination address
{{Shipment Nbr}}   - Shipment number (cleaned)
{{PO}}             - Purchase orders (newline-separated)
{{DO #}}           - Delivery order (10-digit format)
{{VAS}}            - Value added service status
{{Original Qty}}   - Total quantity with "Units"
{{Label Type}}     - Shipment classification
{{Order Type}}     - Order classification
{{Pmt Term}}       - Payment terms
{{Start Ship}}     - Formatted ship date
```

## 🎯 Usage

### Processing Options

#### 1. Manual Entry

- Process specific shipment numbers
- Supports batch input (comma-separated)
- Ideal for selective processing

#### 2. Bulk Processing

- Processes all valid shipments automatically
- Shows progress tracking every 10 shipments
- Requires confirmation before starting
- Perfect for complete dataset processing

### Example Session

```text
[2024-01-15 14:30:25] === Shipping Placard Generator ===
[2024-01-15 14:30:26] Loading file: Data/WM-SPN-CUS105 Open Order Report 2024.xlsx
[2024-01-15 14:30:26] Final dataset: 1465 rows ready for processing
[2024-01-15 14:30:26] Dataset contains 266 unique valid shipments.

Choose an option:
1. Enter specific shipment numbers
2. Generate placards for ALL shipments in dataset
3. Exit

Enter your choice (1-3): 1
Enter Shipment Numbers: 1234567890, 9876543210

[2024-01-15 14:30:45] Processing shipment: 1234567890
[2024-01-15 14:30:52] SUCCESS: Created Placards/Placard_1234567890.docx
[2024-01-15 14:31:15] SUCCESS: Created Placards/Placard_9876543210.docx

=== Processing Summary ===
Documents created: 2
Failed inputs: 0
```

## 📈 Enterprise Features

### Comprehensive CSV Logging

**Automatic Audit Trail**: Every operation logged to timestamped CSV files in `Logs/` folder

**Event Types Tracked**:

- Session lifecycle (start/end with duration)
- Data loading and validation results
- Processing mode selections
- Individual shipment processing details
- Bulk processing progress and summaries
- Error tracking with detailed messages

**Log Structure** (11 columns):

```csv
Timestamp, Session_ID, Event_Type, Shipment_Number, DO_Count, 
Records_Found, Status, Output_File, Error_Message, Processing_Mode, Duration_Seconds
```

### Real-time Timestamped Interface

- All console output includes `[YYYY-MM-DD HH:MM:SS]` timestamps
- Professional interface suitable for enterprise environments
- Real-time progress tracking and status updates
- Performance monitoring and debugging capabilities

### Advanced Data Processing

**Data Transformation Pipeline**:

1. **File Discovery**: Automatic Excel file detection
2. **Data Loading**: Pandas-powered efficient data reading
3. **Data Cleaning**: Removes empty/invalid records
4. **Validation**: Comprehensive format checking
5. **Memory Storage**: Fast in-memory processing

**Document Generation**:

1. **Data Grouping**: Groups records by DO # using pandas
2. **Data Aggregation**: Combines POs, sums quantities per DO #
3. **Template Processing**: Advanced placeholder replacement
4. **Formatting Preservation**: Maintains all Word formatting
5. **Multi-page Assembly**: Creates separate pages per DO #

## 🔧 Advanced Configuration

### Formatting Preservation

The application uses sophisticated techniques to preserve Word document formatting:

- **Run-level formatting**: Fonts, sizes, colors, styles
- **Paragraph formatting**: Alignment, spacing, line spacing
- **Mixed formatting**: Different styles within single paragraphs
- **Cross-run placeholders**: Handles complex placeholder placement
- **Table formatting**: Preserves table styles and cell formatting

### Data Formatting Rules

**Automatic Transformations**:

- `Shipment Nbr`: Removes decimal places (e.g., `1234567890.0` → `1234567890`)
- `DO #`: Padded with leading zeros to 10 digits (e.g., `66455734` → `0066455734`)
- `Start Ship`: Converted to MM/DD/YYYY format
- `VAS`: Converted to "VAS" or "NOT VAS"
- `Original Qty`: Summed with "Units" suffix (e.g., `"7782 Units"`)
- `PO`: Multiple POs joined with newlines

## 🚨 Error Handling & Troubleshooting

### Common Issues

**Setup Problems**:

- **No Excel file found**: Verify file is in `Data/` folder with correct naming
- **Template missing**: Confirm `placard_template.docx` exists in `Template/` folder
- **Missing columns**: Check all 10 required columns exist (case-sensitive)

**Processing Issues**:

- **Invalid shipment format**: Must be exactly 10 digits
- **No data found**: Verify shipment exists and meets DO # validation (8+ digits)
- **File save errors**: Check write permissions and ensure files aren't open elsewhere

### Validation Rules

- **Shipment Numbers**: Exactly 10 digits, no letters/special characters
- **DO # Format**: Minimum 8 digits, validated with regex
- **File Permissions**: Automatic handling of read/write access issues
- **Data Existence**: Validates shipment exists in filtered dataset

## 📊 Performance Optimizations

- **Single File Read**: Excel loaded once at startup for all operations
- **Vectorized Operations**: Pandas operations for efficient data processing
- **Memory-based Processing**: All operations use in-memory datasets
- **Template Reuse**: Efficient document copying with formatting preservation
- **Batch Processing**: Multiple shipments processed in single session

## 🎯 Output Specifications

**Generated Documents**:

- **Location**: `Placards/` folder (auto-created)
- **Naming**: `Placard_[ShipmentNumber].docx`
- **Structure**: Multi-page document (one page per DO #)
- **Formatting**: Complete template formatting preserved

**Quality Assurance**:

- DO # with leading zeros (10 digits total)
- Quantities with "Units" suffix
- Clean shipment numbers (no decimals)
- Properly formatted dates (MM/DD/YYYY)
- All template placeholders replaced accurately

## 🏗 Architecture Overview

Built using object-oriented design with the `PlacardGenerator` class:

- **Memory-efficient processing** with pandas vectorized operations
- **Advanced formatting preservation** across document generations
- **Robust error handling** with detailed validation and user feedback
- **Batch processing capabilities** for enterprise-scale operations
- **Comprehensive logging** for audit trails and performance monitoring
- **Professional user interface** with real-time feedback and timestamps

## 🔄 Development History

This application has evolved from a basic document generator into a professional, enterprise-ready logistics solution with the following major enhancements:

### Core Enhancements

1. **Bulk Processing System**: Complete dataset processing with progress tracking
2. **Enterprise CSV Logging**: 11-column audit trail with session management
3. **Timestamped Interface**: Professional console output with real-time tracking
4. **Advanced Data Validation**: Comprehensive input validation and error handling
5. **Performance Optimization**: Memory-efficient processing with pandas vectorization
6. **Professional Documentation**: Enterprise-ready documentation and user guides
7. **Repository Management**: Professional git integration with data privacy considerations

### Impact

**Transformation**: From basic single-shipment tool → Professional enterprise solution

**Capabilities Added**:

- ✅ Bulk processing (266+ shipments automatically)
- ✅ Complete audit trail and compliance logging
- ✅ Real-time timestamped user interface
- ✅ Enterprise error handling and resilience
- ✅ Professional documentation and support
- ✅ Performance monitoring and optimization

---

**Ready for Enterprise Use**: Full compliance capabilities, professional interface, comprehensive error handling, and complete documentation for production environments.
