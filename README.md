# Logistics Document Generator

A professional Python application that generates multi-page shipping placards from Excel data using Word templates. Features both a modern GUI interface and command-line operation, with enterprise-grade logging, bulk processing capabilities, and advanced document formatting preservation.

## 🚀 Key Features

### 🖥️ Modern GUI Interface (NEW!)

- **Interactive Data Table**: Visual shipment selection with real-time filtering and search
- **Advanced Filtering**: Multi-column filters with search and sort capabilities
- **Real-time Progress**: Visual progress tracking with detailed console logging
- **Professional Styling**: Modern dark theme with consistent button styling and layout
- **Error Handling**: Comprehensive error reporting and recovery mechanisms
- **Cross-platform**: Works on Windows, macOS, and Linux with automatic font detection

### 📊 Core Processing Features

- **Dual Processing Modes**: Manual entry for specific shipments or bulk processing for entire datasets
- **Advanced Template Engine**: Preserves complex Word formatting across all generated documents
- **Enterprise Logging**: Comprehensive CSV audit trail with session tracking and performance metrics
- **Data Validation**: Robust input validation and error handling throughout processing pipeline
- **High Performance**: Memory-efficient processing with pandas vectorization for large datasets

## 📋 Technical Requirements

- **Python**: 3.12.11+ (recommended for optimal performance)
- **Dependencies**: pandas, python-docx, openpyxl, dearpygui
- **Platform**: Cross-platform (Windows, macOS, Linux)
- **Memory**: Sufficient for Excel dataset processing (typically <100MB)

## 🛠 Installation & Setup

### Quick Start with Conda (Recommended)

```bash
# Create and activate environment
conda env create -f environment.yml
conda activate logistics-doc-generator

# Run GUI application (recommended)
python placard_generator_gui.py

# Or run command-line version
python placard_generator.py
```

### Alternative: pip Installation

```bash
pip install -r requirements.txt

# Run GUI application
python placard_generator_gui.py

# Or run command-line version
python placard_generator.py
```

## 📁 Project Structure

```text
Logistics Document Generator/
├── Data/                       # Excel data files (input)
├── Template/                   # Word template files
│   └── placard_template.docx
├── Placards/                  # Generated documents (auto-created)
├── Logs/                      # CSV audit logs (auto-created)
├── placard_generator_gui.py   # Modern GUI application (NEW!)
├── placard_generator.py       # Command-line application
├── environment.yml            # Conda environment
└── requirements.txt           # Python dependencies
```

## 🖥️ GUI Application Guide

### Getting Started with the GUI

1. **Launch the Application**

   ```bash
   python placard_generator_gui.py
   ```

2. **Load Your Data**
   - Click "LOAD DATA" button
   - Application automatically finds Excel files in the `Data/` folder
   - View loaded shipments in the interactive table

3. **Filter and Select Shipments**
   - Use the search box for quick filtering
   - Apply column-specific filters for precise selection
   - Use "SELECT ALL" or "DESELECT ALL" for bulk operations
   - Individual checkboxes for granular control

4. **Generate Documents**
   - Click "GENERATE SELECTED" for chosen shipments
   - Click "GENERATE ALL" for complete dataset processing
   - Monitor progress with the real-time progress bar
   - Review results in the console log

### GUI Features Overview

#### 🎛️ Control Bar

- **LOAD DATA**: Load Excel data from the Data folder
- **CLEAR FILTERS**: Reset all active filters
- **Search Box**: Real-time search across shipments and destinations
- **SELECT ALL / DESELECT ALL**: Bulk selection controls

#### 📊 Data Table

- **Interactive Selection**: Click checkboxes to select individual shipments
- **Column Sorting**: Click headers to sort data
- **Real-time Filtering**: Instantly see filtered results
- **Comprehensive Data View**: All shipment details in organized columns

#### 📈 Status and Progress

- **Selection Counter**: Shows selected vs. total shipments
- **Unit Counter**: Displays total quantity for selected items
- **Progress Bar**: Real-time processing progress
- **Status Messages**: Clear feedback on all operations

#### 🖥️ Console Log

- **Real-time Logging**: All operations logged with timestamps
- **Error Reporting**: Detailed error messages and stack traces
- **Performance Metrics**: Processing rates and timing information
- **Clear Console**: Reset log for new operations

### Advanced GUI Features

#### 🔍 Multi-Column Filtering

- Click column headers to access advanced filters
- Search within specific columns
- Sort filter options alphabetically
- Select multiple values per column
- Combine filters across columns for precise results

#### ⚡ Performance Optimizations

- **Lazy Loading**: Efficient memory usage for large datasets
- **Safe Operations**: Comprehensive error handling prevents crashes
- **Cross-platform Fonts**: Automatic font detection and fallbacks
- **Responsive UI**: Smooth operation even with large datasets

#### 🛡️ Error Handling

- **Data Validation**: Comprehensive checks before processing
- **File System Checks**: Validates directories and permissions
- **Processing Recovery**: Continues processing even if individual shipments fail
- **User Feedback**: Clear error messages and recovery suggestions

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

## 🎯 Usage Examples

### GUI Workflow Example

```text
1. Launch GUI: python placard_generator_gui.py
2. Click "LOAD DATA" → System finds and loads Excel file
3. Use search: "Chicago" → Filters to Chicago shipments
4. Select specific shipments using checkboxes
5. Click "GENERATE SELECTED" → Progress bar shows processing
6. Review results in console log
7. Find generated documents in Placards/ folder
```

### Command-Line Usage

#### 1. Manual Entry

- Process specific shipment numbers
- Supports batch input (comma-separated)
- Ideal for selective processing

#### 2. Bulk Processing

- Processes all valid shipments automatically
- Shows progress tracking every 10 shipments
- Requires confirmation before starting
- Perfect for complete dataset processing

### Example Command-Line Session

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

### Real-time Interface Features

- **GUI Console**: Real-time timestamped logging with color-coded messages
- **Progress Tracking**: Visual progress bars with detailed status updates
- **Performance Monitoring**: Processing rates and timing metrics
- **Error Recovery**: Comprehensive error handling with user-friendly messages

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

### GUI Customization

The GUI features a professional dark theme with:

- **Solid Color Design**: Consistent dark blue-gray background
- **Standardized Buttons**: Uniform styling with proper text centering
- **Responsive Layout**: Automatic centering and scaling
- **Cross-platform Fonts**: Automatic detection of system fonts
- **Accessibility**: High contrast colors and readable text

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
- **GUI won't start**: Ensure all dependencies are installed (`pip install -r requirements.txt`)

**Processing Issues**:

- **Invalid shipment format**: Must be exactly 10 digits
- **No data found**: Verify shipment exists and meets DO # validation (8+ digits)
- **File save errors**: Check write permissions and ensure files aren't open elsewhere
- **GUI freezing**: Check console log for detailed error messages

### GUI-Specific Troubleshooting

**Display Issues**:

- **Fonts not loading**: GUI automatically detects and uses system fonts
- **Layout problems**: Try resizing the window to trigger re-centering
- **Table not updating**: Click "LOAD DATA" to refresh the data

**Performance Issues**:

- **Slow loading**: Large Excel files may take time to process
- **Memory usage**: Close other applications if processing large datasets
- **UI responsiveness**: Check console for background processing status

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
- **GUI Optimizations**: Lazy loading and safe operations prevent UI freezing

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

Built using object-oriented design with comprehensive error handling:

### GUI Architecture

- **Dear PyGui Framework**: Modern, fast GUI with professional styling
- **Threaded Processing**: Background processing prevents UI freezing
- **Safe Operations**: Comprehensive error handling for all GUI operations
- **Cross-platform Compatibility**: Works on Windows, macOS, and Linux

### Core Engine (`PlacardGenerator` class)

- **Memory-efficient processing** with pandas vectorized operations
- **Advanced formatting preservation** across document generations
- **Robust error handling** with detailed validation and user feedback
- **Batch processing capabilities** for enterprise-scale operations
- **Comprehensive logging** for audit trails and performance monitoring

## 🔄 Development History

### Version 2.0.0 - GUI Release (Current)

**Major New Features**:

1. **Professional GUI Interface**: Modern Dear PyGui-based interface with dark theme
2. **Interactive Data Management**: Visual table with filtering, searching, and selection
3. **Real-time Processing**: Progress bars and live console logging
4. **Enhanced Error Handling**: Comprehensive error recovery and user feedback
5. **Cross-platform Support**: Automatic font detection and platform compatibility
6. **Performance Optimizations**: Safe operations and memory-efficient processing

### Version 1.0.0 - Command Line Foundation

**Core Enhancements**:

1. **Bulk Processing System**: Complete dataset processing with progress tracking
2. **Enterprise CSV Logging**: 11-column audit trail with session management
3. **Timestamped Interface**: Professional console output with real-time tracking
4. **Advanced Data Validation**: Comprehensive input validation and error handling
5. **Performance Optimization**: Memory-efficient processing with pandas vectorization
6. **Professional Documentation**: Enterprise-ready documentation and user guides

### Impact

**Transformation**: From basic document generator → Professional enterprise solution with modern GUI

**Current Capabilities**:

- ✅ Modern GUI interface with professional styling
- ✅ Interactive data management and real-time filtering
- ✅ Bulk processing with visual progress tracking
- ✅ Complete audit trail and compliance logging
- ✅ Cross-platform compatibility with automatic font detection
- ✅ Enterprise error handling and recovery mechanisms
- ✅ Performance monitoring and optimization
- ✅ Comprehensive documentation and user support

---

**Ready for Enterprise Use**: Full compliance capabilities, modern GUI interface, comprehensive error handling, and complete documentation for production environments.

**Recommended Usage**: Use the GUI interface (`placard_generator_gui.py`) for interactive work and the command-line interface (`placard_generator.py`) for automation and scripting.
