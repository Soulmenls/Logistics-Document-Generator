# Logistics Document Generator

A professional Python application that generates multi-page shipping placards from Excel data using Word templates. Features both a modern GUI interface and command-line operation, with **enterprise-grade security**, comprehensive logging, bulk processing capabilities, and advanced document formatting preservation.

## 🚀 Key Features

### 🔒 Enterprise Security (NEW!)

- **Comprehensive Input Validation**: Advanced validation for all data fields with malicious content detection
- **Path Traversal Protection**: Secure file operations preventing directory traversal attacks
- **File Integrity Verification**: SHA-256 hashing for file integrity and security audit trails
- **Rate Limiting**: Configurable operation limits to prevent resource exhaustion attacks
- **Memory Management**: Automatic memory cleanup and leak prevention with circular buffers
- **Security Logging**: Detailed security event logging with audit trails and threat detection
- **Dependency Security**: Pinned dependency versions and security-focused package management

### 🖥️ Modern GUI Interface

- **Interactive Data Table**: Visual shipment selection with real-time filtering and search
- **Advanced Filtering**: Multi-column filters with search and sort capabilities
- **Real-time Progress**: Visual progress tracking with detailed console logging
- **Professional Styling**: Modern dark theme with consistent button styling and layout
- **Memory-Safe Operations**: Comprehensive memory management preventing leaks and crashes
- **Thread-Safe Processing**: Secure multi-threaded operations with proper synchronization
- **Cross-platform**: Works on Windows, macOS, and Linux with automatic font detection

### 📊 Core Processing Features

- **Dual Processing Modes**: Manual entry for specific shipments or bulk processing for entire datasets
- **Advanced Template Engine**: Preserves complex Word formatting across all generated documents
- **Enterprise Logging**: Comprehensive CSV audit trail with session tracking and performance metrics
- **Robust Data Validation**: Multi-layered validation system with security-focused input checking
- **High Performance**: Memory-efficient processing with pandas vectorization for large datasets
- **Secure File Handling**: Protected file operations with validation and integrity checks

## 🛡️ Security Features

### Input Validation & Sanitization

- **Shipment Number Validation**: Flexible 8-12 digit validation supporting various formats
- **DO Number Validation**: Secure validation for 6-15 digit delivery order numbers
- **Text Field Sanitization**: XSS prevention and malicious content detection
- **Numeric Field Validation**: Range checking and type validation for all numeric inputs
- **File Path Sanitization**: Prevention of directory traversal and path manipulation attacks

### Security Architecture

- **Multi-layered Validation**: Input validation at multiple processing stages
- **Secure File Operations**: All file operations use validated, sanitized paths
- **Resource Protection**: Rate limiting and processing timeouts prevent resource exhaustion
- **Memory Safety**: Automatic cleanup and leak prevention with monitoring
- **Audit Logging**: Comprehensive security event logging for compliance and monitoring

### Threat Protection

- **Path Traversal Prevention**: Secure file handling prevents directory traversal attacks
- **Input Sanitization**: Protection against script injection and malicious content
- **Resource Exhaustion Protection**: Rate limiting and processing limits prevent DoS attacks
- **File Integrity Verification**: SHA-256 hashing ensures file integrity and detects tampering
- **Dependency Security**: Pinned versions and security-focused package management

## 📋 Technical Requirements

- **Python**: 3.12.11+ (recommended for optimal performance and security)
- **Dependencies**: pandas, python-docx, openpyxl, dearpygui (all versions pinned for security)
- **Platform**: Cross-platform (Windows, macOS, Linux)
- **Memory**: Sufficient for Excel dataset processing (typically <100MB, with monitoring)
- **Security**: File system permissions for secure file operations

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
├── placard_generator_gui.py   # Modern GUI application
├── placard_generator.py       # Command-line application
├── security_utils.py          # Security utilities and validation (NEW!)
├── config.py                  # Configuration and security settings (NEW!)
├── SECURITY.md               # Security documentation (NEW!)
├── environment.yml            # Conda environment
└── requirements.txt           # Python dependencies (security-pinned)
```

## 🔒 Security Configuration

### Security Settings (`config.py`)

The application includes comprehensive security configuration:

```python
# File size limits
MAX_EXCEL_FILE_SIZE = 100 * 1024 * 1024  # 100 MB
MAX_TEMPLATE_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

# Processing limits
MAX_RECORDS_PER_BATCH = 10000
MAX_PROCESSING_TIME = 3600  # 1 hour

# Rate limiting
MAX_OPERATIONS_PER_MINUTE = 60

# Allowed file extensions
ALLOWED_EXCEL_EXTENSIONS = {'.xlsx', '.xls'}
ALLOWED_TEMPLATE_EXTENSIONS = {'.docx'}
```

### Validation Rules

- **Shipment Numbers**: 8-12 digits, handles float format (e.g., `9010157586.0`)
- **DO Numbers**: 6-15 digits, required field with format validation
- **Text Fields**: Length limits, malicious content detection, XSS prevention
- **File Paths**: Directory traversal prevention, path sanitization
- **File Integrity**: SHA-256 verification for all processed files

## 🖥️ GUI Application Guide

### Getting Started with the GUI

1. **Launch the Application**

   ```bash
   python placard_generator_gui.py
   ```

2. **Load Your Data**
   - Click "LOAD DATA" button
   - Application automatically finds and validates Excel files in the `Data/` folder
   - Security validation ensures file integrity and prevents malicious files
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
   - Review results in the secure console log

### Enhanced GUI Security Features

#### 🛡️ Memory Management

- **Circular Buffers**: Console logs use `deque` with maximum size to prevent memory leaks
- **Automatic Cleanup**: Periodic garbage collection and memory monitoring
- **Resource Monitoring**: Real-time memory usage tracking and alerts
- **Thread Safety**: All operations use proper locking mechanisms

#### 🔒 Secure Operations

- **Input Validation**: All user inputs validated before processing
- **File Validation**: Comprehensive file integrity checks before processing
- **Rate Limiting**: Prevents resource exhaustion through operation limiting
- **Error Recovery**: Secure error handling prevents information disclosure

## 📊 Data Requirements

### Excel Input File

**Location**: `Data/` folder  
**Naming**: Must start with `"WM-SPN-CUS105 Open Order Report"`  
**Format**: `.xlsx` or `.xls` (validated for security)
**Size Limit**: 100MB maximum for security

**Required Columns**:

| Column | Description | Validation |
|--------|-------------|------------|
| `Shipment Nbr` | Shipment identifier | 8-12 digits, handles float format |
| `DO #` | Delivery Order number | 6-15 digits, required field |
| `Label Type` | Shipment classification | Text, length limited |
| `Order Type` | Order classification | Text, XSS prevention |
| `Pmt Term` | Payment terms | Text, sanitized |
| `Start Ship` | Ship date | Date format validation |
| `VAS` | Value Added Service | Y/N validation |
| `Ship To` | Destination information | Text, malicious content detection |
| `PO` | Purchase Order numbers | Text, aggregated securely |
| `Original Qty` | Quantity values | Numeric validation |

### Word Template

**Location**: `Template/placard_template.docx`
**Size Limit**: 10MB maximum for security
**Integrity**: SHA-256 hash verification

**Required Placeholders**:

```text
{{Ship To}}        - Destination address (sanitized)
{{Shipment Nbr}}   - Shipment number (validated)
{{PO}}             - Purchase orders (securely aggregated)
{{DO #}}           - Delivery order (validated format)
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
2. Click "LOAD DATA" → System finds, validates, and loads Excel file
3. Security validation: File integrity check and malicious content scan
4. Use search: "Chicago" → Filters to Chicago shipments
5. Select specific shipments using checkboxes
6. Click "GENERATE SELECTED" → Secure processing with progress tracking
7. Review results in console log with security events
8. Find generated documents in Placards/ folder
```

### Command-Line Usage

#### 1. Manual Entry

- Process specific shipment numbers with validation
- Supports batch input (comma-separated)
- Comprehensive security checks for all inputs
- Ideal for selective processing

#### 2. Bulk Processing

- Processes all valid shipments with security validation
- Rate limiting prevents resource exhaustion
- Shows progress tracking every 10 shipments
- Requires confirmation before starting
- Perfect for complete dataset processing

### Example Command-Line Session

```text
[2024-01-15 14:30:25] === Shipping Placard Generator ===
[2024-01-15 14:30:26] Loading file: Data/WM-SPN-CUS105 Open Order Report 2024.xlsx
[2024-01-15 14:30:26] Security validation: File integrity verified (SHA256: abc123...)
[2024-01-15 14:30:26] Data validation passed for 1465 records
[2024-01-15 14:30:26] Final dataset: 1465 rows ready for processing
[2024-01-15 14:30:26] Dataset contains 266 unique valid shipments.

Choose an option:
1. Enter specific shipment numbers
2. Generate placards for ALL shipments in dataset
3. Exit

Enter your choice (1-3): 1
Enter Shipment Numbers: 9010157586, 9010157584

[2024-01-15 14:30:45] Processing shipment: 9010157586
[2024-01-15 14:30:52] SUCCESS: Created Placards/Placard_9010157586.docx
[2024-01-15 14:31:15] SUCCESS: Created Placards/Placard_9010157584.docx

=== Processing Summary ===
Documents created: 2
Failed inputs: 0
```

## 📈 Enterprise Features

### Comprehensive Security Logging

**Security Event Tracking**: All security-related events logged with timestamps

**Security Events Monitored**:

- File integrity verification (SHA-256 hashes)
- Input validation failures and malicious content detection
- Rate limiting violations and resource exhaustion attempts
- Path traversal attempts and file system security events
- Memory usage monitoring and cleanup operations
- Processing timeouts and resource limit violations

### Comprehensive CSV Logging

**Automatic Audit Trail**: Every operation logged to timestamped CSV files in `Logs/` folder

**Event Types Tracked**:

- Session lifecycle (start/end with duration)
- Data loading and validation results
- Security validation events and threat detection
- Processing mode selections
- Individual shipment processing details
- Bulk processing progress and summaries
- Error tracking with detailed messages and security context

**Log Structure** (Enhanced with security fields):

```csv
Timestamp, Session_ID, Event_Type, Shipment_Number, DO_Count, 
Records_Found, Status, Output_File, Error_Message, Processing_Mode, 
Duration_Seconds, Security_Hash, Validation_Status
```

### Real-time Interface Features

- **Secure GUI Console**: Real-time timestamped logging with security event highlighting
- **Progress Tracking**: Visual progress bars with security validation status
- **Performance Monitoring**: Processing rates, memory usage, and security metrics
- **Error Recovery**: Comprehensive error handling with security-aware recovery mechanisms

### Advanced Data Processing

**Secure Data Transformation Pipeline**:

1. **File Discovery**: Automatic Excel file detection with integrity verification
2. **Security Validation**: Comprehensive file and content security checks
3. **Data Loading**: Pandas-powered efficient data reading with validation
4. **Data Cleaning**: Secure removal of empty/invalid records
5. **Validation**: Multi-layered format and security checking
6. **Memory Storage**: Secure in-memory processing with monitoring

**Secure Document Generation**:

1. **Data Grouping**: Secure grouping of records by DO # using validated pandas operations
2. **Data Aggregation**: Secure combination of POs and quantity summation per DO #
3. **Template Processing**: Secure placeholder replacement with input sanitization
4. **Formatting Preservation**: Maintains all Word formatting with security validation
5. **Multi-page Assembly**: Creates separate pages per DO # with integrity checks

## 🔧 Advanced Configuration

### Security Configuration

The application includes comprehensive security settings in `config.py`:

```python
# Security limits
SECURITY_CONFIG = {
    'max_file_size': 100 * 1024 * 1024,  # 100MB
    'max_records': 10000,
    'rate_limit': 60,  # operations per minute
    'processing_timeout': 3600,  # 1 hour
    'memory_limit': 500 * 1024 * 1024,  # 500MB
}

# Validation rules
VALIDATION_RULES = {
    'shipment_number': r'^\d{8,12}$',
    'do_number': r'^\d{6,15}$',
    'text_max_length': 1000,
    'allow_empty_shipment': True,
}
```

### GUI Customization

The GUI features a professional dark theme with security-focused design:

- **Solid Color Design**: Consistent dark blue-gray background
- **Standardized Buttons**: Uniform styling with proper text centering
- **Responsive Layout**: Automatic centering and scaling
- **Cross-platform Fonts**: Automatic detection of system fonts
- **Accessibility**: High contrast colors and readable text
- **Security Indicators**: Visual feedback for security validation status

## 🚨 Error Handling & Troubleshooting

### Security-Related Issues

**File Security**:

- **File integrity failure**: SHA-256 hash verification failed - file may be corrupted or tampered
- **Path traversal detected**: Attempted directory traversal blocked - check file paths
- **File size exceeded**: File exceeds security limits - reduce file size or adjust limits
- **Malicious content detected**: Input contains potentially harmful content - review and sanitize

**Validation Failures**:

- **Rate limit exceeded**: Too many operations - wait before retrying
- **Invalid input format**: Input doesn't match security validation rules
- **Processing timeout**: Operation exceeded time limit - reduce dataset size
- **Memory limit exceeded**: Insufficient memory - close other applications

### Common Issues

**Setup Problems**:

- **No Excel file found**: Verify file is in `Data/` folder with correct naming
- **Template missing**: Confirm `placard_template.docx` exists in `Template/` folder
- **Missing columns**: Check all required columns exist (case-sensitive)
- **Security validation failed**: Ensure files meet security requirements

**Processing Issues**:

- **Invalid shipment format**: Must be 8-12 digits (flexible validation)
- **No data found**: Verify shipment exists and meets validation requirements
- **File save errors**: Check write permissions and ensure files aren't open elsewhere
- **Security blocking**: Check security logs for validation failures

### Enhanced Validation Rules

- **Shipment Numbers**: 8-12 digits, handles float format (e.g., `9010157586.0` → `9010157586`)
- **DO # Format**: 6-15 digits, validated with enhanced regex
- **File Permissions**: Automatic handling of read/write access issues with security logging
- **Data Existence**: Validates shipment exists in filtered dataset with integrity checks

## 📊 Performance & Security Optimizations

### Performance Features

- **Single File Read**: Excel loaded once at startup for all operations
- **Vectorized Operations**: Pandas operations for efficient data processing
- **Memory-based Processing**: All operations use secure in-memory datasets
- **Template Reuse**: Efficient document copying with formatting preservation
- **Batch Processing**: Multiple shipments processed in single session

### Security Optimizations

- **Memory Management**: Circular buffers prevent memory leaks
- **Resource Monitoring**: Real-time tracking of memory and CPU usage
- **Rate Limiting**: Prevents resource exhaustion attacks
- **Input Validation**: Multi-layered validation prevents malicious input
- **File Integrity**: SHA-256 verification ensures file security
- **Secure Operations**: All file operations use validated, sanitized paths

## 🎯 Output Specifications

**Generated Documents**:

- **Location**: `Placards/` folder (auto-created with secure permissions)
- **Naming**: `Placard_[ShipmentNumber].docx` (sanitized filenames)
- **Structure**: Multi-page document (one page per DO #)
- **Formatting**: Complete template formatting preserved with integrity checks
- **Security**: All content validated and sanitized before document generation

**Quality Assurance**:

- DO # with leading zeros (10 digits total)
- Quantities with "Units" suffix
- Clean shipment numbers (no decimals)
- Properly formatted dates (MM/DD/YYYY)
- All template placeholders replaced with validated content
- File integrity verification for all generated documents

## 🏗 Architecture Overview

Built using object-oriented design with comprehensive security and error handling:

### Application Security Architecture

- **Multi-layered Validation**: Input validation at multiple processing stages
- **Secure File Operations**: All file operations use `SecureFileHandler` class
- **Memory Management**: Automatic cleanup with `MemoryManager` utilities
- **Rate Limiting**: `RateLimiter` class prevents resource exhaustion
- **Audit Logging**: Comprehensive security event logging throughout

### GUI Architecture

- **Dear PyGui Framework**: Modern, fast GUI with professional styling
- **Thread-Safe Processing**: Background processing with proper synchronization
- **Memory-Safe Operations**: Circular buffers and automatic cleanup
- **Security Integration**: Real-time security validation and monitoring
- **Cross-platform Compatibility**: Works on Windows, macOS, and Linux

### Core Engine (`PlacardGenerator` class)

- **Memory-efficient processing** with secure pandas vectorized operations
- **Advanced formatting preservation** with security validation
- **Robust error handling** with security-aware validation and user feedback
- **Batch processing capabilities** with rate limiting for enterprise-scale operations
- **Comprehensive logging** with security event tracking for audit trails

## 🔄 Development History

### Version 3.0.0 - Security Hardening (Current)

**Major Security Enhancements**:

1. **Comprehensive Security Framework**: Complete security utilities module with validation, sanitization, and protection
2. **Input Validation System**: Multi-layered validation for all data inputs with malicious content detection
3. **File Security**: SHA-256 integrity verification, path traversal protection, and secure file operations
4. **Memory Management**: Automatic memory cleanup, leak prevention, and resource monitoring
5. **Rate Limiting**: Configurable operation limits to prevent resource exhaustion attacks
6. **Security Logging**: Detailed security event logging with audit trails and threat detection
7. **Dependency Security**: Pinned dependency versions and security-focused package management

**Security Fixes**:

- ✅ **CRITICAL**: Fixed path traversal vulnerabilities in file operations
- ✅ **HIGH**: Implemented secure document template processing with validation
- ✅ **HIGH**: Added comprehensive input validation for all data fields
- ✅ **MEDIUM**: Secured file system access with proper validation
- ✅ **HIGH**: Fixed memory leaks in GUI with circular buffers and cleanup
- ✅ **HIGH**: Resolved race conditions with proper thread synchronization
- ✅ **MEDIUM**: Enhanced exception handling with security-aware error messages

### Version 2.0.0 - GUI Release

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

**Transformation**: From basic document generator → **Enterprise-grade secure solution** with modern GUI and comprehensive security framework

**Current Security Capabilities**:

- ✅ **Enterprise Security**: Comprehensive security framework with threat protection
- ✅ **Input Validation**: Multi-layered validation with malicious content detection
- ✅ **File Security**: SHA-256 integrity verification and secure file operations
- ✅ **Memory Safety**: Automatic cleanup and leak prevention with monitoring
- ✅ **Rate Limiting**: Resource exhaustion protection with configurable limits
- ✅ **Audit Logging**: Complete security event tracking and compliance logging
- ✅ **Dependency Security**: Pinned versions and security-focused package management

## 🛡️ Security Compliance

For detailed security information, vulnerability reporting, and security best practices, see [SECURITY.md](SECURITY.md).

**Security Standards**:

- Input validation and sanitization
- File integrity verification
- Path traversal prevention
- Resource exhaustion protection
- Memory safety and leak prevention
- Comprehensive audit logging
- Secure dependency management

**Recommended Security Practices**:

- Regular security updates
- File integrity monitoring
- Access control and permissions
- Network security considerations
- Data handling and privacy protection
