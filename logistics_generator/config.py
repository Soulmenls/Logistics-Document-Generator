#!/usr/bin/env python3
"""
Configuration File for Logistics Document Generator

Centralized configuration for security settings, performance limits,
and application behavior.
"""

import os
from pathlib import Path

# Application Information
APP_NAME = "Logistics Document Generator"
APP_VERSION = "2.1.0"
APP_AUTHOR = "Logistics Team"

# Security Configuration
SECURITY = {
    # File size limits (in bytes)
    'MAX_EXCEL_FILE_SIZE': 100 * 1024 * 1024,  # 100 MB
    'MAX_TEMPLATE_FILE_SIZE': 10 * 1024 * 1024,  # 10 MB
    'MAX_OUTPUT_FILE_SIZE': 50 * 1024 * 1024,   # 50 MB
    
    # Processing limits
    'MAX_RECORDS_PER_BATCH': 10000,
    'MAX_PROCESSING_TIME': 3600,  # 1 hour in seconds
    'MAX_OPERATIONS_PER_MINUTE': 60,
    
    # Text field limits
    'MAX_TEXT_FIELD_LENGTH': 1000,
    'MAX_FILENAME_LENGTH': 255,
    
    # Allowed file extensions
    'ALLOWED_EXCEL_EXTENSIONS': {'.xlsx', '.xls'},
    'ALLOWED_TEMPLATE_EXTENSIONS': {'.docx'},
    'ALLOWED_OUTPUT_EXTENSIONS': {'.docx'},
    
    # Directory security
    'RESTRICT_TO_WORKSPACE': True,
    'ALLOW_ABSOLUTE_PATHS': False,
}

# Performance Configuration
PERFORMANCE = {
    # Memory management
    'MAX_CONSOLE_LINES': 50,
    'MEMORY_CLEANUP_INTERVAL': 30,  # seconds
    'MAX_MEMORY_USAGE_MB': 512,     # Maximum memory usage
    
    # Threading
    'MAX_WORKER_THREADS': 4,
    'THREAD_TIMEOUT': 300,  # 5 minutes
    
    # GUI performance
    'TABLE_REFRESH_INTERVAL': 0.1,  # seconds
    'PROGRESS_UPDATE_INTERVAL': 0.05,  # seconds
}

# Logging Configuration
LOGGING = {
    'LEVEL': 'INFO',
    'FORMAT': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    'MAX_LOG_SIZE_MB': 10,
    'BACKUP_COUNT': 5,
    'LOG_TO_FILE': True,
    'LOG_TO_CONSOLE': True,
    'SECURITY_LOG_LEVEL': 'WARNING',
}

# Directory Configuration
DIRECTORIES = {
    'DATA_FOLDER': 'Data',
    'TEMPLATE_FOLDER': 'Template',
    'OUTPUT_FOLDER': 'Placards',
    'LOG_FOLDER': 'Logs',
    'TEMP_FOLDER': 'temp',
}

# GUI Configuration
GUI = {
    'WINDOW_WIDTH': 1600,
    'WINDOW_HEIGHT': 1200,
    'MIN_WIDTH': 1400,
    'MIN_HEIGHT': 900,
    'THEME': 'dark',
    'FONT_SIZE': 14,
    'ENABLE_CONSOLE': True,
    'AUTO_SAVE_SETTINGS': True,
}

# Validation Rules
VALIDATION = {
    'SHIPMENT_NUMBER_LENGTH': 10,
    'DO_NUMBER_MIN_LENGTH': 8,
    'DO_NUMBER_MAX_LENGTH': 15,
    'MIN_QUANTITY': 0,
    'MAX_QUANTITY': 1000000,
    'REQUIRED_COLUMNS': [
        'Shipment Nbr', 'DO #', 'Label Type', 'Order Type', 
        'Pmt Term', 'Start Ship', 'VAS', 'Ship To', 'PO', 'Original Qty'
    ],
    'EXCEL_FILE_PREFIX': 'WM-SPN-CUS105 Open Order Report',
    'TEMPLATE_FILENAME': 'placard_template.docx',
}

# Error Messages
ERROR_MESSAGES = {
    'FILE_NOT_FOUND': 'Required file not found: {filename}',
    'INVALID_FILE_FORMAT': 'Invalid file format: {filename}',
    'SECURITY_VIOLATION': 'Security violation detected: {details}',
    'PROCESSING_TIMEOUT': 'Processing timeout exceeded',
    'MEMORY_LIMIT_EXCEEDED': 'Memory limit exceeded',
    'INVALID_DATA': 'Invalid data detected: {details}',
    'PERMISSION_DENIED': 'Permission denied: {operation}',
}

def get_workspace_path() -> Path:
    """Get the workspace root path"""
    return Path(__file__).parent.absolute()

def get_data_path() -> Path:
    """Get the data directory path"""
    return get_workspace_path() / DIRECTORIES['DATA_FOLDER']

def get_template_path() -> Path:
    """Get the template directory path"""
    return get_workspace_path() / DIRECTORIES['TEMPLATE_FOLDER']

def get_output_path() -> Path:
    """Get the output directory path"""
    return get_workspace_path() / DIRECTORIES['OUTPUT_FOLDER']

def get_log_path() -> Path:
    """Get the log directory path"""
    return get_workspace_path() / DIRECTORIES['LOG_FOLDER']

def ensure_directories():
    """Ensure all required directories exist"""
    for dir_name in DIRECTORIES.values():
        dir_path = get_workspace_path() / dir_name
        dir_path.mkdir(exist_ok=True, mode=0o755)

# Environment-specific overrides
def load_environment_config():
    """Load environment-specific configuration overrides"""
    env = os.getenv('LOGISTICS_ENV', 'production').lower()
    
    if env == 'development':
        LOGGING['LEVEL'] = 'DEBUG'
        SECURITY['MAX_RECORDS_PER_BATCH'] = 1000  # Smaller for testing
        PERFORMANCE['MAX_MEMORY_USAGE_MB'] = 256
    elif env == 'testing':
        LOGGING['LEVEL'] = 'WARNING'
        SECURITY['MAX_RECORDS_PER_BATCH'] = 100
        PERFORMANCE['MAX_MEMORY_USAGE_MB'] = 128
        GUI['ENABLE_CONSOLE'] = False

# Load environment config on import
load_environment_config() 