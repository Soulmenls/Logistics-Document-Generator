#!/usr/bin/env python3
"""
Utility functions for the Logistics Document Generator package

This module provides common utility functions used across the package,
including logging setup, file operations, and data validation helpers.
"""

import os
import sys
import logging
import logging.handlers
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

from .config import LOGGING, DIRECTORIES


def setup_logging(log_level: Optional[str] = None, log_to_file: bool = True) -> None:
    """
    Setup logging configuration for the package
    
    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_to_file: Whether to log to file in addition to console
    """
    # Use config values or defaults
    level = log_level or LOGGING.get('LEVEL', 'INFO')
    format_str = LOGGING.get('FORMAT', '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # Convert string level to logging constant
    numeric_level = getattr(logging, level.upper(), logging.INFO)
    
    # Clear any existing handlers
    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    
    # Set root logger level
    root_logger.setLevel(numeric_level)
    
    # Create formatter
    formatter = logging.Formatter(format_str)
    
    # Console handler
    if LOGGING.get('LOG_TO_CONSOLE', True):
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(numeric_level)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
    
    # File handler
    if log_to_file and LOGGING.get('LOG_TO_FILE', True):
        try:
            # Ensure log directory exists
            log_dir = get_package_data_dir() / DIRECTORIES.get('LOG_FOLDER', 'Logs')
            log_dir.mkdir(parents=True, exist_ok=True)
            
            # Create rotating file handler
            log_file = log_dir / 'logistics_generator.log'
            max_bytes = LOGGING.get('MAX_LOG_SIZE_MB', 10) * 1024 * 1024
            backup_count = LOGGING.get('BACKUP_COUNT', 5)
            
            file_handler = logging.handlers.RotatingFileHandler(
                str(log_file),
                maxBytes=max_bytes,
                backupCount=backup_count,
                encoding='utf-8'
            )
            file_handler.setLevel(numeric_level)
            file_handler.setFormatter(formatter)
            root_logger.addHandler(file_handler)
            
        except Exception as e:
            # If file logging fails, at least log to console
            print(f"Warning: Could not setup file logging: {e}")


def get_package_root() -> Path:
    """Get the root directory of the package"""
    return Path(__file__).parent.absolute()


def get_package_data_dir() -> Path:
    """
    Get the data directory for the package
    
    This will be either the package installation directory for system installs,
    or a user data directory for user installs.
    """
    try:
        # Try to use the package directory first
        package_dir = get_package_root()
        
        # For development/editable installs, use the package directory
        if (package_dir.parent / 'setup.py').exists():
            return package_dir.parent
            
        # For system installs, create user data directory
        if sys.platform.startswith('win'):
            # Windows
            data_dir = Path(os.environ.get('APPDATA', Path.home())) / 'LogisticsGenerator'
        elif sys.platform.startswith('darwin'):
            # macOS
            data_dir = Path.home() / 'Library/Application Support/LogisticsGenerator'
        else:
            # Linux/Unix
            data_dir = Path.home() / '.local/share/LogisticsGenerator'
            
        data_dir.mkdir(parents=True, exist_ok=True)
        return data_dir
        
    except Exception:
        # Fallback to current working directory
        return Path.cwd()


def ensure_package_directories() -> Dict[str, Path]:
    """
    Ensure all required package directories exist
    
    Returns:
        Dictionary mapping directory names to Path objects
    """
    base_dir = get_package_data_dir()
    directories = {}
    
    for dir_key, dir_name in DIRECTORIES.items():
        dir_path = base_dir / dir_name
        dir_path.mkdir(parents=True, exist_ok=True, mode=0o755)
        directories[dir_key] = dir_path
    
    return directories


def get_template_path() -> Path:
    """Get the path to the template directory"""
    # First check if templates are in the package data
    package_templates = get_package_root() / 'templates'
    if package_templates.exists():
        return package_templates
    
    # Otherwise use the user data directory
    data_dir = get_package_data_dir()
    template_dir = data_dir / DIRECTORIES.get('TEMPLATE_FOLDER', 'Template')
    template_dir.mkdir(parents=True, exist_ok=True)
    return template_dir


def get_default_template() -> Optional[Path]:
    """Get the path to the default template file"""
    from .config import VALIDATION
    
    template_dir = get_template_path()
    template_name = VALIDATION.get('TEMPLATE_FILENAME', 'placard_template.docx')
    template_path = template_dir / template_name
    
    return template_path if template_path.exists() else None


def validate_package_installation() -> Dict[str, Any]:
    """
    Validate that the package is properly installed and configured
    
    Returns:
        Dictionary with validation results
    """
    results = {
        'is_valid': True,
        'errors': [],
        'warnings': [],
        'directories': {},
        'template_found': False,
    }
    
    try:
        # Check directories
        directories = ensure_package_directories()
        results['directories'] = {k: str(v) for k, v in directories.items()}
        
        # Check for template
        template_path = get_default_template()
        if template_path:
            results['template_found'] = True
            results['template_path'] = str(template_path)
        else:
            results['warnings'].append(
                f"Default template not found. Please place '{VALIDATION.get('TEMPLATE_FILENAME')}' "
                f"in {get_template_path()}"
            )
        
        # Check write permissions
        data_dir = get_package_data_dir()
        if not os.access(data_dir, os.W_OK):
            results['errors'].append(f"No write permission to data directory: {data_dir}")
            results['is_valid'] = False
            
    except Exception as e:
        results['errors'].append(f"Package validation failed: {e}")
        results['is_valid'] = False
    
    return results


def get_version_info() -> Dict[str, str]:
    """Get version information for the package and dependencies"""
    from . import __version__, __author__, __email__
    
    info = {
        'package_version': __version__,
        'author': __author__,
        'email': __email__,
        'python_version': sys.version,
        'platform': sys.platform,
    }
    
    # Try to get dependency versions
    try:
        import pandas
        info['pandas_version'] = pandas.__version__
    except ImportError:
        info['pandas_version'] = 'Not installed'
    
    try:
        import docx
        info['python_docx_version'] = docx.__version__
    except (ImportError, AttributeError):
        info['python_docx_version'] = 'Not available'
    
    try:
        import openpyxl
        info['openpyxl_version'] = openpyxl.__version__
    except ImportError:
        info['openpyxl_version'] = 'Not installed'
    
    return info


def format_timestamp(dt: Optional[datetime] = None) -> str:
    """
    Format timestamp for display
    
    Args:
        dt: Datetime object to format (uses current time if None)
        
    Returns:
        Formatted timestamp string
    """
    if dt is None:
        dt = datetime.now()
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def safe_filename(filename: str, max_length: int = 200) -> str:
    """
    Create a safe filename by removing/replacing problematic characters
    
    Args:
        filename: Original filename
        max_length: Maximum allowed length
        
    Returns:
        Sanitized filename
    """
    import re
    
    # Remove or replace problematic characters
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
    
    # Remove leading/trailing dots and spaces
    safe_name = safe_name.strip('. ')
    
    # Truncate if too long
    if len(safe_name) > max_length:
        name, ext = os.path.splitext(safe_name)
        max_name_len = max_length - len(ext)
        safe_name = name[:max_name_len] + ext
    
    return safe_name or 'unnamed_file' 