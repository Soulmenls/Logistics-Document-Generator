#!/usr/bin/env python3
"""
Security Utilities Module

Centralized security functions for input validation, path sanitization,
and secure file operations to prevent common vulnerabilities.
"""

import os
import re
import hashlib
import logging
from pathlib import Path
from typing import Optional, List, Any
import pandas as pd

# Configure security logger
security_logger = logging.getLogger('security')
security_logger.setLevel(logging.INFO)

class SecurityError(Exception):
    """Custom exception for security-related errors"""
    pass

class InputValidator:
    """Comprehensive input validation utilities"""
    
    @staticmethod
    def validate_shipment_number(shipment_num: Any) -> bool:
        """Validate shipment number with realistic criteria based on actual data"""
        if pd.isna(shipment_num) or shipment_num is None:
            return True  # Allow empty shipment numbers as they are optional
        
        # Convert to string and handle float format
        shipment_str = str(shipment_num).strip()
        
        # Remove .0 suffix if it's a float
        if shipment_str.endswith('.0'):
            shipment_str = shipment_str[:-2]
        
        # Must be 8-12 digits (allowing for various shipment number formats)
        if not shipment_str.isdigit() or not (8 <= len(shipment_str) <= 12):
            return False
            
        # Check for obviously invalid patterns
        # Reject numbers with all same digits (e.g., 1111111111)
        if len(set(shipment_str)) == 1:
            return False
            
        return True
    
    @staticmethod
    def validate_do_number(do_num: Any) -> bool:
        """Validate DO number with enhanced security checks"""
        if pd.isna(do_num) or do_num is None:
            return False  # DO numbers are required
            
        do_str = str(do_num).strip()
        
        # Remove .0 suffix if it's a float
        if do_str.endswith('.0'):
            do_str = do_str[:-2]
        
        # Must be at least 6 digits, max 15 (reasonable limit, based on actual data)
        if not re.match(r'^\d{6,15}$', do_str):
            return False
            
        # Additional validation: check for suspicious patterns
        # Reject numbers with all same digits
        if len(set(do_str)) == 1:
            return False
            
        return True
    
    @staticmethod
    def validate_text_field(text: Any, max_length: int = 1000, allow_empty: bool = True) -> bool:
        """Validate text fields with length and content checks"""
        if pd.isna(text) or text is None:
            return allow_empty
            
        text_str = str(text).strip()
        
        if not allow_empty and not text_str:
            return False
            
        # Check length
        if len(text_str) > max_length:
            return False
            
        # Check for potentially malicious content
        suspicious_patterns = [
            r'<script[^>]*>.*?</script>',  # Script tags
            r'javascript:',                # JavaScript URLs
            r'vbscript:',                 # VBScript URLs
            r'on\w+\s*=',                 # Event handlers
            r'\\x[0-9a-fA-F]{2}',         # Hex encoded characters
        ]
        
        for pattern in suspicious_patterns:
            if re.search(pattern, text_str, re.IGNORECASE):
                security_logger.warning(f"Suspicious content detected: {pattern}")
                return False
                
        return True
    
    @staticmethod
    def validate_numeric_field(value: Any, min_val: float = 0, max_val: float = 1e9) -> bool:
        """Validate numeric fields with range checks"""
        if pd.isna(value) or value is None:
            return False
            
        try:
            num_val = float(value)
            return min_val <= num_val <= max_val
        except (ValueError, TypeError):
            return False

class PathSanitizer:
    """Secure path operations to prevent directory traversal attacks"""
    
    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """Sanitize filename to prevent path traversal and invalid characters"""
        if not filename:
            raise SecurityError("Filename cannot be empty")
            
        # Remove path separators and dangerous characters
        sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
        
        # Remove leading/trailing dots and spaces
        sanitized = sanitized.strip('. ')
        
        # Prevent reserved names on Windows
        reserved_names = {
            'CON', 'PRN', 'AUX', 'NUL',
            'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
            'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
        }
        
        if sanitized.upper() in reserved_names:
            sanitized = f"_{sanitized}"
            
        # Ensure filename is not empty after sanitization
        if not sanitized:
            raise SecurityError("Filename becomes empty after sanitization")
            
        return sanitized
    
    @staticmethod
    def validate_file_path(file_path: str, base_directory: str) -> bool:
        """Validate file path to prevent directory traversal attacks"""
        try:
            # Normalize and resolve paths
            base_path = Path(base_directory).resolve()
            target_path = Path(file_path).resolve()
            
            # Check if target path is within base directory
            return str(target_path).startswith(str(base_path))
            
        except (OSError, ValueError) as e:
            security_logger.error(f"Path validation error: {e}")
            return False
    
    @staticmethod
    def safe_join_path(base_directory: str, *path_parts: str) -> str:
        """Safely join path components with validation"""
        # Sanitize each path part
        sanitized_parts = []
        for part in path_parts:
            if not part or part in ('.', '..'):
                raise SecurityError(f"Invalid path component: {part}")
            sanitized_parts.append(PathSanitizer.sanitize_filename(part))
        
        # Join paths
        full_path = os.path.join(base_directory, *sanitized_parts)
        
        # Validate the result
        if not PathSanitizer.validate_file_path(full_path, base_directory):
            raise SecurityError("Path traversal attempt detected")
            
        return full_path

class SecureFileHandler:
    """Secure file operations with validation and logging"""
    
    def __init__(self, base_directory: str):
        self.base_directory = os.path.abspath(base_directory)
        if not os.path.exists(self.base_directory):
            os.makedirs(self.base_directory, mode=0o755)
    
    def safe_file_exists(self, file_path: str) -> bool:
        """Safely check if file exists with path validation"""
        try:
            if not PathSanitizer.validate_file_path(file_path, self.base_directory):
                security_logger.warning(f"Path validation failed: {file_path}")
                return False
            return os.path.exists(file_path)
        except Exception as e:
            security_logger.error(f"File existence check failed: {e}")
            return False
    
    def safe_list_files(self, pattern: str) -> List[str]:
        """Safely list files with pattern matching and validation"""
        try:
            import glob
            # Construct safe pattern within base directory
            safe_pattern = os.path.join(self.base_directory, pattern)
            
            # Use glob with validation
            files = glob.glob(safe_pattern)
            
            # Additional validation for each file
            validated_files = []
            for file_path in files:
                if PathSanitizer.validate_file_path(file_path, self.base_directory):
                    validated_files.append(file_path)
                else:
                    security_logger.warning(f"Filtered out invalid path: {file_path}")
            
            return validated_files
            
        except Exception as e:
            security_logger.error(f"File listing failed: {e}")
            return []
    
    def calculate_file_hash(self, file_path: str) -> Optional[str]:
        """Calculate SHA-256 hash of file for integrity verification"""
        try:
            if not self.safe_file_exists(file_path):
                return None
                
            hash_sha256 = hashlib.sha256()
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_sha256.update(chunk)
            return hash_sha256.hexdigest()
            
        except Exception as e:
            security_logger.error(f"File hash calculation failed: {e}")
            return None

class SecurityConfig:
    """Security configuration and constants"""
    
    # File size limits (in bytes)
    MAX_EXCEL_FILE_SIZE = 100 * 1024 * 1024  # 100 MB
    MAX_TEMPLATE_FILE_SIZE = 10 * 1024 * 1024  # 10 MB
    
    # Processing limits
    MAX_RECORDS_PER_BATCH = 10000
    MAX_PROCESSING_TIME = 3600  # 1 hour in seconds
    
    # Allowed file extensions
    ALLOWED_EXCEL_EXTENSIONS = {'.xlsx', '.xls'}
    ALLOWED_TEMPLATE_EXTENSIONS = {'.docx'}
    
    # Rate limiting
    MAX_OPERATIONS_PER_MINUTE = 60
    
    @staticmethod
    def validate_file_size(file_path: str, max_size: int) -> bool:
        """Validate file size against limits"""
        try:
            file_size = os.path.getsize(file_path)
            return file_size <= max_size
        except OSError:
            return False
    
    @staticmethod
    def validate_file_extension(file_path: str, allowed_extensions: set) -> bool:
        """Validate file extension against allowed list"""
        file_ext = os.path.splitext(file_path)[1].lower()
        return file_ext in allowed_extensions

# Rate limiting utility
class RateLimiter:
    """Simple rate limiter for operations"""
    
    def __init__(self, max_operations: int, time_window: int = 60):
        self.max_operations = max_operations
        self.time_window = time_window
        self.operations = []
    
    def allow_operation(self) -> bool:
        """Check if operation is allowed based on rate limits"""
        import time
        current_time = time.time()
        
        # Remove old operations outside time window
        self.operations = [op_time for op_time in self.operations 
                          if current_time - op_time < self.time_window]
        
        # Check if under limit
        if len(self.operations) < self.max_operations:
            self.operations.append(current_time)
            return True
        
        return False 