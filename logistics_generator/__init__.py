#!/usr/bin/env python3
"""
Logistics Document Generator Package

A secure, high-performance logistics document generator for creating shipping 
placards from Excel data using Word templates.

Features:
- Secure input validation and sanitization
- Path traversal protection
- File size and type validation
- Rate limiting and resource management
- Comprehensive error handling and logging
- GUI and CLI interfaces
"""

from .core import PlacardGenerator
from .security import (
    InputValidator, 
    PathSanitizer, 
    SecureFileHandler, 
    SecurityConfig,
    SecurityError,
    RateLimiter
)
from .config import (
    APP_NAME,
    APP_VERSION,
    SECURITY,
    PERFORMANCE,
    LOGGING,
    DIRECTORIES,
    VALIDATION
)

__version__ = "2.1.0"
__author__ = "Logistics Team"
__email__ = "logistics@company.com"
__license__ = "MIT"

__all__ = [
    # Core functionality
    "PlacardGenerator",
    
    # Security utilities
    "InputValidator",
    "PathSanitizer", 
    "SecureFileHandler",
    "SecurityConfig",
    "SecurityError",
    "RateLimiter",
    
    # Configuration
    "APP_NAME",
    "APP_VERSION",
    "SECURITY",
    "PERFORMANCE", 
    "LOGGING",
    "DIRECTORIES",
    "VALIDATION",
    
    # Package metadata
    "__version__",
    "__author__",
    "__email__",
    "__license__",
]

# Initialize logging when package is imported
import logging
from .utils import setup_logging

# Setup default logging configuration
setup_logging()

# Package-level logger
logger = logging.getLogger(__name__)
logger.info(f"Initialized {APP_NAME} v{__version__}") 