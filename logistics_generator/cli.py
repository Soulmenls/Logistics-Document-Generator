#!/usr/bin/env python3
"""
Command Line Interface for Logistics Document Generator

This module provides the command-line interface for the package,
allowing users to run the placard generator from the terminal.
"""

import sys
import argparse
import logging
from pathlib import Path
from typing import List, Optional

from .core import PlacardGenerator
from .utils import (
    setup_logging, 
    validate_package_installation,
    get_version_info,
    ensure_package_directories
)
from . import __version__, __author__


def create_parser() -> argparse.ArgumentParser:
    """Create and configure the argument parser"""
    parser = argparse.ArgumentParser(
        prog='placard-generator',
        description='Generate shipping placards from Excel data using Word templates',
        epilog=f'Logistics Document Generator v{__version__} by {__author__}',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    # Version information
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__}'
    )
    
    # Logging options
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging (DEBUG level)'
    )
    
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress all output except errors'
    )
    
    # Processing options
    parser.add_argument(
        '-s', '--shipments',
        nargs='+',
        help='Specific shipment numbers to process (space-separated)'
    )
    
    parser.add_argument(
        '-a', '--all',
        action='store_true',
        help='Process all shipments in the dataset'
    )
    
    parser.add_argument(
        '-f', '--file',
        type=str,
        help='Specify Excel file path (overrides auto-detection)'
    )
    
    parser.add_argument(
        '-t', '--template',
        type=str,
        help='Specify template file path (overrides default)'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=str,
        help='Specify output directory (overrides default)'
    )
    
    # Information commands
    parser.add_argument(
        '--info',
        action='store_true',
        help='Show package information and exit'
    )
    
    parser.add_argument(
        '--validate',
        action='store_true',
        help='Validate package installation and exit'
    )
    
    parser.add_argument(
        '--setup',
        action='store_true',
        help='Setup directories and show configuration'
    )
    
    return parser


def setup_cli_logging(verbose: bool = False, quiet: bool = False) -> None:
    """Setup logging for CLI usage"""
    if quiet:
        log_level = 'ERROR'
    elif verbose:
        log_level = 'DEBUG'
    else:
        log_level = 'INFO'
    
    setup_logging(log_level=log_level, log_to_file=True)


def show_info() -> None:
    """Display package information"""
    info = get_version_info()
    
    print("Logistics Document Generator - Package Information")
    print("=" * 50)
    print(f"Version: {info['package_version']}")
    print(f"Author: {info['author']}")
    print(f"Email: {info['email']}")
    print(f"Python: {info['python_version']}")
    print(f"Platform: {info['platform']}")
    print()
    print("Dependencies:")
    print(f"  - pandas: {info['pandas_version']}")
    print(f"  - python-docx: {info['python_docx_version']}")
    print(f"  - openpyxl: {info['openpyxl_version']}")


def show_validation() -> bool:
    """Validate and display package installation status"""
    print("Validating package installation...")
    print("=" * 40)
    
    validation = validate_package_installation()
    
    # Show directories
    print("Directories:")
    for name, path in validation['directories'].items():
        print(f"  {name}: {path}")
    
    print()
    
    # Show template status
    if validation['template_found']:
        print(f"✓ Template found: {validation.get('template_path', 'Unknown')}")
    else:
        print("⚠ Template not found")
    
    print()
    
    # Show warnings
    if validation['warnings']:
        print("Warnings:")
        for warning in validation['warnings']:
            print(f"  ⚠ {warning}")
        print()
    
    # Show errors
    if validation['errors']:
        print("Errors:")
        for error in validation['errors']:
            print(f"  ✗ {error}")
        print()
    
    # Overall status
    if validation['is_valid']:
        print("✓ Package installation is valid")
        return True
    else:
        print("✗ Package installation has issues")
        return False


def setup_directories() -> None:
    """Setup package directories and show configuration"""
    print("Setting up directories...")
    
    try:
        directories = ensure_package_directories()
        
        print("Created/verified directories:")
        for name, path in directories.items():
            print(f"  {name}: {path}")
            
        print("\nNext steps:")
        print("1. Place your Excel data files in the Data directory")
        print("2. Place your template file (placard_template.docx) in the Template directory")
        print("3. Run the generator with your shipment numbers")
        
    except Exception as e:
        print(f"Error setting up directories: {e}")
        sys.exit(1)


def run_generator(args: argparse.Namespace) -> int:
    """Run the placard generator with the given arguments"""
    try:
        generator = PlacardGenerator()
        
        # Initialize directories and logging
        if not generator.setup_directories():
            print("ERROR: Could not setup required directories")
            return 1
        
        if not generator.initialize_log():
            print("WARNING: Could not initialize logging")
        
        # Load data
        print("Loading and preparing data...")
        if not generator.load_and_prepare_data():
            print("ERROR: Failed to load data")
            return 1
        
        # Process shipments
        if args.all:
            # Process all shipments
            successful, failed = generator.process_all_shipments()
            
            print(f"\nProcessing complete:")
            print(f"  Successful: {successful}")
            print(f"  Failed: {failed}")
            
            return 0 if failed == 0 else 1
            
        elif args.shipments:
            # Process specific shipments
            successful = 0
            failed = 0
            
            for shipment in args.shipments:
                if generator.process_shipment(shipment.strip()):
                    successful += 1
                else:
                    failed += 1
            
            print(f"\nProcessing complete:")
            print(f"  Successful: {successful}")
            print(f"  Failed: {failed}")
            
            return 0 if failed == 0 else 1
            
        else:
            # Interactive mode
            print("Starting interactive mode...")
            generator.run()
            return 0
            
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
        return 1
    except Exception as e:
        print(f"ERROR: {e}")
        logging.exception("Unexpected error in CLI")
        return 1


def main(argv: Optional[List[str]] = None) -> int:
    """
    Main CLI entry point
    
    Args:
        argv: Command line arguments (uses sys.argv if None)
        
    Returns:
        Exit code (0 for success, non-zero for failure)
    """
    if argv is None:
        argv = sys.argv[1:]
    
    parser = create_parser()
    args = parser.parse_args(argv)
    
    # Setup logging early
    setup_cli_logging(args.verbose, args.quiet)
    
    # Handle information commands
    if args.info:
        show_info()
        return 0
    
    if args.validate:
        is_valid = show_validation()
        return 0 if is_valid else 1
    
    if args.setup:
        setup_directories()
        return 0
    
    # Run the generator
    return run_generator(args)


if __name__ == '__main__':
    sys.exit(main()) 