#!/usr/bin/env python3
"""
Logistics Document Generator GUI

A professional GUI interface for the logistics document generator using Dear PyGui.
Features an interactive table for selecting shipments, real-time progress tracking,
and comprehensive error handling for enterprise-grade reliability.

Author: Logistics Team
Version: 2.0.0
"""

import os
import sys
import threading
import time
import traceback
import logging
import weakref
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Any, Set, cast
from collections import deque
import gc

import dearpygui.dearpygui as dpg
import pandas as pd

# Import the existing placard generator functionality
try:
    from placard_generator import PlacardGenerator
    from security_utils import (
        InputValidator, PathSanitizer, SecurityError, 
        RateLimiter, SecurityConfig, security_logger
    )
except ImportError as e:
    print(f"CRITICAL ERROR: Could not import required modules: {e}")
    print("Please ensure placard_generator.py and security_utils.py are in the same directory.")
    sys.exit(1)

# Configure GUI-specific logging
gui_logger = logging.getLogger('gui_app')
gui_logger.setLevel(logging.INFO)


class PlacardGeneratorGUI:
    """Professional GUI interface for the Logistics Document Generator with comprehensive error handling"""
    
    def __init__(self):
        """Initialize the GUI application with comprehensive security and memory management"""
        try:
            # Core application state
            self.generator = PlacardGenerator()
            self.shipment_data: List[Dict[str, Any]] = []
            self.filtered_data: List[Dict[str, Any]] = []
            self.selected_shipments: Set[str] = set()
            self.is_processing = False
            self.data_loaded = False
            self.load_error = ""
            self.search_text = ""
            
            # Thread safety and security
            self._processing_lock = threading.RLock()
            self._shutdown_event = threading.Event()
            self.rate_limiter = RateLimiter(SecurityConfig.MAX_OPERATIONS_PER_MINUTE)
            
            # Memory management
            self._memory_cleanup_interval = 30  # seconds
            self._last_cleanup = time.time()
            self._weak_references: Set[weakref.ref] = set()
            
            # Column filters with size limits
            self.column_filters = {
                'shipment_nbr': '',
                'do_numbers': '',
                'ship_to': '',
                'po': '',
                'vas': '',
                'label_type': '',
                'order_type': '',
                'pmt_term': ''
            }
            self.selected_total_units = 0
            self.dropdown_options: Dict[str, List[str]] = {}
            self.updating_master_checkbox = False
            
            # Multi-select filter selections with memory management
            self.multi_select_filters = {
                'shipment_nbr': set(),
                'do_numbers': set(),
                'ship_to': set(),
                'po': set(),
                'vas': set(),
                'label_type': set(),
                'order_type': set(),
                'pmt_term': set()
            }
            
            # Filter search text for each dropdown
            self.filter_search_text = {
                'shipment_nbr': '',
                'do_numbers': '',
                'ship_to': '',
                'po': '',
                'vas': '',
                'label_type': '',
                'order_type': '',
                'pmt_term': ''
            }
            
            # Sort state for each dropdown filter
            self.filter_sort_state = {
                'shipment_nbr': 'default',
                'do_numbers': 'default',
                'ship_to': 'default',
                'po': 'default',
                'vas': 'default',
                'label_type': 'default',
                'order_type': 'default',
                'pmt_term': 'default'
            }
            
            # GUI theme colors - solid professional theme
            self.colors = {
                'primary': [41, 74, 122],      # Dark blue
                'secondary': [70, 130, 180],   # Steel blue
                'accent': [100, 149, 237],     # Cornflower blue
                'success': [40, 167, 69],      # Green
                'warning': [255, 193, 7],      # Amber
                'danger': [220, 53, 69],       # Red
                'light': [248, 249, 250],      # Light gray
                'dark': [33, 37, 41],          # Dark gray
                'background': [40, 44, 52],    # Solid dark blue-gray background
                'surface': [50, 54, 62],       # Solid surface color
            }
            
            # Console log with memory management - use deque for O(1) operations
            self.console_logs = deque(maxlen=50)  # Limit to 50 messages to prevent memory leaks
            self.max_console_lines = 50  # Reduced from 100 for better memory management
            
            # Performance monitoring
            self._operation_times = deque(maxlen=100)
            self._memory_usage = deque(maxlen=20)
            
            # Initialize Dear PyGui with error handling
            self._initialize_gui()
            
        except Exception as e:
            self._handle_critical_error("GUI Initialization", e)
            sys.exit(1)
    
    def _initialize_gui(self):
        """Initialize Dear PyGui with comprehensive error handling"""
        try:
            dpg.create_context()
            self.setup_themes()
            self.setup_fonts()
            self.log_to_console("GUI initialization completed successfully", "success")
        except Exception as e:
            raise RuntimeError(f"Failed to initialize Dear PyGui: {e}")
    
    def _handle_critical_error(self, operation: str, error: Exception):
        """Handle critical errors that prevent application startup"""
        error_msg = f"CRITICAL ERROR in {operation}: {str(error)}"
        print(f"\n{'='*60}")
        print(error_msg)
        print(f"{'='*60}")
        print("Stack trace:")
        print(traceback.format_exc())
        print(f"{'='*60}")
        print("Application cannot continue. Please check the error above.")
        
    def _safe_dpg_operation(self, operation_name: str, operation_func, *args, **kwargs):
        """Safely execute Dear PyGui operations with error handling"""
        try:
            return operation_func(*args, **kwargs)
        except Exception as e:
            self.log_to_console(f"GUI operation '{operation_name}' failed: {str(e)}", "error")
            return None
    
    def _validate_data_integrity(self) -> bool:
        """Validate data integrity before processing"""
        try:
            if not self.shipment_data:
                self.update_status("No data loaded for validation", "warning")
                return False
            
            # Check for required fields in each shipment record
            required_fields = ['shipment_nbr', 'do_numbers', 'ship_to', 'original_qty']
            invalid_records = []
            
            for i, shipment in enumerate(self.shipment_data):
                missing_fields = [field for field in required_fields if not shipment.get(field)]
                if missing_fields:
                    invalid_records.append(f"Record {i+1}: missing {', '.join(missing_fields)}")
            
            if invalid_records:
                self.log_to_console(f"Data validation failed: {len(invalid_records)} invalid records", "error")
                for record in invalid_records[:5]:  # Show first 5 errors
                    self.log_to_console(f"  - {record}", "error")
                if len(invalid_records) > 5:
                    self.log_to_console(f"  ... and {len(invalid_records) - 5} more errors", "error")
                return False
            
            self.log_to_console(f"Data validation passed: {len(self.shipment_data)} records validated", "success")
            return True
            
        except Exception as e:
            self.log_to_console(f"Data validation error: {str(e)}", "error")
            return False
    
    def setup_themes(self):
        """Setup custom themes for the application"""
        # Main theme with solid color design
        with dpg.theme() as main_theme:
            with dpg.theme_component(dpg.mvAll):
                dpg.add_theme_color(dpg.mvThemeCol_WindowBg, [40, 44, 52])  # Solid dark blue-gray
                dpg.add_theme_color(dpg.mvThemeCol_ChildBg, [40, 44, 52])  # Same solid color
                dpg.add_theme_color(dpg.mvThemeCol_PopupBg, [45, 49, 57])  # Slightly lighter solid
                dpg.add_theme_color(dpg.mvThemeCol_FrameBg, [50, 54, 62])  # Input backgrounds solid
                dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, [60, 64, 72])  # Solid hover
                dpg.add_theme_color(dpg.mvThemeCol_FrameBgActive, [70, 74, 82])   # Solid active
                dpg.add_theme_color(dpg.mvThemeCol_TitleBg, [35, 39, 47])        # Solid title bg
                dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, [40, 44, 52])  # Solid active title
                dpg.add_theme_color(dpg.mvThemeCol_MenuBarBg, [40, 44, 52])      # Solid menu bar
                dpg.add_theme_color(dpg.mvThemeCol_Button, [52, 95, 150])      # Standard default blue
                dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [62, 105, 160])  # +10 lighter
                dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, [42, 85, 140])    # -10 darker
                dpg.add_theme_color(dpg.mvThemeCol_Header, [55, 59, 67])           # Solid header
                dpg.add_theme_color(dpg.mvThemeCol_HeaderHovered, [65, 69, 77])    # Solid hover
                dpg.add_theme_color(dpg.mvThemeCol_HeaderActive, [75, 79, 87])     # Solid active
                dpg.add_theme_color(dpg.mvThemeCol_Tab, [45, 49, 57])             # Solid tab
                dpg.add_theme_color(dpg.mvThemeCol_TabHovered, [55, 59, 67])      # Solid tab hover
                dpg.add_theme_color(dpg.mvThemeCol_TabActive, [65, 69, 77])       # Solid tab active
                dpg.add_theme_color(dpg.mvThemeCol_TableHeaderBg, [50, 54, 62])   # Solid table header
                dpg.add_theme_color(dpg.mvThemeCol_TableBorderStrong, [70, 74, 82]) # Solid border
                dpg.add_theme_color(dpg.mvThemeCol_TableBorderLight, [55, 59, 67])  # Solid light border
                dpg.add_theme_color(dpg.mvThemeCol_TableRowBg, [40, 44, 52])      # Solid row background
                dpg.add_theme_color(dpg.mvThemeCol_TableRowBgAlt, [45, 49, 57])   # Solid alternating row
                dpg.add_theme_color(dpg.mvThemeCol_Text, [220, 220, 220])  # Light text
                dpg.add_theme_color(dpg.mvThemeCol_CheckMark, [70, 130, 200])  # Blue checkmarks
                dpg.add_theme_color(dpg.mvThemeCol_Border, [60, 60, 60])
                dpg.add_theme_color(dpg.mvThemeCol_Separator, [50, 50, 50])
                # Sharp, modern styling with better button text centering
                dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 3)  # Slightly more rounding for buttons
                dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 0)  # Sharp corners
                dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 2)
                dpg.add_theme_style(dpg.mvStyleVar_PopupRounding, 2)
                dpg.add_theme_style(dpg.mvStyleVar_ScrollbarRounding, 0)
                dpg.add_theme_style(dpg.mvStyleVar_GrabRounding, 2)
                dpg.add_theme_style(dpg.mvStyleVar_TabRounding, 2)
                dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 12, 10)  # Better padding for text centering
                dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 20, 20)  # Generous padding
                dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, 10)  # Clean spacing
                dpg.add_theme_style(dpg.mvStyleVar_ItemInnerSpacing, 10, 6)  # Better inner spacing for buttons
                dpg.add_theme_style(dpg.mvStyleVar_IndentSpacing, 25)
                dpg.add_theme_style(dpg.mvStyleVar_CellPadding, 8, 8)  # Table cell padding
                dpg.add_theme_style(dpg.mvStyleVar_WindowBorderSize, 1)
                dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 1)
                dpg.add_theme_style(dpg.mvStyleVar_ButtonTextAlign, 0.5, 0.5)  # Center button text
        
        # Standardized button base styling
        button_base_styles = {
            'rounding': 3,
            'padding': (12, 8),
            'text_align': (0.5, 0.5),
            'text_color': [255, 255, 255],  # White text for all buttons
        }
        
        # Success button theme - Standardized green
        with dpg.theme() as success_theme:
            with dpg.theme_component(dpg.mvButton):
                dpg.add_theme_color(dpg.mvThemeCol_Button, [34, 139, 34])      # Standard green
                dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [44, 149, 44])  # +10 lighter
                dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, [24, 129, 24])   # -10 darker
                dpg.add_theme_color(dpg.mvThemeCol_Text, button_base_styles['text_color'])
                dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, button_base_styles['rounding'])
                dpg.add_theme_style(dpg.mvStyleVar_FramePadding, *button_base_styles['padding'])
                dpg.add_theme_style(dpg.mvStyleVar_ButtonTextAlign, *button_base_styles['text_align'])
        
        # Warning button theme - Standardized orange
        with dpg.theme() as warning_theme:
            with dpg.theme_component(dpg.mvButton):
                dpg.add_theme_color(dpg.mvThemeCol_Button, [255, 140, 0])      # Standard orange
                dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [255, 150, 10])  # +10 lighter
                dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, [245, 130, 0])    # -10 darker
                dpg.add_theme_color(dpg.mvThemeCol_Text, button_base_styles['text_color'])
                dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, button_base_styles['rounding'])
                dpg.add_theme_style(dpg.mvStyleVar_FramePadding, *button_base_styles['padding'])
                dpg.add_theme_style(dpg.mvStyleVar_ButtonTextAlign, *button_base_styles['text_align'])
        
        # Danger button theme - Standardized red
        with dpg.theme() as danger_theme:
            with dpg.theme_component(dpg.mvButton):
                dpg.add_theme_color(dpg.mvThemeCol_Button, [220, 53, 69])      # Standard red
                dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [230, 63, 79])  # +10 lighter
                dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, [210, 43, 59])   # -10 darker
                dpg.add_theme_color(dpg.mvThemeCol_Text, button_base_styles['text_color'])
                dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, button_base_styles['rounding'])
                dpg.add_theme_style(dpg.mvStyleVar_FramePadding, *button_base_styles['padding'])
                dpg.add_theme_style(dpg.mvStyleVar_ButtonTextAlign, *button_base_styles['text_align'])
        
        # Primary action button theme - Standardized blue
        with dpg.theme() as primary_theme:
            with dpg.theme_component(dpg.mvButton):
                dpg.add_theme_color(dpg.mvThemeCol_Button, [0, 123, 255])      # Standard blue
                dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [10, 133, 255])  # +10 lighter
                dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, [0, 113, 245])    # -10 darker
                dpg.add_theme_color(dpg.mvThemeCol_Text, button_base_styles['text_color'])
                dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, button_base_styles['rounding'])
                dpg.add_theme_style(dpg.mvStyleVar_FramePadding, *button_base_styles['padding'])
                dpg.add_theme_style(dpg.mvStyleVar_ButtonTextAlign, *button_base_styles['text_align'])
        
        self.main_theme = main_theme
        self.success_theme = success_theme
        self.warning_theme = warning_theme
        self.danger_theme = danger_theme
        self.primary_theme = primary_theme
        
    def setup_fonts(self):
        """Setup custom fonts with comprehensive error handling"""
        try:
            # Create a font registry
            with dpg.font_registry():
                # Try to load system fonts with fallbacks
                font_paths = [
                    "C:/Windows/Fonts/segoeui.ttf",     # Windows
                    "/System/Library/Fonts/Arial.ttf",  # macOS
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"  # Linux
                ]
                
                default_font = None
                large_font = None
                bold_font = None
                
                # Try to load default font
                for font_path in font_paths:
                    if os.path.exists(font_path):
                        try:
                            default_font = dpg.add_font(font_path, 16)
                            large_font = dpg.add_font(font_path, 20)
                            self.log_to_console(f"Loaded fonts from: {font_path}", "success")
                            break
                        except Exception as e:
                            self.log_to_console(f"Failed to load font {font_path}: {e}", "warning")
                            continue
                
                # Try to load bold font
                bold_paths = [
                    "C:/Windows/Fonts/segoeuib.ttf",
                    "/System/Library/Fonts/Arial Bold.ttf",
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
                ]
                
                for bold_path in bold_paths:
                    if os.path.exists(bold_path):
                        try:
                            bold_font = dpg.add_font(bold_path, 16)
                            break
                        except Exception as e:
                            continue
                
                # Use default font as fallback for bold if needed
                if bold_font is None and default_font is not None:
                    bold_font = default_font
                    self.log_to_console("Using default font as bold fallback", "info")
            
            # Set font attributes with safe fallbacks
            self.default_font = default_font
            self.large_font = large_font if large_font else default_font
            self.bold_font = bold_font if bold_font else default_font
            
            if default_font is None:
                self.log_to_console("Warning: No custom fonts loaded, using system defaults", "warning")
                
        except Exception as e:
            self.log_to_console(f"Font setup error: {str(e)}", "error")
            # Set None values to prevent binding errors
            self.default_font = None
            self.large_font = None
            self.bold_font = None
    
    def _safe_bind_font(self, item, font):
        """Safely bind font to item with None check"""
        if font is not None:
            try:
                dpg.bind_item_font(item, font)
            except Exception as e:
                self.log_to_console(f"Font binding failed for item: {e}", "warning")
        
    def update_status(self, message: str, status_type: str = "info"):
        """Update the status bar with a message"""
        if dpg.does_item_exist("status_text"):
            dpg.set_value("status_text", f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
            
            # Change color based on status type
            if status_type == "success":
                color = self.colors['success']
            elif status_type == "warning":
                color = self.colors['warning']
            elif status_type == "error":
                color = self.colors['danger']
            else:
                color = [255, 255, 255]  # White for info
                
            dpg.configure_item("status_text", color=color)
        
        # Also log to console
        self.log_to_console(message, status_type)
    
    def log_to_console(self, message: str, log_type: str = "info"):
        """Add a message to the console log with security validation and memory management"""
        try:
            # Rate limiting for console logs
            if not self.rate_limiter.allow_operation():
                return
            
            # Validate and sanitize input
            if not InputValidator.validate_text_field(message, max_length=1000):
                message = "INVALID_LOG_MESSAGE"
                log_type = "error"
                security_logger.warning("Invalid log message filtered")
            
            timestamp = datetime.now().strftime('%H:%M:%S')
            log_entry = f"[{timestamp}] {message}"
            
            # Add to console logs - deque automatically handles size limit
            self.console_logs.append({
                'message': log_entry,
                'type': log_type,
                'timestamp': timestamp
            })
            
            # Update console display if it exists
            self.update_console_display()
            
            # Log to appropriate logger based on type
            if log_type == "error":
                gui_logger.error(message)
            elif log_type == "warning":
                gui_logger.warning(message)
            elif log_type == "success":
                gui_logger.info(f"SUCCESS: {message}")
            else:
                gui_logger.info(message)
            
            # Periodic memory cleanup
            current_time = time.time()
            if current_time - self._last_cleanup > self._memory_cleanup_interval:
                self._cleanup_memory()
                self._last_cleanup = current_time
                
        except Exception as e:
            # Fallback logging to prevent crash
            print(f"Console logging error: {e}")
            gui_logger.error(f"Console logging failed: {e}")
    
    def _cleanup_memory(self):
        """Perform periodic memory cleanup"""
        try:
            # Force garbage collection
            gc.collect()
            
            # Clean up weak references
            dead_refs = [ref for ref in self._weak_references if ref() is None]
            for ref in dead_refs:
                self._weak_references.discard(ref)
            
            # Limit data structures if they get too large
            if len(self.shipment_data) > SecurityConfig.MAX_RECORDS_PER_BATCH:
                gui_logger.warning("Large dataset detected, consider data cleanup")
            
            # Log memory usage if psutil is available
            try:
                import psutil
                process = psutil.Process()
                memory_mb = process.memory_info().rss / 1024 / 1024
                self._memory_usage.append(memory_mb)
                
                if len(self._memory_usage) > 10:
                    avg_memory = sum(self._memory_usage) / len(self._memory_usage)
                    if memory_mb > avg_memory * 1.5:  # 50% increase
                        gui_logger.warning(f"Memory usage spike detected: {memory_mb:.1f}MB")
            except ImportError:
                # psutil not available, skip memory monitoring
                pass
            
        except Exception as e:
            gui_logger.error(f"Memory cleanup failed: {e}")
    
    def update_console_display(self):
        """Update the console text display"""
        if dpg.does_item_exist("console_text"):
            # Create console text from logs
            console_text = "\n".join([log['message'] for log in self.console_logs])
            dpg.set_value("console_text", console_text)
            
            # Auto-scroll to bottom
            if dpg.does_item_exist("console_window"):
                dpg.set_y_scroll("console_window", -1)  # Scroll to bottom
    
    def clear_console_callback(self, sender, app_data):
        """Clear the console log"""
        self.console_logs.clear()
        self.update_console_display()
        self.log_to_console("Console cleared", "info")
    
    def load_data_callback(self, sender, app_data):
        """Callback to load Excel data with comprehensive security and error handling"""
        # Thread safety check
        with self._processing_lock:
            if self.is_processing:
                self.update_status("Already processing, please wait...", "warning")
                return
            self.is_processing = True
        
        try:
            self.update_status("Loading data...", "info")
            
            # Rate limiting check
            if not self.rate_limiter.allow_operation():
                self.update_status("Rate limit exceeded. Please wait before retrying.", "warning")
                return
            
            # Disable load button during loading
            self._safe_dpg_operation("disable_load_button", 
                                     dpg.configure_item, "load_data_btn", enabled=False, label="Loading...")
            
            # Clear previous data securely
            self.shipment_data.clear()
            self.filtered_data.clear()
            self.selected_shipments.clear()
            self.data_loaded = False
            
            # Load data with comprehensive error handling
            self.log_to_console("Starting data loading process...", "info")
            
            # Check if Data directory exists
            data_dir = os.path.join(os.getcwd(), "Data")
            if not os.path.exists(data_dir):
                raise FileNotFoundError(f"Data directory not found: {data_dir}")
            
            # Check for Excel files in Data directory
            excel_files = [f for f in os.listdir(data_dir) if f.endswith(('.xlsx', '.xls'))]
            if not excel_files:
                raise FileNotFoundError("No Excel files found in Data directory")
            
            self.log_to_console(f"Found {len(excel_files)} Excel file(s) in Data directory", "info")
            
            if self.generator.load_and_prepare_data():
                self.log_to_console("Excel data loaded successfully", "success")
                # Get all shipments with their details
                all_shipments = self.generator.get_all_unique_shipments()
                self.log_to_console(f"Found {len(all_shipments)} unique shipments", "info")
                
                self.shipment_data = []
                for shipment in all_shipments:
                    if self.generator.df is not None:
                        # Use the same conversion logic as get_all_unique_shipments()
                        df_shipment_clean = self.generator.df['Shipment Nbr'].astype(float).astype(int).astype(str)
                        shipment_df = self.generator.df[df_shipment_clean == str(shipment)]
                        
                        if not shipment_df.empty:
                            first_record = shipment_df.iloc[0]
                            do_count = len(shipment_df)
                            
                            # Get all DOs for this shipment
                            do_series = cast(pd.Series, shipment_df['DO #'])
                            do_numbers = do_series.unique().tolist()
                            do_list = ', '.join([str(int(do)) for do in do_numbers if pd.notna(do)])
                            
                            # Get total original quantity
                            qty_series = cast(pd.Series, shipment_df['Original Qty'])
                            total_qty = qty_series.sum()
                            
                            # Get all POs for this shipment
                            po_series = cast(pd.Series, shipment_df['PO'])
                            po_list = po_series.dropna().unique().tolist()
                            po_display = ', '.join([str(po) for po in po_list[:3]])  # Show first 3 POs
                            if len(po_list) > 3:
                                po_display += f" (+{len(po_list) - 3} more)"
                            
                            self.shipment_data.append({
                                'selected': False,
                                'shipment_nbr': shipment,
                                'do_numbers': do_list,
                                'do_count': do_count,
                                'ship_to': str(first_record.get('Ship To', 'N/A')),
                                'po': po_display if po_display else 'N/A',
                                'vas': str(first_record.get('VAS', 'N/A')),
                                'original_qty': str(int(total_qty)) if pd.notna(total_qty) else '0',
                                'label_type': str(first_record.get('Label Type', 'N/A')),
                                'order_type': str(first_record.get('Order Type', 'N/A')),
                                'pmt_term': str(first_record.get('Pmt Term', 'N/A')),
                                'start_ship': str(first_record.get('Start Ship', 'N/A'))
                            })
                
                # Initialize filtered data
                self.filtered_data = self.shipment_data.copy()
                
                # Validate data integrity now that shipment_data is populated
                if not self._validate_data_integrity():
                    self.update_status("Data validation failed - check console for details", "error")
                    return
                
                # Mark data as loaded
                self.data_loaded = True
                
                # Populate dropdown filter options - CRITICAL for filtering to work!
                self.populate_dropdown_options()
                self.log_to_console("Dropdown filter options populated", "info")
                
                # Update GUI immediately
                self.update_status(f"Loaded {len(self.shipment_data)} shipments", "success")
                self.refresh_table()
                self.update_selection_count()
                self.log_to_console("GUI table refreshed and updated", "info")
                
                # Enable processing buttons
                dpg.configure_item("generate_selected_btn", enabled=True)
                dpg.configure_item("generate_all_btn", enabled=True)
                dpg.configure_item("select_all_btn", enabled=True)
                dpg.configure_item("deselect_all_btn", enabled=True)
                self.log_to_console("Processing buttons enabled", "info")
                
            else:
                self.update_status("Failed to load data. Check Data folder and file format.", "error")
                self.log_to_console("Data loading failed - check Data folder", "error")
                
        except SecurityError as e:
            self.update_status(f"Security error loading data: {str(e)}", "error")
            self.log_to_console(f"Security violation during data loading: {str(e)}", "error")
            security_logger.error(f"Data loading security error: {e}")
            
        except Exception as e:
            self.update_status(f"Error loading data: {str(e)}", "error")
            self.log_to_console(f"Exception during data loading: {str(e)}", "error")
            gui_logger.error(f"Data loading failed: {e}", exc_info=True)
            
        finally:
            # Re-enable load button and reset processing state
            dpg.configure_item("load_data_btn", enabled=True, label="Load Data")
            with self._processing_lock:
                self.is_processing = False
    
    def search_callback(self, sender, app_data):
        """Callback for search functionality"""
        self.search_text = app_data.lower().strip()
        self.apply_all_filters()
    
    def column_filter_callback(self, sender, app_data, user_data):
        """Callback for column-specific filters"""
        column_name = user_data
        filter_text = app_data.lower().strip()
        self.column_filters[column_name] = filter_text
        
        # Apply all filters
        self.apply_all_filters()
    
    def apply_all_filters(self):
        """Apply both search and multi-select column filters"""
        initial_count = len(self.shipment_data)
        filtered_data = self.shipment_data.copy()
        self.log_to_console(f"Applying filters to {initial_count} shipments", "info")
        
        # Apply main search filter
        if self.search_text:
            temp_filtered = []
            for shipment in filtered_data:
                search_fields = [
                    shipment['shipment_nbr'],
                    shipment['do_numbers'],
                    shipment['ship_to'],
                    shipment['po']
                ]
                if any(self.search_text in str(field).lower() for field in search_fields):
                    temp_filtered.append(shipment)
            filtered_data = temp_filtered
        
        # Apply multi-select column filters
        for column, selected_values in self.multi_select_filters.items():
            if selected_values:  # If any values are selected for this column
                print(f"DEBUG: Processing filter for {column} with selected values: {list(selected_values)}")
                temp_filtered = []
                for shipment in filtered_data:
                    shipment_value = str(shipment.get(column, '')).strip()
                    
                    # Special handling for comma-separated values (DO numbers, POs)
                    if column in ['do_numbers', 'po'] and ',' in shipment_value:
                        # Split comma-separated values and check if any match
                        shipment_values = [v.strip() for v in shipment_value.split(',')]
                        print(f"DEBUG: Checking {column} values {shipment_values} against selected {list(selected_values)}")
                        if any(selected_val in selected_values for selected_val in shipment_values):
                            temp_filtered.append(shipment)
                            print(f"DEBUG: Match found! Keeping shipment {shipment['shipment_nbr']}")
                    else:
                        # Direct match for single values
                        print(f"DEBUG: Checking {column} value '{shipment_value}' against selected {list(selected_values)}")
                        if shipment_value in selected_values:
                            temp_filtered.append(shipment)
                            print(f"DEBUG: Match found! Keeping shipment {shipment['shipment_nbr']}")
                
                filtered_data = temp_filtered
                print(f"DEBUG: After {column} filter: {len(filtered_data)} shipments remaining")
        
        self.filtered_data = filtered_data
        
        # Log filter results
        self.log_to_console(f"Filter results: {len(self.filtered_data)} of {len(self.shipment_data)} shipments", "info")
        active_filters = sum(1 for filters in self.multi_select_filters.values() if filters)
        if active_filters > 0:
            self.log_to_console(f"Active filters: {active_filters} column filters", "info")
        if self.search_text:
            self.log_to_console(f"Search filter: '{self.search_text}'", "info")
        
        self.refresh_table()
        self.update_selection_count()
        self.update_master_checkbox()
        
        # Update status
        if self.search_text or active_filters > 0:
            self.update_status(f"Showing {len(self.filtered_data)} of {len(self.shipment_data)} shipments (filters active)", "info")
        else:
            self.update_status(f"Showing all {len(self.shipment_data)} shipments", "info")
    
    def clear_all_filters_callback(self, sender, app_data):
        """Clear all multi-select column filters"""
        # Clear all multi-select filter values
        for key in self.multi_select_filters:
            self.multi_select_filters[key].clear()
        
        # Reset search text as well
        self.search_text = ''
        
        # Reset search input
        if dpg.does_item_exist("main_search_input"):
            dpg.set_value("main_search_input", "")
        
        # Update all filter button displays
        for column in self.multi_select_filters.keys():
            self.update_filter_display(column)
        
        # Reapply filters (which will show all data)
        self.apply_all_filters()
        
        # Update status
        self.update_status("All filters cleared", "success")
    
    def populate_dropdown_options(self):
        """Populate dropdown options from loaded data"""
        if not self.shipment_data:
            return
        
        print(f"DEBUG populate_dropdown_options: Processing {len(self.shipment_data)} shipments")
        
        # Initialize dropdown options as dict of sets temporarily
        temp_options = {
            'shipment_nbr': set(),
            'do_numbers': set(),
            'ship_to': set(),
            'po': set(),
            'vas': set(),
            'label_type': set(),
            'order_type': set(),
            'pmt_term': set()
        }
        
        # Collect unique values from all data
        for i, shipment in enumerate(self.shipment_data):
            if i < 3:  # Debug first few shipments
                print(f"DEBUG shipment {i}: {shipment}")
            for key in temp_options.keys():
                if key in shipment:
                    value = str(shipment[key]).strip()
                    if value and value != 'N/A':
                        # For DO numbers and PO, split comma-separated values
                        if key in ['do_numbers', 'po']:
                            parts = [part.strip() for part in value.split(',')]
                            for part in parts:
                                if part and not part.endswith('...'):
                                    temp_options[key].add(part)
                        else:
                            temp_options[key].add(value)
        
        # Convert sets to sorted lists and add "All" option
        self.dropdown_options = {}
        for key in temp_options:
            sorted_options = sorted(list(temp_options[key]))
            self.dropdown_options[key] = ["All"] + sorted_options
            print(f"DEBUG dropdown options for {key}: {len(sorted_options)} options - {sorted_options[:5] if sorted_options else []}")
        
        print("DEBUG populate_dropdown_options completed")
    
    def show_multi_select_popup(self, sender, app_data, user_data):
        """Show multi-select filter popup"""
        column_name = user_data
        popup_tag = f"filter_popup_{column_name}"
        
        # Delete existing popup if it exists
        if dpg.does_item_exist(popup_tag):
            dpg.delete_item(popup_tag)
        
        # Create popup window
        with dpg.window(
            label=f"Filter: {column_name.replace('_', ' ').title()}",
            tag=popup_tag,
            modal=True,
            width=400,
            height=500,
            pos=[200, 200]
        ):
            # Search box for filtering options
            dpg.add_text("Search options:")
            search_tag = f"search_{column_name}"
            dpg.add_input_text(
                tag=search_tag,
                hint="Type to filter options...",
                callback=self.filter_search_callback,
                user_data=column_name,
                width=-1
            )
            
            dpg.add_separator()
            
            # Sort options
            with dpg.group(horizontal=True):
                dpg.add_text("Sort:")
                dpg.add_button(
                    label="A-Z",
                    callback=self.sort_filter_options,
                    user_data=(column_name, "asc"),
                    width=50,
                    height=25
                )
                dpg.add_button(
                    label="Z-A",
                    callback=self.sort_filter_options,
                    user_data=(column_name, "desc"),
                    width=50,
                    height=25
                )
                # Numeric sort for shipment numbers and DO numbers
                if column_name in ['shipment_nbr', 'do_numbers']:
                    dpg.add_button(
                        label="1-9",
                        callback=self.sort_filter_options,
                        user_data=(column_name, "numeric_asc"),
                        width=50,
                        height=25
                    )
                    dpg.add_button(
                        label="9-1",
                        callback=self.sort_filter_options,
                        user_data=(column_name, "numeric_desc"),
                        width=50,
                        height=25
                    )
            
            dpg.add_separator()
            
            # Action buttons
            with dpg.group(horizontal=True):
                dpg.add_button(
                    label="Select All",
                    callback=self.select_all_filter_options,
                    user_data=column_name,
                    width=90,
                    height=30
                )
                dpg.bind_item_theme(dpg.last_item(), self.success_theme)
                dpg.add_button(
                    label="Clear All",
                    callback=self.clear_all_filter_options,
                    user_data=column_name,
                    width=90,
                    height=30
                )
                dpg.bind_item_theme(dpg.last_item(), self.warning_theme)
                dpg.add_button(
                    label="Apply",
                    callback=self.apply_multi_select_filter,
                    user_data=column_name,
                    width=90,
                    height=30
                )
                dpg.bind_item_theme(dpg.last_item(), self.success_theme)
                dpg.add_button(
                    label="Cancel",
                    callback=lambda: dpg.delete_item(popup_tag),
                    width=90,
                    height=30
                )
                dpg.bind_item_theme(dpg.last_item(), self.danger_theme)
            
            dpg.add_separator()
            
            # Selected count
            count_tag = f"selected_count_{column_name}"
            selected_count = len(self.multi_select_filters[column_name])
            dpg.add_text(f"Selected: {selected_count}", tag=count_tag, color=[100, 149, 237])
            
            dpg.add_separator()
            
            # Scrollable list of options with checkboxes
            with dpg.child_window(height=300):
                options_group_tag = f"options_group_{column_name}"
                with dpg.group(tag=options_group_tag):
                    self.populate_filter_options(column_name, options_group_tag)
    
    def filter_search_callback(self, sender, app_data, user_data):
        """Filter the options in the multi-select popup based on search"""
        column_name = user_data
        search_text = app_data.lower()
        self.filter_search_text[column_name] = search_text
        
        # Refresh the options list
        options_group_tag = f"options_group_{column_name}"
        self.populate_filter_options(column_name, options_group_tag)
    
    def sort_filter_options(self, sender, app_data, user_data):
        """Sort the filter options based on the selected sort method"""
        column_name, sort_type = user_data
        self.filter_sort_state[column_name] = sort_type
        
        # Refresh the options list with new sort order
        options_group_tag = f"options_group_{column_name}"
        self.populate_filter_options(column_name, options_group_tag)
    
    def populate_filter_options(self, column_name, group_tag):
        """Populate the filter options with checkboxes"""
        # Clear existing options
        children = dpg.get_item_children(group_tag, slot=1)
        if children:
            for child in children:
                dpg.delete_item(child)
        
        if column_name not in self.dropdown_options:
            return
        
        search_text = self.filter_search_text.get(column_name, '').lower()
        options = self.dropdown_options[column_name]
        
        # Filter options based on search
        if search_text:
            filtered_options = [opt for opt in options if opt != "All" and search_text in opt.lower()]
        else:
            filtered_options = [opt for opt in options if opt != "All"]
        
        # Apply sorting based on current sort state
        sort_state = self.filter_sort_state.get(column_name, 'default')
        
        if sort_state == 'asc':
            filtered_options.sort()
        elif sort_state == 'desc':
            filtered_options.sort(reverse=True)
        elif sort_state == 'numeric_asc':
            # Numeric sort for shipment numbers and DO numbers
            try:
                filtered_options.sort(key=lambda x: int(x.split(',')[0].strip()) if x.replace(',', '').replace(' ', '').isdigit() else float('inf'))
            except (ValueError, AttributeError):
                filtered_options.sort()  # Fallback to alphabetical
        elif sort_state == 'numeric_desc':
            # Numeric sort descending
            try:
                filtered_options.sort(key=lambda x: int(x.split(',')[0].strip()) if x.replace(',', '').replace(' ', '').isdigit() else float('inf'), reverse=True)
            except (ValueError, AttributeError):
                filtered_options.sort(reverse=True)  # Fallback to alphabetical descending
        # Default: keep original order from dropdown_options
        
        # Add checkboxes for each option
        for option in filtered_options:
            is_selected = option in self.multi_select_filters[column_name]
            checkbox_tag = f"checkbox_{column_name}_{hash(option)}"
            
            with dpg.group(horizontal=True, parent=group_tag):
                dpg.add_checkbox(
                    label=option,
                    tag=checkbox_tag,
                    default_value=is_selected,
                    callback=self.toggle_filter_option,
                    user_data=(column_name, option)
                )
    
    def toggle_filter_option(self, sender, app_data, user_data):
        """Toggle selection of a filter option"""
        column_name, option = user_data
        
        print(f"DEBUG toggle_filter_option: {column_name} = {option}, checked = {app_data}")
        
        if app_data:  # Checked
            self.multi_select_filters[column_name].add(option)
            print(f"DEBUG: Added {option} to {column_name} filter. Now has: {list(self.multi_select_filters[column_name])}")
        else:  # Unchecked
            self.multi_select_filters[column_name].discard(option)
            print(f"DEBUG: Removed {option} from {column_name} filter. Now has: {list(self.multi_select_filters[column_name])}")
        
        # Update selected count
        count_tag = f"selected_count_{column_name}"
        if dpg.does_item_exist(count_tag):
            selected_count = len(self.multi_select_filters[column_name])
            dpg.set_value(count_tag, f"Selected: {selected_count}")
    
    def select_all_filter_options(self, sender, app_data, user_data):
        """Select all visible filter options"""
        column_name = user_data
        search_text = self.filter_search_text.get(column_name, '').lower()
        options = self.dropdown_options.get(column_name, [])
        
        # Filter options based on search
        if search_text:
            filtered_options = [opt for opt in options if opt != "All" and search_text in opt.lower()]
        else:
            filtered_options = [opt for opt in options if opt != "All"]
        
        # Add all filtered options to selection
        self.multi_select_filters[column_name].update(filtered_options)
        
        # Update checkboxes and count
        options_group_tag = f"options_group_{column_name}"
        self.populate_filter_options(column_name, options_group_tag)
        
        count_tag = f"selected_count_{column_name}"
        if dpg.does_item_exist(count_tag):
            selected_count = len(self.multi_select_filters[column_name])
            dpg.set_value(count_tag, f"Selected: {selected_count}")
    
    def clear_all_filter_options(self, sender, app_data, user_data):
        """Clear all filter options"""
        column_name = user_data
        self.multi_select_filters[column_name].clear()
        
        # Update checkboxes and count
        options_group_tag = f"options_group_{column_name}"
        self.populate_filter_options(column_name, options_group_tag)
        
        count_tag = f"selected_count_{column_name}"
        if dpg.does_item_exist(count_tag):
            dpg.set_value(count_tag, "Selected: 0")
    
    def apply_multi_select_filter(self, sender, app_data, user_data):
        """Apply the multi-select filter and close popup"""
        column_name = user_data
        popup_tag = f"filter_popup_{column_name}"
        
        # Apply all filters (this will also update button displays via refresh_table)
        self.apply_all_filters()
        
        # Close popup
        dpg.delete_item(popup_tag)
    
    def update_filter_display(self, column_name):
        """Update the display text for a multi-select filter"""
        button_tag = f"filter_button_{column_name}"
        selected_count = len(self.multi_select_filters[column_name])
        
        if selected_count == 0:
            display_text = "All"
        elif selected_count == 1:
            option = list(self.multi_select_filters[column_name])[0]
            display_text = f"{option[:12]}..." if len(option) > 12 else option
        else:
            display_text = f"({selected_count}) selected"
        
        if dpg.does_item_exist(button_tag):
            dpg.configure_item(button_tag, label=display_text)
    
    def update_filter_dropdowns(self):
        """Update dropdown filter contents with actual data"""
        dropdown_tags = {
            'shipment_nbr': 'filter_shipment_combo',
            'do_numbers': 'filter_do_combo',
            'ship_to': 'filter_shipto_combo',
            'po': 'filter_po_combo',
            'vas': 'filter_vas_combo',
            'label_type': 'filter_label_combo',
            'order_type': 'filter_order_combo',
            'pmt_term': 'filter_pmt_combo'
        }
        
        for key, tag in dropdown_tags.items():
            if dpg.does_item_exist(tag) and key in self.dropdown_options:
                # Configure the combo with new items
                dpg.configure_item(tag, items=self.dropdown_options[key])
        
    def load_data_thread(self):
        """Load data in a separate thread"""
        try:
            print("DEBUG: Starting data load...")  # Debug print
            if self.generator.load_and_prepare_data():
                # Get all shipments with their details
                all_shipments = self.generator.get_all_unique_shipments()
                print(f"DEBUG: Found {len(all_shipments)} shipments")  # Debug print
                
                temp_shipment_data = []
                for shipment in all_shipments:
                    if self.generator.df is not None:
                        shipment_df = self.generator.df[
                            self.generator.df['Shipment Nbr'].astype(str) == shipment
                        ]
                        if not shipment_df.empty:
                            first_record = shipment_df.iloc[0]
                            do_count = len(shipment_df)
                            
                            # Get all DOs for this shipment
                            do_series = cast(pd.Series, shipment_df['DO #'])
                            do_numbers = do_series.unique().tolist()
                            do_list = ', '.join([str(int(do)) for do in do_numbers if pd.notna(do)])
                            
                            # Get total original quantity
                            qty_series = cast(pd.Series, shipment_df['Original Qty'])
                            total_qty = qty_series.sum()
                            
                            # Get all POs for this shipment
                            po_series = cast(pd.Series, shipment_df['PO'])
                            po_list = po_series.dropna().unique().tolist()
                            po_display = ', '.join([str(po) for po in po_list[:3]])  # Show first 3 POs
                            if len(po_list) > 3:
                                po_display += f" (+{len(po_list) - 3} more)"
                            
                            temp_shipment_data.append({
                                'selected': False,
                                'shipment_nbr': shipment,
                                'do_numbers': do_list,
                                'do_count': do_count,
                                'ship_to': str(first_record.get('Ship To', 'N/A')),
                                'po': po_display if po_display else 'N/A',
                                'vas': str(first_record.get('VAS', 'N/A')),
                                'original_qty': str(int(total_qty)) if pd.notna(total_qty) else '0',
                                'label_type': str(first_record.get('Label Type', 'N/A')),
                                'order_type': str(first_record.get('Order Type', 'N/A')),
                                'pmt_term': str(first_record.get('Pmt Term', 'N/A')),
                                'start_ship': str(first_record.get('Start Ship', 'N/A'))
                            })
                
                print(f"DEBUG: Prepared {len(temp_shipment_data)} shipment records")  # Debug print
                
                # Store data and schedule GUI update on main thread
                self.shipment_data = temp_shipment_data
                self.data_loaded = True
                print("DEBUG: Data loading completed successfully")  # Debug print
                
            else:
                print("DEBUG: Data loading failed")  # Debug print
                self.data_loaded = False
                self.load_error = "Failed to load data. Check Data folder and file format."
                
        except Exception as e:
            print(f"DEBUG: Exception during data loading: {e}")  # Debug print
            self.data_loaded = False
            self.load_error = f"Error loading data: {str(e)}"
    
    def check_loading_completion(self):
        """Check if data loading is complete and update GUI accordingly"""
        def check():
            print(f"DEBUG: Checking loading completion. data_loaded = {getattr(self, 'data_loaded', 'not set')}")  # Debug print
            
            if hasattr(self, 'data_loaded'):
                if self.data_loaded:
                    print(f"DEBUG: Data loaded successfully, updating GUI with {len(self.shipment_data)} shipments")  # Debug print
                    
                    # Data loaded successfully
                    dpg.set_value("data_status", f"✓ Loaded {len(self.shipment_data)} shipments")
                    self.update_status(f"Successfully loaded {len(self.shipment_data)} shipments", "success")
                    self.refresh_table()
                    self.update_selection_count()
                    
                    # Enable processing buttons
                    dpg.configure_item("generate_selected_btn", enabled=True)
                    dpg.configure_item("generate_all_btn", enabled=True)
                    dpg.configure_item("select_all_btn", enabled=True)
                    dpg.configure_item("deselect_all_btn", enabled=True)
                    
                    # Re-enable load button
                    dpg.configure_item("load_data_btn", enabled=True, label="Load Data")
                    
                elif hasattr(self, 'load_error') and self.load_error:
                    print(f"DEBUG: Data loading failed: {self.load_error}")  # Debug print
                    
                    # Data loading failed
                    dpg.set_value("data_status", "Failed to load data")
                    self.update_status(self.load_error, "error")
                    dpg.configure_item("load_data_btn", enabled=True, label="Load Data")
                    
                else:
                    print("DEBUG: Still loading, checking again...")  # Debug print
                    # Still loading, check again
                    threading.Timer(0.1, check).start()
            else:
                print("DEBUG: data_loaded attribute not found, checking again...")  # Debug print
                # Still loading, check again
                threading.Timer(0.1, check).start()
        
        # Start checking
        threading.Timer(0.1, check).start()
    
    def refresh_table(self):
        """Refresh the shipment table with current data"""
        print(f"DEBUG refresh_table called with {len(self.filtered_data)} filtered shipments")
        if not dpg.does_item_exist("shipment_table"):
            print("DEBUG: Table doesn't exist!")
            return
            
        # Clear all existing rows (only data rows, not the header)
        try:
            table_children = dpg.get_item_children("shipment_table", slot=1)
            if table_children:
                print(f"DEBUG: Clearing {len(table_children)} existing table rows")
                for child in table_children:
                    dpg.delete_item(child)
        except Exception as e:
            print(f"DEBUG: Error clearing table: {e}")
        
        # Add data rows directly - no special master checkbox row
        print(f"DEBUG: Adding {len(self.filtered_data)} data rows")
        
        # Data rows only
        for i, shipment in enumerate(self.filtered_data):
            # Find original index for selection callback
            original_index = next((j for j, orig in enumerate(self.shipment_data) 
                                 if orig['shipment_nbr'] == shipment['shipment_nbr']), i)
            
            with dpg.table_row(parent="shipment_table"):
                with dpg.table_cell():
                    dpg.add_checkbox(
                        default_value=shipment['selected'],
                        callback=self.toggle_shipment_selection,
                        user_data=original_index,
                        tag=f"checkbox_{i}"
                    )
                with dpg.table_cell():
                    dpg.add_text(shipment['shipment_nbr'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['do_numbers'][:50] + "..." if len(shipment['do_numbers']) > 50 else shipment['do_numbers'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(str(shipment['do_count']), color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['ship_to'][:25] + "..." if len(shipment['ship_to']) > 25 else shipment['ship_to'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['po'][:25] + "..." if len(shipment['po']) > 25 else shipment['po'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['vas'][:15] + "..." if len(shipment['vas']) > 15 else shipment['vas'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['original_qty'], color=[34, 139, 34])
                with dpg.table_cell():
                    dpg.add_text(shipment['label_type'][:15] + "..." if len(shipment['label_type']) > 15 else shipment['label_type'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['order_type'][:15] + "..." if len(shipment['order_type']) > 15 else shipment['order_type'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['pmt_term'][:15] + "..." if len(shipment['pmt_term']) > 15 else shipment['pmt_term'], color=[240, 240, 240])
                with dpg.table_cell():
                    dpg.add_text(shipment['start_ship'][:15] + "..." if len(shipment['start_ship']) > 15 else shipment['start_ship'], color=[240, 240, 240])
        
        print("DEBUG: Table refresh completed")
    
    def update_individual_checkboxes(self):
        """Update individual row checkboxes to match their data state"""
        print(f"DEBUG: update_individual_checkboxes called for {len(self.filtered_data)} shipments")
        updated_count = 0
        for i, shipment in enumerate(self.filtered_data):
            checkbox_tag = f"checkbox_{i}"
            if dpg.does_item_exist(checkbox_tag):
                dpg.set_value(checkbox_tag, shipment['selected'])
                updated_count += 1
                print(f"DEBUG: Updated checkbox {checkbox_tag} to {shipment['selected']}")
            else:
                print(f"DEBUG: Checkbox {checkbox_tag} does not exist!")
        print(f"DEBUG: Updated {updated_count} checkboxes")
    
    def toggle_shipment_selection(self, sender, app_data, user_data):
        """Toggle selection of a specific shipment"""
        index = user_data
        self.shipment_data[index]['selected'] = app_data
        
        shipment_nbr = self.shipment_data[index]['shipment_nbr']
        
        # Also update the filtered data
        for filtered_shipment in self.filtered_data:
            if filtered_shipment['shipment_nbr'] == shipment_nbr:
                filtered_shipment['selected'] = app_data
                break
        
        if app_data:
            self.selected_shipments.add(shipment_nbr)
        else:
            self.selected_shipments.discard(shipment_nbr)
        
        self.update_selection_count()
        self.update_master_checkbox()
    
    def update_master_checkbox(self):
        """Update master checkbox state based on current selections - no longer needed but kept for compatibility"""
        pass
    
    def select_all_callback(self, sender, app_data):
        """Select all visible (filtered) shipments"""
        print(f"DEBUG: select_all_callback triggered with {len(self.filtered_data)} filtered shipments")
        self.log_to_console(f"Selecting all {len(self.filtered_data)} visible shipments", "info")
        
        for shipment in self.filtered_data:
            shipment_nbr = shipment['shipment_nbr']
            # Update both filtered and original data
            for orig_shipment in self.shipment_data:
                if orig_shipment['shipment_nbr'] == shipment_nbr:
                    orig_shipment['selected'] = True
                    break
            shipment['selected'] = True
            self.selected_shipments.add(shipment_nbr)
        
        self.update_individual_checkboxes()
        self.update_selection_count()
        self.update_status(f"Selected all {len(self.filtered_data)} visible shipments", "success")
    
    def deselect_all_callback(self, sender, app_data):
        """Deselect all visible (filtered) shipments"""
        print(f"DEBUG: deselect_all_callback triggered with {len(self.filtered_data)} filtered shipments")
        self.log_to_console(f"Deselecting all {len(self.filtered_data)} visible shipments", "info")
        
        for shipment in self.filtered_data:
            shipment_nbr = shipment['shipment_nbr']
            # Update both filtered and original data
            for orig_shipment in self.shipment_data:
                if orig_shipment['shipment_nbr'] == shipment_nbr:
                    orig_shipment['selected'] = False
                    break
            shipment['selected'] = False
            self.selected_shipments.discard(shipment_nbr)
        
        self.update_individual_checkboxes()
        self.update_selection_count()
        self.update_status(f"Deselected all {len(self.filtered_data)} visible shipments", "success")
    
    def toggle_select_all(self, sender, app_data):
        """Toggle selection of all shipments"""
        select_all = app_data
        for shipment in self.shipment_data:
            shipment['selected'] = select_all
        
        if select_all:
            self.selected_shipments = set(s['shipment_nbr'] for s in self.shipment_data)
        else:
            self.selected_shipments.clear()
        
        self.refresh_table()
        self.update_selection_count()
    
    def update_selection_count(self):
        """Update the selection count display"""
        count = len(self.selected_shipments)
        total = len(self.filtered_data)
        
        # Calculate total units for selected shipments
        total_units = 0
        for shipment in self.shipment_data:
            if shipment['selected']:
                try:
                    units = int(shipment['original_qty']) if shipment['original_qty'].isdigit() else 0
                    total_units += units
                except (ValueError, AttributeError):
                    pass
        
        self.selected_total_units = total_units
        
        dpg.set_value("selection_count", f"{count} of {total} selected")
        if dpg.does_item_exist("total_units"):
            dpg.set_value("total_units", f"{total_units:,} units")
    
    def generate_selected_callback(self, sender, app_data):
        """Generate placards for selected shipments"""
        if not self.selected_shipments:
            self.update_status("No shipments selected", "warning")
            return
        
        if self.is_processing:
            self.update_status("Processing already in progress", "warning")
            return
        
        self.log_to_console(f"Starting placard generation for {len(self.selected_shipments)} selected shipments", "info")
        # Start processing in separate thread
        threading.Thread(
            target=self.process_shipments_thread,
            args=(list(self.selected_shipments),),
            daemon=True
        ).start()
    
    def generate_all_callback(self, sender, app_data):
        """Generate placards for all shipments"""
        if not self.shipment_data:
            self.update_status("No data loaded", "warning")
            return
        
        if self.is_processing:
            self.update_status("Processing already in progress", "warning")
            return
        
        all_shipments = [s['shipment_nbr'] for s in self.shipment_data]
        self.log_to_console(f"Starting placard generation for ALL {len(all_shipments)} shipments", "info")
        
        # Start processing in separate thread
        threading.Thread(
            target=self.process_shipments_thread,
            args=(all_shipments,),
            daemon=True
        ).start()
    
    def process_shipments_thread(self, shipments: List[str]):
        """Process shipments in a separate thread with comprehensive security and error handling"""
        # Thread safety - use lock to prevent concurrent processing
        with self._processing_lock:
            if self.is_processing:
                self.log_to_console("Processing already in progress, skipping request", "warning")
                return
            self.is_processing = True
        
        try:
            self.log_to_console(f"Starting background processing thread for {len(shipments)} shipments", "info")
            
            # Security validation for shipment list
            if len(shipments) > SecurityConfig.MAX_RECORDS_PER_BATCH:
                raise SecurityError(f"Too many shipments requested: {len(shipments)} > {SecurityConfig.MAX_RECORDS_PER_BATCH}")
            
            # Validate each shipment number
            validated_shipments = []
            for shipment in shipments:
                if InputValidator.validate_shipment_number(shipment):
                    validated_shipments.append(shipment)
                else:
                    security_logger.warning(f"Invalid shipment number filtered: {shipment}")
            
            if not validated_shipments:
                raise SecurityError("No valid shipment numbers provided")
            
            shipments = validated_shipments
            self.log_to_console(f"Validated {len(shipments)} shipment numbers", "info")
            
            # Validate prerequisites
            if self.generator.df is None:
                raise RuntimeError("No data loaded - cannot process shipments")
            
            # Check output directory
            output_dir = os.path.join(os.getcwd(), "Placards")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                self.log_to_console(f"Created output directory: {output_dir}", "info")
            
            # Update UI to show processing state
            self._safe_dpg_operation("disable_generate_buttons", 
                                     dpg.configure_item, "generate_selected_btn", enabled=False)
            self._safe_dpg_operation("disable_generate_all_button", 
                                     dpg.configure_item, "generate_all_btn", enabled=False)
            self._safe_dpg_operation("reset_progress_bar", 
                                     dpg.set_value, "progress_bar", 0.0)
            self._safe_dpg_operation("show_progress_container", 
                                     dpg.configure_item, "progress_container", show=True)
            
            successful_count = 0
            failed_count = 0
            failed_shipments = []
            
            for i, shipment_num in enumerate(shipments):
                try:
                    # Update progress with safe operations
                    progress = (i + 1) / len(shipments)
                    self._safe_dpg_operation("update_progress_bar", 
                                             dpg.set_value, "progress_bar", progress)
                    self._safe_dpg_operation("update_progress_text", 
                                             dpg.set_value, "progress_text", 
                                             f"Processing {i + 1} of {len(shipments)}: {shipment_num}")
                    
                    self.update_status(f"Processing shipment {shipment_num} ({i + 1}/{len(shipments)})", "info")
                    
                    # Validate shipment number format
                    if not shipment_num.isdigit() or len(shipment_num) != 10:
                        raise ValueError(f"Invalid shipment number format: {shipment_num}")
                    
                    # Process the shipment with timeout protection
                    start_time = time.time()
                    if self.generator.process_shipment(shipment_num):
                        processing_time = time.time() - start_time
                        successful_count += 1
                        self.log_to_console(f"Successfully processed shipment {shipment_num} in {processing_time:.2f}s", "success")
                    else:
                        failed_count += 1
                        failed_shipments.append(shipment_num)
                        self.log_to_console(f"Failed to process shipment {shipment_num} - no data found", "error")
                    
                    # Small delay to make progress visible and prevent UI freezing
                    time.sleep(0.1)
                    
                except Exception as e:
                    failed_count += 1
                    failed_shipments.append(shipment_num)
                    self.log_to_console(f"Error processing shipment {shipment_num}: {str(e)}", "error")
                    continue
            
            # Update final status with detailed reporting
            total = len(shipments)
            if failed_count == 0:
                self.update_status(f"✅ Processing complete: All {successful_count} shipments processed successfully", "success")
                self.log_to_console(f"BATCH COMPLETE: {successful_count}/{total} shipments processed successfully", "success")
            else:
                self.update_status(f"⚠️ Processing complete: {successful_count} successful, {failed_count} failed", "warning")
                self.log_to_console(f"BATCH COMPLETE: {successful_count}/{total} successful, {failed_count} failed", "warning")
                
                # Log failed shipments for troubleshooting
                if failed_shipments:
                    self.log_to_console(f"Failed shipments: {', '.join(failed_shipments[:10])}", "error")
                    if len(failed_shipments) > 10:
                        self.log_to_console(f"... and {len(failed_shipments) - 10} more failed shipments", "error")
            
            # Log performance metrics
            processing_rate = len(shipments) / max(1, time.time() - start_time) if 'start_time' in locals() else 0
            self.log_to_console(f"Processing rate: {processing_rate:.2f} shipments/second", "info")
            
        except SecurityError as e:
            self.update_status(f"❌ Security error during processing: {str(e)}", "error")
            self.log_to_console(f"SECURITY ERROR in processing thread: {str(e)}", "error")
            security_logger.error(f"Processing security violation: {e}")
            
        except Exception as e:
            self.update_status(f"❌ Critical error during processing: {str(e)}", "error")
            self.log_to_console(f"CRITICAL ERROR in processing thread: {str(e)}", "error")
            self.log_to_console(f"Stack trace: {traceback.format_exc()}", "error")
            gui_logger.error(f"Processing thread failed: {e}", exc_info=True)
            
        finally:
            # Re-enable buttons and hide progress with safe operations
            self._safe_dpg_operation("enable_generate_selected", 
                                     dpg.configure_item, "generate_selected_btn", enabled=True)
            self._safe_dpg_operation("enable_generate_all", 
                                     dpg.configure_item, "generate_all_btn", enabled=True)
            self._safe_dpg_operation("hide_progress_container", 
                                     dpg.configure_item, "progress_container", show=False)
            
            # Thread safety - reset processing state
            with self._processing_lock:
                self.is_processing = False
            
            self.log_to_console("Processing thread completed", "info")
    
    def create_main_window(self):
        """Create the main application window with properly centered design"""
        with dpg.window(label="LOGISTICS DOCUMENT GENERATOR", tag="main_window", no_scrollbar=True, no_scroll_with_mouse=True):
            
            dpg.add_spacer(height=15)
            
            # Title - centered using drawlist positioning
            with dpg.child_window(height=40, border=False, tag="title_container", no_scrollbar=True, no_scroll_with_mouse=True):
                dpg.add_text("LOGISTICS DOCUMENT GENERATOR", color=[70, 130, 200], tag="main_title", pos=[0, 10])
                self._safe_bind_font(dpg.last_item(), self.large_font)
            
            dpg.add_spacer(height=10)
            
            # Control bar - centered using child window
            with dpg.child_window(height=50, border=False, tag="control_container", no_scrollbar=True, no_scroll_with_mouse=True):
                with dpg.group(horizontal=True, tag="control_group", pos=[0, 10]):
                    # Action buttons
                    dpg.add_button(
                        label="LOAD DATA",
                        tag="load_data_btn",
                        callback=self.load_data_callback,
                        width=120,
                        height=35
                    )
                    dpg.bind_item_theme(dpg.last_item(), self.primary_theme)
                    
                    dpg.add_spacer(width=20)
                    
                    dpg.add_button(
                        label="CLEAR FILTERS",
                        callback=self.clear_all_filters_callback,
                        width=120,
                        height=35
                    )
                    dpg.bind_item_theme(dpg.last_item(), self.danger_theme)
                    
                    dpg.add_spacer(width=20)
                    
                    # Search
                    dpg.add_input_text(
                        hint="Search shipments, destinations...",
                        callback=self.search_callback,
                        width=400,
                        tag="main_search_input"
                    )
                    
                    dpg.add_spacer(width=20)
                    
                    dpg.add_spacer(width=20)
                    
                    # Selection controls
                    dpg.add_button(
                        label="SELECT ALL",
                        callback=self.select_all_callback,
                        width=110,
                        height=35,
                        enabled=False,
                        tag="select_all_btn"
                    )
                    dpg.bind_item_theme(dpg.last_item(), self.success_theme)
                    
                    dpg.add_spacer(width=10)
                    
                    dpg.add_button(
                        label="DESELECT ALL",
                        callback=self.deselect_all_callback,
                        width=130,
                        height=35,
                        enabled=False,
                        tag="deselect_all_btn"
                    )
                    dpg.bind_item_theme(dpg.last_item(), self.warning_theme)
            
            dpg.add_spacer(height=10)
            
            # Stats bar - centered using child window
            with dpg.child_window(height=30, border=False, tag="stats_container", no_scrollbar=True, no_scroll_with_mouse=True):
                with dpg.group(horizontal=True, tag="stats_group", pos=[0, 5]):
                    dpg.add_text("SELECTED:", color=[140, 140, 140])
                    self._safe_bind_font(dpg.last_item(), self.bold_font)
                    
                    dpg.add_spacer(width=5)
                    
                    dpg.add_text("0 of 0", tag="selection_count", color=[70, 130, 200])
                    self._safe_bind_font(dpg.last_item(), self.bold_font)
                    
                    dpg.add_spacer(width=25)
                    
                    dpg.add_text("UNITS:", color=[140, 140, 140])
                    self._safe_bind_font(dpg.last_item(), self.bold_font)
                    
                    dpg.add_spacer(width=5)
                    
                    dpg.add_text("0", tag="total_units", color=[34, 139, 34])
                    self._safe_bind_font(dpg.last_item(), self.bold_font)
            
            dpg.add_spacer(height=15)
            
            # Table - full width
            with dpg.child_window(height=420, border=True):
                with dpg.table(
                    tag="shipment_table",
                    header_row=True,
                    borders_innerH=True,
                    borders_outerH=True,
                    borders_innerV=True,
                    borders_outerV=True,
                    row_background=True,
                    scrollX=True,
                    scrollY=True,
                    resizable=True,
                    policy=dpg.mvTable_SizingStretchProp
                ):
                    dpg.add_table_column(label="SELECT", width_fixed=True, init_width_or_weight=75)
                    dpg.add_table_column(label="SHIPMENT", init_width_or_weight=0.1)
                    dpg.add_table_column(label="DO NUMBERS", init_width_or_weight=0.15)
                    dpg.add_table_column(label="COUNT", init_width_or_weight=0.06)
                    dpg.add_table_column(label="SHIP TO", init_width_or_weight=0.22)
                    dpg.add_table_column(label="PO", init_width_or_weight=0.15)
                    dpg.add_table_column(label="VAS", init_width_or_weight=0.05)
                    dpg.add_table_column(label="QTY", init_width_or_weight=0.08)
                    dpg.add_table_column(label="LABEL", init_width_or_weight=0.07)
                    dpg.add_table_column(label="ORDER", init_width_or_weight=0.07)
                    dpg.add_table_column(label="PAYMENT", init_width_or_weight=0.08)
                    dpg.add_table_column(label="START SHIP", init_width_or_weight=0.12)
            
            dpg.add_spacer(height=20)
            
            # Action buttons - centered using child window
            with dpg.child_window(height=50, border=False, tag="action_container", no_scrollbar=True, no_scroll_with_mouse=True):
                with dpg.group(horizontal=True, tag="action_group", pos=[0, 5]):
                    dpg.add_button(
                        label="GENERATE SELECTED",
                        tag="generate_selected_btn",
                        callback=self.generate_selected_callback,
                        enabled=False,
                        width=180,
                        height=40
                    )
                    dpg.bind_item_theme(dpg.last_item(), self.success_theme)
                    
                    dpg.add_spacer(width=20)
                    
                    dpg.add_button(
                        label="GENERATE ALL",
                        tag="generate_all_btn",
                        callback=self.generate_all_callback,
                        enabled=False,
                        width=140,
                        height=40
                    )
                    dpg.bind_item_theme(dpg.last_item(), self.primary_theme)
            
            dpg.add_spacer(height=20)
            
            # Progress - centered using child window
            with dpg.child_window(height=30, border=False, tag="progress_container_outer", no_scrollbar=True, no_scroll_with_mouse=True):
                with dpg.group(tag="progress_container", show=False, horizontal=True, pos=[0, 3]):
                    dpg.add_progress_bar(tag="progress_bar", width=300, height=25)
                    dpg.add_spacer(width=10)
                    dpg.add_text("Processing...", tag="progress_text", color=[70, 130, 200])
                    self._safe_bind_font(dpg.last_item(), self.bold_font)
            
            dpg.add_spacer(height=15)
            
            # Status - centered using child window
            with dpg.child_window(height=40, border=False, tag="status_container", no_scrollbar=True, no_scroll_with_mouse=True):
                with dpg.child_window(height=30, width=400, border=True, tag="status_inner", pos=[0, 5], no_scrollbar=True, no_scroll_with_mouse=True):
                    with dpg.group(horizontal=True):
                        dpg.add_spacer(width=1)  # Flexible spacer inside status window
                        dpg.add_text("STATUS: Ready", tag="status_text", color=[70, 130, 200])
                        self._safe_bind_font(dpg.last_item(), self.bold_font)
                        dpg.add_spacer(width=1)  # Flexible spacer inside status window
            
            dpg.add_spacer(height=10)
            
            # Console section
            with dpg.group(horizontal=True):
                dpg.add_text("CONSOLE", color=[140, 140, 140])
                self._safe_bind_font(dpg.last_item(), self.bold_font)
                
                dpg.add_spacer(width=10)
                
                dpg.add_button(
                    label="CLEAR CONSOLE",
                    callback=self.clear_console_callback,
                    width=130,
                    height=30
                )
                dpg.bind_item_theme(dpg.last_item(), self.danger_theme)
            
            dpg.add_spacer(height=5)
            
            # Console window - full width
            with dpg.child_window(height=150, border=True, tag="console_window"):
                dpg.add_input_text(
                    tag="console_text",
                    multiline=True,
                    readonly=True,
                    width=-1,
                    height=-1,
                    default_value="Console initialized - application starting...\n"
                )
            
            # Register resize callback to recalculate centering
            dpg.set_viewport_resize_callback(self.on_viewport_resize)
    
    def on_viewport_resize(self):
        """Callback to recenter content when viewport is resized"""
        self.center_content()
    
    def center_content(self):
        """Dynamically center all content by calculating proper positions"""
        try:
            # Get viewport and window width
            viewport_width = dpg.get_viewport_width()
            window_width = dpg.get_item_width("main_window") if dpg.does_item_exist("main_window") else None
            if window_width is None:
                window_width = viewport_width
            
            # Use window width for calculations (accounting for window padding)
            available_width = window_width - 40  # Account for window padding
            
            # Title centering
            if dpg.does_item_exist("main_title"):
                title_width = 420  # More accurate width of title text
                title_x = max(0, (available_width - title_width) // 2)
                dpg.configure_item("main_title", pos=[title_x, 10])
            
            # Control bar centering
            if dpg.does_item_exist("control_group"):
                # Calculate actual control bar width:
                # LOAD DATA(120) + spacer(20) + CLEAR FILTERS(120) + spacer(20) + 
                # Search(400) + spacer(20) + spacer(20) + SELECT ALL(110) + spacer(10) + DESELECT ALL(130)
                control_bar_width = 120 + 20 + 120 + 20 + 400 + 20 + 20 + 110 + 10 + 130  # = 970
                control_x = max(0, (available_width - control_bar_width) // 2)
                dpg.configure_item("control_group", pos=[control_x, 10])
            
            # Stats bar centering
            if dpg.does_item_exist("stats_group"):
                stats_bar_width = 200  # Approximate width of stats elements
                stats_x = max(0, (available_width - stats_bar_width) // 2)
                dpg.configure_item("stats_group", pos=[stats_x, 5])
            
            # Action bar centering
            if dpg.does_item_exist("action_group"):
                action_bar_width = 320  # Approximate width of action buttons
                action_x = max(0, (available_width - action_bar_width) // 2)
                dpg.configure_item("action_group", pos=[action_x, 5])
            
            # Progress bar centering
            if dpg.does_item_exist("progress_container"):
                progress_bar_width = 400  # Approximate width of progress elements
                progress_x = max(0, (available_width - progress_bar_width) // 2)
                dpg.configure_item("progress_container", pos=[progress_x, 3])
            
            # Status bar centering
            if dpg.does_item_exist("status_inner"):
                status_bar_width = 400  # Width of status window
                status_x = max(0, (available_width - status_bar_width) // 2)
                dpg.configure_item("status_inner", pos=[status_x, 5])
                
        except Exception as e:
            print(f"Error centering content: {e}")

    # Essential functionality methods  
    def sort_table(self, direction):
        """Sort the entire table by shipment number"""
        if not self.filtered_data:
            return
            
        try:
            reverse = (direction == "desc")
            self.filtered_data.sort(key=lambda x: str(x.get('shipment_nbr', '')), reverse=reverse)
            self.refresh_table()
            self.update_status("Sorted", "success")
        except Exception as e:
            self.update_status(f"Sort error: {str(e)}", "error")
        
    def copy_selected_to_clipboard(self, sender, app_data):
        """Copy selected rows to clipboard"""
        if not self.selected_shipments:
            self.update_status("No rows selected for copying", "warning")
            return
            
        try:
            # Gather selected data
            selected_data = [ship for ship in self.filtered_data 
                           if ship['shipment_nbr'] in self.selected_shipments]
            
            if not selected_data:
                self.update_status("No data to copy", "warning")
                return
                
            # Create clipboard text (tab-delimited)
            headers = ["Shipment #", "DO Numbers", "Count", "Ship To", "PO", "VAS", 
                      "Quantity", "Label Type", "Order Type", "Payment Terms", "Start Ship"]
            
            clipboard_text = "\t".join(headers) + "\n"
            
            for ship in selected_data:
                row_data = [
                    ship.get('shipment_nbr', ''),
                    ship.get('do_numbers', ''),
                    str(ship.get('do_count', '')),
                    ship.get('ship_to', ''),
                    ship.get('po', ''),
                    ship.get('vas', ''),
                    str(ship.get('original_qty', '')),
                    ship.get('label_type', ''),
                    ship.get('order_type', ''),
                    ship.get('pmt_term', ''),
                    ship.get('start_ship', '')
                ]
                clipboard_text += "\t".join(row_data) + "\n"
            
            # Copy to clipboard
            dpg.set_clipboard_text(clipboard_text)
            self.update_status(f"Copied {len(selected_data)} rows", "success")
                
        except Exception as e:
            self.update_status(f"Copy error: {str(e)}", "error")
    
    def export_to_excel(self, sender, app_data):
        """Export current view to Excel file"""
        try:
            if not self.filtered_data:
                self.update_status("No data to export", "warning")
                return
                
            # Create DataFrame from filtered data
            export_data = []
            for ship in self.filtered_data:
                export_data.append({
                    'Shipment #': ship.get('shipment_nbr', ''),
                    'DO Numbers': ship.get('do_numbers', ''),
                    'DO Count': ship.get('do_count', ''),
                    'Ship To': ship.get('ship_to', ''),
                    'Purchase Orders': ship.get('po', ''),
                    'VAS': ship.get('vas', ''),
                    'Original Quantity': ship.get('original_qty', ''),
                    'Label Type': ship.get('label_type', ''),
                    'Order Type': ship.get('order_type', ''),
                    'Payment Terms': ship.get('pmt_term', ''),
                    'Start Ship Date': ship.get('start_ship', '')
                })
            
            df = pd.DataFrame(export_data)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"logistics_export_{timestamp}.xlsx"
            filepath = os.path.join(os.getcwd(), filename)
            
            # Export to Excel
            df.to_excel(filepath, index=False, engine='openpyxl')
            self.update_status(f"Exported {len(export_data)} rows to {filename}", "success")
            
        except Exception as e:
            self.update_status(f"Export error: {str(e)}", "error")
    
    def run(self):
        """Run the GUI application"""
        try:
            # Setup directories
            self.generator.setup_directories()
            self.generator.initialize_log()
            
            # Create main window
            self.create_main_window()
            
            # Setup viewport with sharp, modern design (increased height for console)
            dpg.create_viewport(
                title="LOGISTICS DOCUMENT GENERATOR",
                width=1600,
                height=1200,  # Increased to accommodate console
                min_width=1400,
                min_height=900,  # Increased minimum height
                resizable=True,
                always_on_top=False
            )
            
            # Apply themes
            dpg.bind_theme(self.main_theme)
            if self.default_font is not None:
                dpg.bind_font(self.default_font)
            
            # Setup and show
            dpg.setup_dearpygui()
            dpg.show_viewport()
            dpg.set_primary_window("main_window", True)
            
            # Center content after setup
            dpg.set_frame_callback(1, self.center_content)  # Center after first frame
            
            self.update_status("Ready", "success")
            
            # Initialize console with welcome messages
            self.log_to_console("Logistics Document Generator initialized", "success")
            self.log_to_console("Ready to load Excel data from Data folder", "info")
            self.log_to_console("Console logging enabled - all activities will be tracked here", "info")
            
            # Start main loop
            dpg.start_dearpygui()
            
        except Exception as e:
            print(f"Error starting GUI: {e}")
            
        finally:
            dpg.destroy_context()


def main():
    """Main entry point for the GUI application"""
    try:
        app = PlacardGeneratorGUI()
        app.run()
    except Exception as e:
        print(f"Failed to start application: {e}")
        input("Press Enter to exit...")


if __name__ == "__main__":
    main() 