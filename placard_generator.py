#!/usr/bin/env python3
"""
Logistics Document Generator

This script generates multi-page shipping placards from Excel data using Word templates.
Reads from Data folder, uses Template folder for templates, outputs to Placards folder.
"""

import os
import sys
import glob
import re
from datetime import datetime
from typing import List, Dict, Optional, Tuple, Any, cast

import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


class PlacardGenerator:
    """High-performance logistics document generator"""
    
    def __init__(self):
        self.df: Optional[pd.DataFrame] = None
        self.required_columns = [
            'Shipment Nbr', 'DO #', 'Label Type', 'Order Type', 
            'Pmt Term', 'Start Ship', 'VAS', 'Ship To', 'PO', 'Original Qty'
        ]
        self.data_folder = "Data"
        self.template_folder = "Template"
        self.output_folder = "Placards"
        
    def setup_directories(self) -> bool:
        """Ensure required directories exist"""
        try:
            for folder in [self.data_folder, self.template_folder, self.output_folder]:
                if not os.path.exists(folder):
                    os.makedirs(folder)
                    print(f"Created directory: {folder}")
            return True
        except Exception as e:
            print(f"Error creating directories: {e}")
            return False
    
    def find_excel_file(self) -> Optional[str]:
        """Find Excel file starting with 'WM-SPN-CUS105 Open Order Report' in Data folder"""
        pattern = os.path.join(self.data_folder, "WM-SPN-CUS105 Open Order Report*.xlsx")
        files = glob.glob(pattern)
        
        if not files:
            # Also check for .xls files
            pattern = os.path.join(self.data_folder, "WM-SPN-CUS105 Open Order Report*.xls")
            files = glob.glob(pattern)
        
        if not files:
            print(f"ERROR: No Excel file found in '{self.data_folder}' folder starting with 'WM-SPN-CUS105 Open Order Report'")
            return None
        
        if len(files) > 1:
            print(f"Multiple Excel files found. Using: {files[0]}")
        
        return files[0]
    
    def validate_do_number(self, do_num: Any) -> bool:
        """Validate DO # format: at least 8 digits (adjusted for actual data)"""
        if pd.isna(do_num):
            return False
        do_str = str(do_num).strip()
        return bool(re.match(r'^\d{8,}$', do_str))
    
    def validate_shipment_number(self, shipment_num: Any) -> bool:
        """Validate shipment number: exactly 10 characters, all numbers"""
        if not shipment_num:
            return False
        shipment_str = str(shipment_num).strip()
        return len(shipment_str) == 10 and shipment_str.isdigit()
    
    def load_and_prepare_data(self) -> bool:
        """Load Excel file and prepare data with validation"""
        print("Loading and preparing data...")
        
        # Find Excel file
        excel_file = self.find_excel_file()
        if not excel_file:
            return False
        
        try:
            # Load Excel file
            print(f"Loading file: {excel_file}")
            df = pd.read_excel(excel_file)
            print(f"Loaded {len(df)} rows from Excel file")
            
            # Check for required columns
            missing_columns = [col for col in self.required_columns if col not in df.columns]
            if missing_columns:
                print(f"ERROR: Missing required columns: {missing_columns}")
                print(f"Available columns: {list(df.columns)}")
                return False
            
            # Filter out rows with empty Shipment Nbr
            initial_count = len(df)
            df = df[df['Shipment Nbr'].notna()]
            # Cast to Series to access .str accessor
            shipment_series = cast(pd.Series, df['Shipment Nbr'].astype(str))
            df = df[shipment_series.str.strip() != '']
            print(f"Removed {initial_count - len(df)} rows with empty Shipment Nbr")
            
            # Validate DO # format (exactly 10 digits)
            before_do_filter = len(df)
            # Cast to Series to access .apply method
            do_series = cast(pd.Series, df['DO #'])
            df = df[do_series.apply(self.validate_do_number)]
            print(f"Removed {before_do_filter - len(df)} rows with invalid DO # format")
            
            # Assign to instance variable - cast to DataFrame to satisfy type checker
            self.df = cast(pd.DataFrame, df)
            if self.df is not None:
                print(f"Final dataset: {len(self.df)} rows ready for processing")
            return True
            
        except Exception as e:
            print(f"ERROR loading Excel file: {e}")
            return False
    
    def format_date(self, date_value: Any) -> str:
        """Format date as MM/DD/YYYY"""
        if pd.isna(date_value):
            return ""
        
        try:
            if isinstance(date_value, str):
                # Try to parse string date
                date_obj = pd.to_datetime(date_value)
            else:
                date_obj = pd.to_datetime(date_value)
            return date_obj.strftime("%m/%d/%Y")
        except:
            return str(date_value)
    
    def get_vas_value(self, vas_raw: Any) -> str:
        """Convert VAS value to 'VAS' or 'NOT VAS'"""
        if pd.isna(vas_raw):
            return "NOT VAS"
        vas_str = str(vas_raw).strip().upper()
        return "VAS" if vas_str == "Y" else "NOT VAS"
    
    def replace_placeholders_in_document(self, doc: Any, replacements: Dict[str, str]) -> None:
        """Replace all placeholders in the document while preserving formatting"""
        
        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            for placeholder, value in replacements.items():
                if placeholder in paragraph.text:
                    self.replace_placeholder_in_paragraph(paragraph, placeholder, value)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for placeholder, value in replacements.items():
                            if placeholder in paragraph.text:
                                self.replace_placeholder_in_paragraph(paragraph, placeholder, value)
        
        # Replace in headers
        for section in doc.sections:
            header = section.header
            for paragraph in header.paragraphs:
                for placeholder, value in replacements.items():
                    if placeholder in paragraph.text:
                        self.replace_placeholder_in_paragraph(paragraph, placeholder, value)
        
        # Replace in footers
        for section in doc.sections:
            footer = section.footer
            for paragraph in footer.paragraphs:
                for placeholder, value in replacements.items():
                    if placeholder in paragraph.text:
                        self.replace_placeholder_in_paragraph(paragraph, placeholder, value)
    
    def replace_placeholder_in_paragraph(self, paragraph: Any, placeholder: str, replacement: str) -> None:
        """Replace placeholder while preserving mixed formatting"""
        if placeholder not in paragraph.text:
            return
        
        # Try simple replacement first (best case - placeholder is in one run)
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)
                return
        
        # Complex case: placeholder spans multiple runs
        # We need to reconstruct the paragraph preserving different formatting for different parts
        self.replace_across_runs(paragraph, placeholder, replacement)
    
    def replace_across_runs(self, paragraph: Any, placeholder: str, replacement: str) -> None:
        """Replace placeholder that spans across multiple runs while preserving original formatting"""
        
        # Store all original runs with their formatting
        original_runs = []
        for run in paragraph.runs:
            original_runs.append({
                'text': run.text,
                'bold': run.bold,
                'italic': run.italic,
                'underline': run.underline,
                'font_name': run.font.name,
                'font_size': run.font.size,
                'font_color': run.font.color.rgb if run.font.color.rgb else None
            })
        
        # Get the full text and find placeholder position
        full_text = paragraph.text
        placeholder_start = full_text.find(placeholder)
        if placeholder_start == -1:
            return
        
        placeholder_end = placeholder_start + len(placeholder)
        
        # Split text into three parts: before placeholder, placeholder, after placeholder
        before_text = full_text[:placeholder_start]
        after_text = full_text[placeholder_end:]
        
        # Find the run that contains most of the placeholder for formatting reference
        char_pos = 0
        placeholder_run_format = None
        
        for run_info in original_runs:
            run_end = char_pos + len(run_info['text'])
            if char_pos <= placeholder_start < run_end:
                placeholder_run_format = run_info
                break
            char_pos = run_end
        
        # Fallback to first run if we can't find the placeholder run
        if placeholder_run_format is None and original_runs:
            placeholder_run_format = original_runs[0]
        
        # Clear the paragraph
        paragraph.clear()
        
        # Method 1: Try to reconstruct the original structure intelligently
        if before_text or after_text:
            self.reconstruct_mixed_formatting(paragraph, before_text, replacement, after_text, original_runs, placeholder_run_format)
        else:
            # Simple case: just the placeholder
            new_run = paragraph.add_run(replacement)
            if placeholder_run_format:
                self.apply_run_formatting(new_run, placeholder_run_format)
    
    def reconstruct_mixed_formatting(self, paragraph: Any, before_text: str, replacement: str, after_text: str, original_runs: List[Dict[str, Any]], placeholder_format: Optional[Dict[str, Any]]) -> None:
        """Reconstruct paragraph with mixed formatting preserved"""
        
        # For now, use a simplified approach that preserves the most important formatting
        # Add before text (try to preserve original formatting)
        if before_text:
            # Use the first run's formatting for the before text
            before_run = paragraph.add_run(before_text)
            if original_runs:
                self.apply_run_formatting(before_run, original_runs[0])
        
        # Add replacement text with placeholder formatting  
        replacement_run = paragraph.add_run(replacement)
        if placeholder_format:
            self.apply_run_formatting(replacement_run, placeholder_format)
        
        # Add after text (try to preserve original formatting)
        if after_text:
            # Use the last run's formatting for the after text
            after_run = paragraph.add_run(after_text)
            if original_runs:
                self.apply_run_formatting(after_run, original_runs[-1])
    
    def apply_run_formatting(self, run: Any, format_info: Dict[str, Any]) -> None:
        """Apply formatting from format_info dictionary to a run"""
        if format_info.get('bold') is not None:
            run.bold = format_info['bold']
        if format_info.get('italic') is not None:
            run.italic = format_info['italic']
        if format_info.get('underline') is not None:
            run.underline = format_info['underline']
        if format_info.get('font_name'):
            run.font.name = format_info['font_name']
        if format_info.get('font_size'):
            run.font.size = format_info['font_size']
        if format_info.get('font_color'):
            run.font.color.rgb = format_info['font_color']
    
    def replace_in_paragraph(self, paragraph: Any, placeholder: str, replacement: str) -> None:
        """Replace placeholder in paragraph while preserving formatting"""
        if placeholder not in paragraph.text:
            return
        
        # Method 1: Try to replace within existing runs first (best formatting preservation)
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)
                return
        
        # Method 2: Handle placeholder spanning multiple runs with smart formatting detection
        full_text = paragraph.text
        if placeholder not in full_text:
            return
        
        # Find placeholder position
        placeholder_start = full_text.find(placeholder)
        placeholder_end = placeholder_start + len(placeholder)
        
        # Collect all runs and their positions
        runs_info = []
        current_pos = 0
        
        for run in paragraph.runs:
            run_start = current_pos
            run_end = current_pos + len(run.text)
            runs_info.append({
                'run': run,
                'start': run_start,
                'end': run_end,
                'text': run.text
            })
            current_pos = run_end
        
        # Find which runs contain the placeholder
        affected_runs = []
        for run_info in runs_info:
            if (run_info['start'] < placeholder_end and run_info['end'] > placeholder_start):
                affected_runs.append(run_info)
        
        if not affected_runs:
            return
        
        # Use the formatting from the run that contains the start of the placeholder
        default_run = None
        for run_info in affected_runs:
            if run_info['start'] <= placeholder_start < run_info['end']:
                default_run = run_info['run']
                break
        
        if default_run is None:
            default_run = affected_runs[0]['run']
        
        # Replace with mixed formatting preservation
        self.replace_with_mixed_formatting(paragraph, placeholder, replacement, default_run)
    
    def apply_formatting(self, run: Any, bold: Optional[bool], italic: Optional[bool], underline: Optional[bool], font_name: Optional[str], font_size: Optional[Any], font_color: Optional[Any]) -> None:
        """Apply formatting to a run"""
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        if underline is not None:
            run.underline = underline
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = font_size
        if font_color:
            run.font.color.rgb = font_color
    
    def replace_with_mixed_formatting(self, paragraph: Any, placeholder: str, replacement: str, default_run: Any) -> None:
        """Replace placeholder while preserving mixed formatting in the paragraph"""
        # Get the full paragraph text
        full_text = paragraph.text
        
        # Find placeholder position
        placeholder_start = full_text.find(placeholder)
        if placeholder_start == -1:
            return
        
        placeholder_end = placeholder_start + len(placeholder)
        
        # Split into before, placeholder, and after
        before_text = full_text[:placeholder_start]
        after_text = full_text[placeholder_end:]
        
        # Clear paragraph and rebuild with formatting
        paragraph.clear()
        
        # Add before text (preserve original formatting from first run)
        if before_text:
            before_run = paragraph.add_run(before_text)
            # Copy formatting from first original run
            if paragraph.runs:  # This will be empty since we just cleared
                pass  # We'll use default formatting
        
        # Add replacement text with default run formatting
        replacement_run = paragraph.add_run(replacement)
        self.apply_formatting(
            replacement_run,
            default_run.bold if default_run else None,
            default_run.italic if default_run else None,
            default_run.underline if default_run else None,
            default_run.font.name if default_run else None,
            default_run.font.size if default_run else None,
            default_run.font.color.rgb if default_run and default_run.font.color.rgb else None
        )
    
    def copy_template_content(self, template_path: str) -> Optional[Any]:
        """Load template and return a copy"""
        try:
            if not os.path.exists(template_path):
                print(f"ERROR: Template file not found: {template_path}")
                return None
            
            doc = Document(template_path)  # type: ignore
            return doc
        except Exception as e:
            print(f"ERROR loading template: {e}")
            return None
    
    def copy_formatted_content(self, source_doc: Any, target_doc: Any) -> None:
        """Copy all content from source document to target document while preserving formatting"""
        try:
            # Copy paragraphs with full formatting
            for paragraph in source_doc.paragraphs:
                new_paragraph = target_doc.add_paragraph()
                
                # Copy paragraph-level formatting
                new_paragraph.alignment = paragraph.alignment
                new_paragraph.paragraph_format.space_before = paragraph.paragraph_format.space_before
                new_paragraph.paragraph_format.space_after = paragraph.paragraph_format.space_after
                new_paragraph.paragraph_format.line_spacing = paragraph.paragraph_format.line_spacing
                
                # Copy all runs with their formatting
                for run in paragraph.runs:
                    new_run = new_paragraph.add_run(run.text)
                    
                    # Copy run formatting
                    new_run.bold = run.bold
                    new_run.italic = run.italic
                    new_run.underline = run.underline
                    if run.font.name:
                        new_run.font.name = run.font.name
                    if run.font.size:
                        new_run.font.size = run.font.size
                    if run.font.color.rgb:
                        new_run.font.color.rgb = run.font.color.rgb
            
            # Copy tables with formatting
            for table in source_doc.tables:
                new_table = target_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                
                # Copy table style
                if table.style:
                    new_table.style = table.style
                
                # Copy cell content and formatting
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        new_cell = new_table.cell(i, j)
                        
                        # Clear default paragraph and copy all paragraphs from source cell
                        new_cell.paragraphs[0].clear()
                        
                        for paragraph in cell.paragraphs:
                            if paragraph == cell.paragraphs[0]:
                                # Use the existing first paragraph
                                target_paragraph = new_cell.paragraphs[0]
                            else:
                                # Add new paragraph for additional ones
                                target_paragraph = new_cell.add_paragraph()
                            
                            # Copy paragraph formatting
                            target_paragraph.alignment = paragraph.alignment
                            
                            # Copy runs with formatting
                            for run in paragraph.runs:
                                new_run = target_paragraph.add_run(run.text)
                                new_run.bold = run.bold
                                new_run.italic = run.italic
                                new_run.underline = run.underline
                                if run.font.name:
                                    new_run.font.name = run.font.name
                                if run.font.size:
                                    new_run.font.size = run.font.size
                                if run.font.color.rgb:
                                    new_run.font.color.rgb = run.font.color.rgb
                                    
        except Exception as e:
            print(f"Warning: Error copying formatted content: {e}")
            # Fallback to simple text copy
            for paragraph in source_doc.paragraphs:
                target_doc.add_paragraph(paragraph.text)
            for table in source_doc.tables:
                new_table = target_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        new_table.cell(i, j).text = cell.text
    
    def process_shipment(self, shipment_num: str) -> bool:
        """Process a single shipment number and generate placard"""
        print(f"\nProcessing shipment: {shipment_num}")
        
        # Validate shipment number
        if not self.validate_shipment_number(shipment_num):
            print(f"ERROR: Invalid shipment number format: {shipment_num} (must be exactly 10 digits)")
            return False
        
        # Ensure df is not None
        if self.df is None:
            print("ERROR: No data loaded")
            return False
        
        # Find shipment data - handle float values like 9010157586.0
        # Convert both the column and search value to integers for comparison
        df_shipment_clean = self.df['Shipment Nbr'].astype(float).astype(int).astype(str)
        shipment_data = self.df[df_shipment_clean == shipment_num]
        if shipment_data.empty:
            print(f"ERROR: No data found for shipment number: {shipment_num}")
            return False
        
        print(f"Found {len(shipment_data)} records for shipment {shipment_num}")
        
        # Get template path
        template_path = os.path.join(self.template_folder, "placard_template.docx")
        
        # Create main document by copying template
        main_doc = self.copy_template_content(template_path)
        if not main_doc:
            return False
        
        # Get shipment-level data (same for all pages)
        first_row = shipment_data.iloc[0]
        shipment_level_data = {
            'Shipment Nbr': str(first_row['Shipment Nbr']),
            'Label Type': str(first_row['Label Type']),
            'Order Type': str(first_row['Order Type']),
            'Pmt Term': str(first_row['Pmt Term']),
            'Start Ship': self.format_date(first_row['Start Ship']),
            'VAS': self.get_vas_value(first_row['VAS'])
        }
        
        # Group by DO #
        do_groups = shipment_data.groupby('DO #')
        total_dos = len(do_groups)
        
        print(f"Processing {total_dos} DO #s for shipment {shipment_num}")
        
        # Process each DO # and create pages
        main_doc = None
        
        for do_index, (do_num, do_group) in enumerate(do_groups, 1):
            print(f"  Processing DO # {do_num} ({do_index}/{total_dos})")
            
            # Get page-level data
            first_do_row = do_group.iloc[0]
            
            # Aggregate POs for this DO #
            # Cast to Series to access .dropna method
            po_series = cast(pd.Series, do_group['PO'])
            unique_pos = po_series.dropna().unique()
            po_list = '\n'.join([str(po) for po in unique_pos if str(po).strip()])
            
            # Calculate total Original Qty for this DO #
            total_original_qty = do_group['Original Qty'].sum()
            
            page_level_data = {
                'DO #': str(do_num),
                'Ship To': str(first_do_row['Ship To']),
                'PO': po_list,
                'Original Qty': str(int(total_original_qty)) if not pd.isna(total_original_qty) else '0'
            }
            
            # Combine all replacement data
            replacements = {
                '{{Ship To}}': page_level_data['Ship To'],
                '{{Shipment Nbr}}': str(int(float(shipment_level_data['Shipment Nbr']))),  # Remove .0 from float
                '{{PO}}': page_level_data['PO'],
                '{{DO #}}': f"{int(page_level_data['DO #']):010d}",  # Format with leading zeros (10 digits)
                '{{VAS}}': shipment_level_data['VAS'],
                '{{Original Qty}}': page_level_data['Original Qty'] + ' Units',  # Add "Units" after quantity
                '{{Label Type}}': shipment_level_data['Label Type'],
                '{{Order Type}}': shipment_level_data['Order Type'],
                '{{Pmt Term}}': shipment_level_data['Pmt Term'],
                '{{Start Ship}}': shipment_level_data['Start Ship']
            }
            
            # Create a fresh copy of template for this page
            page_doc = Document(template_path)  # type: ignore
            
            # Replace placeholders in this page
            self.replace_placeholders_in_document(page_doc, replacements)
            
            # Handle multi-page document creation
            if do_index == 1:
                # First page: use this as the main document
                main_doc = page_doc
            else:
                # Subsequent pages: add page break and copy all content with formatting
                if main_doc is not None:
                    main_doc.add_page_break()
                    self.copy_formatted_content(page_doc, main_doc)
        
        # Save document
        output_filename = f"Placard_{shipment_num}.docx"
        output_path = os.path.join(self.output_folder, output_filename)
        
        try:
            if main_doc is not None:
                main_doc.save(output_path)
                print(f"SUCCESS: Created placard document: {output_path}")
                return True
            else:
                print("ERROR: No document was created")
                return False
        except Exception as e:
            print(f"ERROR saving document: {e}")
            return False
    
    def get_user_input(self) -> List[str]:
        """Get shipment numbers from user input"""
        while True:
            try:
                user_input = input("\nEnter one or more Shipment Numbers (comma-separated): ").strip()
                if not user_input:
                    print("Please enter at least one shipment number.")
                    continue
                
                # Split by comma and clean up
                shipment_numbers = [num.strip() for num in user_input.split(',')]
                shipment_numbers = [num for num in shipment_numbers if num]  # Remove empty strings
                
                if not shipment_numbers:
                    print("Please enter valid shipment numbers.")
                    continue
                
                return shipment_numbers
                
            except KeyboardInterrupt:
                print("\nOperation cancelled by user.")
                sys.exit(0)
            except Exception as e:
                print(f"Error reading input: {e}")
                continue
    
    def run(self) -> None:
        """Main execution method"""
        print("=== Shipping Placard Generator ===")
        print("Loading data and preparing system...")
        
        # Setup directories
        if not self.setup_directories():
            return
        
        # Load and prepare data (one-time operation)
        if not self.load_and_prepare_data():
            print("Failed to load data. Please check the Data folder and file format.")
            return
        
        # Check template exists
        template_path = os.path.join(self.template_folder, "placard_template.docx")
        if not os.path.exists(template_path):
            print(f"ERROR: Template file not found: {template_path}")
            print("Please place the template file 'placard_template.docx' in the Template folder.")
            return
        
        print("\nData loaded successfully! Ready to generate placards.")
        
        # Process user requests
        successful_count = 0
        failed_count = 0
        
        while True:
            try:
                # Get user input
                shipment_numbers = self.get_user_input()
                
                # Process each shipment
                for shipment_num in shipment_numbers:
                    if self.process_shipment(shipment_num):
                        successful_count += 1
                    else:
                        failed_count += 1
                
                # Print summary
                print(f"\n=== Processing Summary ===")
                print(f"Documents created: {successful_count}")
                print(f"Failed inputs: {failed_count}")
                
                # Ask if user wants to continue
                continue_choice = input("\nProcess more shipments? (y/n): ").strip().lower()
                if continue_choice not in ['y', 'yes']:
                    break
                    
            except KeyboardInterrupt:
                print("\n\nOperation cancelled by user.")
                break
            except Exception as e:
                print(f"Unexpected error: {e}")
                break
        
        print(f"\n=== Final Summary ===")
        print(f"Total documents created: {successful_count}")
        print(f"Total failed inputs: {failed_count}")
        print("Thank you for using the Shipping Placard Generator!")


def main() -> None:
    """Main entry point"""
    generator = PlacardGenerator()
    generator.run()


if __name__ == "__main__":
    main() 