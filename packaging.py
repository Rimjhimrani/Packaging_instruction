import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.drawing.image import Image
from openpyxl.cell.cell import MergedCell
import io
import base64
from PIL import Image as PILImage
import zipfile
import os
import tempfile
import streamlit as st
from openpyxl.utils import get_column_letter

class ExactPackagingTemplateManager:
    def __init__(self):
        self.template_fields = {
            # Header Information
            'Revision No.': '',
            'Date': '',
            
            # Vendor Information
            'Vendor Code': '',
            'Vendor Name': '',
            'Vendor Location': '',
            
            # Part Information
            'Part No.': '',
            'Part Description': '',
            'Part Unit Weight': '',
            'Part Weight Unit': '',
            'Part L': '',
            'Part W': '',
            'Part H': '',
            
            # Primary Packaging
            'Primary Packaging Type': '',
            'Primary L-mm': '',
            'Primary W-mm': '',
            'Primary H-mm': '',
            'Primary Qty/Pack': '',
            'Primary Empty Weight': '',
            'Primary Pack Weight': '',
            
            # Secondary Packaging
            'Secondary Packaging Type': '',
            'Secondary L-mm': '',
            'Secondary W-mm': '',
            'Secondary H-mm': '',
            'Secondary Qty/Pack': '',
            'Secondary Empty Weight': '',
            'Secondary Pack Weight': '',
            
            # Packaging Procedures (10 steps)
            'Procedure Step 1': '',
            'Procedure Step 2': '',
            'Procedure Step 3': '',
            'Procedure Step 4': '',
            'Procedure Step 5': '',
            'Procedure Step 6': '',
            'Procedure Step 7': '',
            'Procedure Step 8': '',
            'Procedure Step 9': '',
            'Procedure Step 10': '',
            
            # Approval
            'Issued By': '',
            'Reviewed By': '',
            'Approved By': '',
            
            # Additional fields
            'Problem If Any': '',
            'Caution': ''
        }
    
    def apply_border_to_range(self, ws, start_cell, end_cell):
        """Apply borders to a range of cells"""
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Parse cell references
        start_col = ord(start_cell[0]) - ord('A')
        start_row = int(start_cell[1:])
        end_col = ord(end_cell[0]) - ord('A')
        end_row = int(end_cell[1:])
        
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col+1)
                cell.border = border
    
    def create_exact_template_excel(self):
        """Create the exact Excel template matching the image"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Packaging Instruction"
        
        # Define styles
        blue_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        light_blue_fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True, size=12)
        black_font = Font(color="000000", bold=True, size=14)
        regular_font = Font(color="000000", size=12)
        bold_font = Font(color="000000", bold=True, size=12)  # Added bold font for headers
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                       top=Side(style='thin'), bottom=Side(style='thin'))
        center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        title_font = Font(bold=True, size=12)
        header_font = Font(bold=True)
        
        # Set column widths to match the image exactly
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 12
        ws.column_dimensions['I'].width = 12
        ws.column_dimensions['J'].width = 12
        ws.column_dimensions['K'].width = 12
        ws.column_dimensions['L'].width = 24

        # Set row heights
        for row in range(1, 51):
            ws.row_dimensions[row].height = 16

        # Header Row - "Packaging Instruction"
        ws.merge_cells('A1:K1')
        ws['A1'] = "Packaging Instruction"
        ws['A1'].fill = blue_fill
        ws['A1'].font = white_font
        ws['A1'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A1', 'K1')

        # Current Packaging header (right side)
        ws['L1'] = "CURRENT PACKAGING"
        ws['L1'].fill = blue_fill
        ws['L1'].font = white_font
        ws['L1'].border = border
        ws['L1'].alignment = center_alignment

        # Revision information row
        ws['A2'] = "Revision No."
        ws['A2'].border = border
        ws['A2'].alignment = left_alignment
        ws['A2'].font = bold_font  # Made bold

        ws.merge_cells('B2:E2')
        ws['B2'] = "Revision 1"
        ws['B2'].border = border
        self.apply_border_to_range(ws, 'B2', 'E2')

        # Date field
        ws['F2'] = "Date"
        ws['F2'].border = border
        ws['F2'].alignment = left_alignment
        ws['F2'].font = bold_font  # Made bold

        # Merge cells for date value
        ws.merge_cells('G2:K2')
        ws['G2'] = ""
        ws['G2'].border = border
        self.apply_border_to_range(ws, 'G2', 'K2')

        ws['L2'] = ""
        ws['L2'].border = border

        # Row 3 - empty with borders
        ws.merge_cells('B3:E3')
        ws['B3'] = ""
        self.apply_border_to_range(ws, 'A3', 'L3')

        # Row 4 - Section headers
        ws.merge_cells('A4:D4')
        ws['A4'] = "Vendor Information"
        ws['A4'].font = title_font
        ws['A4'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A4', 'D4')

        ws['E4'] = ""
        ws['E4'].border = border

        ws.merge_cells('F4:I4')
        ws['F4'] = "Part Information"
        ws['F4'].font = title_font
        ws['F4'].alignment = center_alignment
        self.apply_border_to_range(ws, 'F4', 'I4')

        # Apply borders to remaining cells in row 4
        for col in ['J', 'K', 'L']:
            ws[f'{col}4'] = ""
            ws[f'{col}4'].border = border

        # Vendor Code Row
        ws['A5'] = "Code"
        ws['A5'].font = bold_font  # Made bold
        ws['A5'].alignment = left_alignment
        ws['A5'].border = border

        ws.merge_cells('B5:D5')
        ws['B5'] = ""
        self.apply_border_to_range(ws, 'B5', 'D5')

        ws['E5'] = ""
        ws['E5'].border = border
        
        # Part fields
        ws['F5'] = "Part No."
        ws['F5'].border = border
        ws['F5'].alignment = left_alignment
        ws['F5'].font = bold_font  # Made bold

        ws.merge_cells('G5:K5')
        ws['G5'] = ""
        self.apply_border_to_range(ws, 'G5', 'K5')

        ws['L5'] = ""
        ws['L5'].border = border

        # Vendor Name Row
        ws['A6'] = "Name"
        ws['A6'].font = bold_font  # Made bold
        ws['A6'].alignment = left_alignment
        ws['A6'].border = border

        ws.merge_cells('B6:D6')
        ws['B6'] = ""
        self.apply_border_to_range(ws, 'B6', 'D6')

        ws['E6'] = ""
        ws['E6'].border = border

        ws['F6'] = "Description"
        ws['F6'].border = border
        ws['F6'].alignment = left_alignment
        ws['F6'].font = bold_font  # Made bold

        ws.merge_cells('G6:K6')
        ws['G6'] = ""
        self.apply_border_to_range(ws, 'G6', 'K6')

        ws['L6'] = ""
        ws['L6'].border = border

        # Vendor Location Row
        ws['A7'] = "Location"
        ws['A7'].font = bold_font  # Made bold
        ws['A7'].alignment = left_alignment
        ws['A7'].border = border

        ws.merge_cells('B7:D7')
        ws['B7'] = ""
        self.apply_border_to_range(ws, 'B7', 'D7')

        ws['E7'] = ""
        ws['E7'].border = border

        ws['F7'] = "Unit Weight"
        ws['F7'].border = border
        ws['F7'].alignment = left_alignment
        ws['F7'].font = bold_font  # Made bold

        ws.merge_cells('G7:K7')
        ws['G7'] = ""
        self.apply_border_to_range(ws, 'G7', 'K7')

        ws['L7'] = ""
        ws['L7'].border = border

        # Additional row after Unit Weight (Row 8) for L, W, H
        ws['F8'] = "L"
        ws['F8'].border = border
        ws['F8'].alignment = left_alignment
        ws['F8'].font = bold_font  # Made bold

        ws['G8'] = ""
        ws['G8'].border = border

        ws['H8'] = "W"
        ws['H8'].border = border
        ws['H8'].alignment = center_alignment
        ws['H8'].font = bold_font  # Made bold

        ws['I8'] = ""
        ws['I8'].border = border

        ws['J8'] = "H"
        ws['J8'].border = border
        ws['J8'].alignment = center_alignment
        ws['J8'].font = bold_font  # Made bold

        ws['K8'] = ""
        ws['K8'].border = border

        # Empty cells for A-E and L in row 8
        for col in ['A', 'B', 'C', 'D', 'E', 'L']:
            ws[f'{col}8'] = ""
            ws[f'{col}8'].border = border

        # Title row for Primary Packaging
        ws.merge_cells('A9:K9')
        ws['A9'] = "Primary Packaging Instruction (Primary / Internal)"
        ws['A9'].fill = blue_fill
        ws['A9'].font = white_font
        ws['A9'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A9', 'K9')

        ws['L9'] = "CURRENT PACKAGING"
        ws['L9'].fill = blue_fill
        ws['L9'].font = white_font
        ws['L9'].border = border
        ws['L9'].alignment = center_alignment

        # Primary packaging headers
        headers = ["Packaging Type", "L-mm", "W-mm", "H-mm", "Qty/Pack", "Empty Weight", "Pack Weight"]
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}10'] = header
            ws[f'{col}10'].border = border
            ws[f'{col}10'].alignment = center_alignment
            ws[f'{col}10'].font = bold_font  # Made bold

        # Empty cells for remaining columns in row 10
        for col in ['H', 'I', 'J', 'K', 'L']:
            ws[f'{col}10'] = ""
            ws[f'{col}10'].border = border

        # Primary packaging data rows (11-13)
        for row in range(11, 14):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border

        # TOTAL row
        ws['D13'] = "TOTAL"
        ws['D13'].border = border
        ws['D13'].font = black_font
        ws['D13'].alignment = center_alignment

        # Secondary Packaging Instruction header
        ws.merge_cells('A14:J14')
        ws['A14'] = "Secondary Packaging Instruction (Outer / External)"
        ws['A14'].fill = blue_fill
        ws['A14'].font = white_font
        ws['A14'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A14', 'J14')

        ws['K14'] = ""
        ws['K14'].border = border

        ws['L10'] = "PROBLEM IF ANY:"
        ws['L10'].border = border
        ws['L10'].font = bold_font  # Made bold
        ws['L10'].alignment = left_alignment

        # Secondary packaging headers
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}15'] = header
            ws[f'{col}15'].border = border
            ws[f'{col}15'].alignment = center_alignment
            ws[f'{col}15'].font = bold_font  # Made bold

        # Empty cells for remaining columns in row 15
        for col in ['H', 'I', 'J', 'K']:
            ws[f'{col}15'] = ""
            ws[f'{col}15'].border = border

        ws['L11'] = "CAUTION: REVISED DESIGN"
        ws['L11'].fill = red_fill
        ws['L11'].font = white_font
        ws['L11'].border = border
        ws['L11'].alignment = center_alignment

        # Secondary packaging data rows (16-18)
        for row in range(16, 19):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border

        # TOTAL row for secondary
        ws['D18'] = "TOTAL"
        ws['D18'].border = border
        ws['D18'].font = black_font
        ws['D18'].alignment = center_alignment

        # Packaging Procedure section
        ws.merge_cells('A19:K19')
        ws['A19'] = "Packaging Procedure"
        ws['A19'].fill = blue_fill
        ws['A19'].font = white_font
        ws['A19'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A19', 'K19')

        ws['L19'] = ""
        ws['L19'].border = border

        # Packaging procedure steps (rows 20-29) - WITH MERGED CELLS
        for i in range(1, 11):
            row = 19 + i
            ws[f'A{row}'] = str(i)
            ws[f'A{row}'].border = border
            ws[f'A{row}'].alignment = center_alignment
            ws[f'A{row}'].font = bold_font  # Made bold

            # MERGE CELLS B to J for each procedure step
            ws.merge_cells(f'B{row}:J{row}')
            ws[f'B{row}'] = ""
            ws[f'B{row}'].alignment = left_alignment
            self.apply_border_to_range(ws, f'B{row}', f'J{row}')

            # K and L columns
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Reference Images/Pictures section
        ws.merge_cells('A30:K30')
        ws['A30'] = "Reference Images/Pictures"
        ws['A30'].fill = blue_fill
        ws['A30'].font = white_font
        ws['A30'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A30', 'K30')

        ws['L30'] = ""
        ws['L30'].border = border

        # Image section headers
        ws.merge_cells('A31:C31')
        ws['A31'] = "Primary Packaging"
        ws['A31'].alignment = center_alignment
        ws['A31'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'A31', 'C31')

        ws.merge_cells('D31:G31')
        ws['D31'] = "Secondary Packaging"
        ws['D31'].alignment = center_alignment
        ws['D31'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'D31', 'G31')

        ws.merge_cells('H31:J31')
        ws['H31'] = "Label"
        ws['H31'].alignment = center_alignment
        ws['H31'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'H31', 'J31')

        ws['K31'] = ""
        ws['K31'].border = border
        ws['L31'] = ""
        ws['L31'].border = border

        # Image placeholder areas (rows 32-37)
        ws.merge_cells('A32:C37')
        ws['A32'] = "Primary\nPackaging"
        ws['A32'].alignment = center_alignment
        ws['A32'].font = regular_font
        self.apply_border_to_range(ws, 'A32', 'C37')

        # Arrow 1
        ws['D35'] = "→"
        ws['D35'].border = border
        ws['D35'].alignment = center_alignment
        ws['D35'].font = Font(size=20, bold=True)

        # Secondary Packaging image area
        ws.merge_cells('E32:F37')
        ws['E32'] = "SECONDARY\nPACKAGING"
        ws['E32'].alignment = center_alignment
        ws['E32'].font = regular_font
        ws['E32'].fill = light_blue_fill
        self.apply_border_to_range(ws, 'E32', 'F37')

        # Arrow 2
        ws['G35'] = "→"
        ws['G35'].border = border
        ws['G35'].alignment = center_alignment
        ws['G35'].font = Font(size=20, bold=True)

        # Label image area
        ws.merge_cells('H32:K37')
        ws['H32'] = "LABEL"
        ws['H32'].alignment = center_alignment
        ws['H32'].font = regular_font
        self.apply_border_to_range(ws, 'H32', 'K37')

        # Add borders to remaining cells in image section
        for row in range(32, 38):
            for col in ['D', 'G', 'L']:
                if row != 35 or col != 'D':  # Skip D35 and G35 which have arrows
                    if row != 35 or col != 'G':
                        ws[f'{col}{row}'] = ""
                        ws[f'{col}{row}'].border = border

        # Approval Section
        ws.merge_cells('A38:C38')
        ws['A38'] = "Issued By"
        ws['A38'].alignment = center_alignment
        ws['A38'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'A38', 'C38')

        ws.merge_cells('D38:G38')
        ws['D38'] = "Reviewed By"
        ws['D38'].alignment = center_alignment
        ws['D38'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'D38', 'G38')

        ws.merge_cells('H38:K38')
        ws['H38'] = "Approved By"
        ws['H38'].alignment = center_alignment
        ws['H38'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'H38', 'K38')

        ws['L38'] = ""
        ws['L38'].border = border

        # Signature boxes (rows 39-42)
        ws.merge_cells('A39:C42')
        ws['A39'] = ""
        self.apply_border_to_range(ws, 'A39', 'C42')

        ws.merge_cells('D39:G42')
        ws['D39'] = ""
        self.apply_border_to_range(ws, 'D39', 'G42')

        ws.merge_cells('H39:K42')
        ws['H39'] = ""
        self.apply_border_to_range(ws, 'H39', 'K42')

        # Apply borders for L column in signature section
        for row in range(39, 43):
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Second Approval Section
        ws.merge_cells('A43:C43')
        ws['A43'] = "Issued By"
        ws['A43'].alignment = center_alignment
        ws['A43'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'A43', 'C43')

        ws.merge_cells('D43:G43')
        ws['D43'] = "Reviewed By"
        ws['D43'].alignment = center_alignment
        ws['D43'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'D43', 'G43')

        ws.merge_cells('H43:J43')
        ws['H43'] = "Approved By"
        ws['H43'].alignment = center_alignment
        ws['H43'].font = bold_font  # Made bold
        self.apply_border_to_range(ws, 'H43', 'J43')

        ws['K43'] = ""
        ws['K43'].border = border
        ws['L43'] = ""
        ws['L43'].border = border

        # Second signature boxes (rows 44-47)
        ws.merge_cells('A44:C47')
        ws['A44'] = ""
        self.apply_border_to_range(ws, 'A44', 'C47')

        ws.merge_cells('D44:G47')
        ws['D44'] = ""
        self.apply_border_to_range(ws, 'D44', 'G47')

        ws.merge_cells('H44:J47')
        ws['H44'] = ""
        self.apply_border_to_range(ws, 'H44', 'J47')

        # Apply borders for K and L columns in second signature section
        for row in range(44, 48):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Final rows (48-50) - empty with borders
        for row in range(48, 51):
            for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
                ws[f'{col}{row}'] = ""
                ws[f'{col}{row}'].border = border

        # Return the workbook
        return wb
    
    def extract_data_from_uploaded_file(self, uploaded_file):
        """Extract comprehensive data from uploaded CSV/Excel file"""
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            # Convert dataframe to dictionary - extract all available data
            data_dict = {}
            
            # Map common variations of column names
            column_mappings = {
                # Header Information
                'revision no': 'Revision No.',
                'revision_no': 'Revision No.',
                'revision number': 'Revision No.',
                'date': 'Date',
                
                # Vendor Information
                'vendor code': 'Vendor Code',
                'vendor_code': 'Vendor Code',
                'supplier code': 'Vendor Code',
                'vendor name': 'Vendor Name',
                'vendor_name': 'Vendor Name',
                'supplier name': 'Vendor Name',
                'vendor location': 'Vendor Location',
                'vendor_location': 'Vendor Location',
                'supplier location': 'Vendor Location',
                
                # Part Information
                'part no': 'Part No.',
                'part_no': 'Part No.',
                'part number': 'Part No.',
                'part description': 'Part Description',
                'part_description': 'Part Description',
                'description': 'Part Description',
                'part unit weight': 'Part Unit Weight',
                'part_unit_weight': 'Part Unit Weight',
                'unit weight': 'Part Unit Weight',
                'weight': 'Part Unit Weight',
                'part weight unit': 'Part Weight Unit',
                'weight unit': 'Part Weight Unit',
                'part l': 'Part L',
                'part_l': 'Part L',
                'part length': 'Part L',
                'part w': 'Part W',
                'part_w': 'Part W',
                'part width': 'Part W',
                'part h': 'Part H',
                'part_h': 'Part H',
                'part height': 'Part H',
                
                # Primary Packaging
                'primary packaging type': 'Primary Packaging Type',
                'primary_packaging_type': 'Primary Packaging Type',
                'primary type': 'Primary Packaging Type',
                'primary l-mm': 'Primary L-mm',
                'primary_l_mm': 'Primary L-mm',
                'primary length': 'Primary L-mm',
                'primary w-mm': 'Primary W-mm',
                'primary_w_mm': 'Primary W-mm',
                'primary width': 'Primary W-mm',
                'primary h-mm': 'Primary H-mm',
                'primary_h_mm': 'Primary H-mm',
                'primary height': 'Primary H-mm',
                'primary qty/pack': 'Primary Qty/Pack',
                'primary_qty_pack': 'Primary Qty/Pack',
                'primary quantity': 'Primary Qty/Pack',
                'primary empty weight': 'Primary Empty Weight',
                'primary_empty_weight': 'Primary Empty Weight',
                'primary pack weight': 'Primary Pack Weight',
                'primary_pack_weight': 'Primary Pack Weight',
                
                # Secondary Packaging
                'secondary packaging type': 'Secondary Packaging Type',
                'secondary_packaging_type': 'Secondary Packaging Type',
                'secondary type': 'Secondary Packaging Type',
                'secondary l-mm': 'Secondary L-mm',
                'secondary_l_mm': 'Secondary L-mm',
                'secondary length': 'Secondary L-mm',
                'secondary w-mm': 'Secondary W-mm',
                'secondary_w_mm': 'Secondary W-mm',
                'secondary width': 'Secondary W-mm',
                'secondary h-mm': 'Secondary H-mm',
                'secondary_h_mm': 'Secondary H-mm',
                'secondary height': 'Secondary H-mm',
                'secondary qty/pack': 'Secondary Qty/Pack',
                'secondary_qty_pack': 'Secondary Qty/Pack',
                'secondary quantity': 'Secondary Qty/Pack',
                'secondary empty weight': 'Secondary Empty Weight',
                'secondary_empty_weight': 'Secondary Empty Weight',
                'secondary pack weight': 'Secondary Pack Weight',
                'secondary_pack_weight': 'Secondary Pack Weight',
                
                # Approval
                'issued by': 'Issued By',
                'issued_by': 'Issued By',
                'reviewed by': 'Reviewed By',
                'reviewed_by': 'Reviewed By',
                'approved by': 'Approved By',
                'approved_by': 'Approved By',
                
                # Additional
                'problem if any': 'Problem If Any',
                'problem_if_any': 'Problem If Any',
                'problems': 'Problem If Any',
                'caution': 'Caution',
                'warning': 'Caution'
            }
            
            # Add procedure steps mapping
            for i in range(1, 11):
                column_mappings[f'procedure step {i}'] = f'Procedure Step {i}'
                column_mappings[f'procedure_step_{i}'] = f'Procedure Step {i}'
                column_mappings[f'step {i}'] = f'Procedure Step {i}'
                column_mappings[f'step_{i}'] = f'Procedure Step {i}'
            
            # Extract data using mappings
            for col in df.columns:
                col_clean = col.strip().lower()
                mapped_field = column_mappings.get(col_clean, col.strip())
                
                # Extract first non-null value from the column
                if not df.empty and col in df.columns:
                    for idx in range(len(df)):
                        value = df[col].iloc[idx]
                        if pd.notna(value) and str(value).strip():
                            data_dict[mapped_field] = str(value).strip()
                            break
            
            return data_dict
            
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
            return {}
    
    def populate_template_with_data(self, wb, data_dict):
        """Populate the Excel template with extracted data"""
        ws = wb.active
        
        # Define cell mappings for data population
        cell_mappings = {
            # Header Information
            'Date': 'G2',
            
            # Vendor Information
            'Vendor Code': 'B5',
            'Vendor Name': 'B6',
            'Vendor Location': 'B7',
            
            # Part Information
            'Part No.': 'G5',
            'Part Description': 'G6',
            'Part Unit Weight': 'G7',
            'Part L': 'G8',
            'Part W': 'I8',
            'Part H': 'K8',
            
            # Primary Packaging
            'Primary Packaging Type': 'A11',
            'Primary L-mm': 'B11',
            'Primary W-mm': 'C11',
            'Primary H-mm': 'D11',
            'Primary Qty/Pack': 'E11',
            'Primary Empty Weight': 'F11',
            'Primary Pack Weight': 'G11',
            
            # Secondary Packaging
            'Secondary Packaging Type': 'A16',
            'Secondary L-mm': 'B16',
            'Secondary W-mm': 'C16',
            'Secondary H-mm': 'D16',
            'Secondary Qty/Pack': 'E16',
            'Secondary Empty Weight': 'F16',
            'Secondary Pack Weight': 'G16',
            
            # Approval
            'Issued By': 'A39',
            'Reviewed By': 'D39',
            'Approved By': 'H39',
            
            # Additional fields
            'Problem If Any': 'L12',
            'Caution': 'L13'
        }
        
        # Populate procedure steps
        for i in range(1, 11):
            cell_mappings[f'Procedure Step {i}'] = f'B{19+i}'
        
        # Apply data to cells
        for field, cell_ref in cell_mappings.items():
            if field in data_dict and data_dict[field]:
                try:
                    ws[cell_ref] = data_dict[field]
                except Exception as e:
                    st.warning(f"Could not populate {field}: {str(e)}")
        
        return wb
    
    def create_filled_template(self, uploaded_file=None):
        """Create template and optionally fill with uploaded data"""
        # Create the base template
        wb = self.create_exact_template_excel()
        
        # If file is uploaded, extract data and populate template
        if uploaded_file is not None:
            data_dict = self.extract_data_from_uploaded_file(uploaded_file)
            if data_dict:
                wb = self.populate_template_with_data(wb, data_dict)
                st.success(f"Template populated with data from {uploaded_file.name}")
            else:
                st.warning("No valid data found in uploaded file")
        
        return wb
    
    def save_workbook_to_bytes(self, wb):
        """Save workbook to bytes for download"""
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

# Streamlit App
def main():
    st.set_page_config(page_title="Packaging Template Generator", layout="wide")
    
    st.title("📦 Exact Packaging Instruction Template Generator")
    st.markdown("---")
    
    # Initialize template manager
    template_manager = ExactPackagingTemplateManager()
    
    # Sidebar for options
    st.sidebar.header("Options")
    
    # File upload option
    st.sidebar.subheader("Upload Data File (Optional)")
    uploaded_file = st.sidebar.file_uploader(
        "Choose CSV or Excel file to auto-populate template",
        type=['csv', 'xlsx', 'xls'],
        help="Upload a file with packaging data to automatically fill the template"
    )
    
    # Display file info if uploaded
    if uploaded_file:
        st.sidebar.success(f"✅ File uploaded: {uploaded_file.name}")
        
        # Show preview of uploaded data
        try:
            if uploaded_file.name.endswith('.csv'):
                preview_df = pd.read_csv(uploaded_file)
            else:
                preview_df = pd.read_excel(uploaded_file)
            
            st.sidebar.subheader("Data Preview")
            st.sidebar.dataframe(preview_df.head(3))
            
        except Exception as e:
            st.sidebar.error(f"Error previewing file: {str(e)}")
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("Generate Packaging Template")
        st.write("""
        This tool generates an exact replica of the packaging instruction template 
        with proper formatting, borders, and layout matching the original design.
        """)
        
        if uploaded_file:
            st.info("📋 Template will be auto-populated with data from your uploaded file")
        else:
            st.info("📋 Generate blank template (you can upload a data file to auto-populate)")
    
    with col2:
        st.header("Template Features")
        st.markdown("""
        ✅ **Exact Layout Match**
        - Proper cell merging
        - Accurate borders
        - Correct colors & fonts
        
        ✅ **Auto-Population**
        - CSV/Excel data import
        - Smart column mapping
        - Error handling
        
        ✅ **Professional Format**
        - Print-ready layout
        - Standard Excel format
        - Easy to edit
        """)
    
    # Generate button
    st.markdown("---")
    
    if st.button("🚀 Generate Packaging Template", type="primary", use_container_width=True):
        with st.spinner("Generating your packaging template..."):
            try:
                # Create the template
                wb = template_manager.create_filled_template(uploaded_file)
                
                # Convert to bytes for download
                excel_bytes = template_manager.save_workbook_to_bytes(wb)
                
                # Success message
                st.success("✅ Template generated successfully!")
                
                # Download button
                filename = f"Packaging_Instruction_Template_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                st.download_button(
                    label="📥 Download Excel Template",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Additional info
                st.info(f"💾 File saved as: {filename}")
                
            except Exception as e:
                st.error(f"❌ Error generating template: {str(e)}")
                st.error("Please check your data file format and try again.")

# Instructions section
    st.markdown("---")
    st.header("📖 How to Use")
    
    with st.expander("Step-by-Step Instructions"):
        st.markdown("""
        ### Option 1: Generate Blank Template
        1. Click "Generate Packaging Template" button
        2. Download the generated Excel file
        3. Fill in your data manually
        
        ### Option 2: Auto-Populate with Data
        1. Prepare your data in CSV or Excel format
        2. Upload your file using the sidebar
        3. Click "Generate Packaging Template"
        4. Download the pre-filled template
        
        ### Data File Format
        Your CSV/Excel file should have columns like:
        - `Vendor Code`, `Vendor Name`, `Vendor Location`
        - `Part No.`, `Part Description`, `Part Unit Weight`
        - `Primary Packaging Type`, `Primary L-mm`, etc.
        - `Procedure Step 1`, `Procedure Step 2`, etc.
        
        ### Supported File Types
        - CSV (.csv)
        - Excel (.xlsx, .xls)
        """)

if __name__ == "__main__":
    main()
