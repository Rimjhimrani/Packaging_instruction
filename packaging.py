import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
import tempfile
import io
import os
from PIL import Image as PILImage

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
        
        # Predefined packaging procedures for different types
        self.packaging_procedures = {
            "RIM (R-1)": [
                "Put 16 quantity of part on plastic pallet",
                "Apply pet strap over it and put Traceability label as per PMSPL standard guideline",
                "Stretch wrap complete pack",
                "Prepare additional carton boxes in line with procurement schedule (multiple of primary pack quantity -- 16)",
                "Load parts on base plastic pallet -- 4 parts per layer & max 4 level",
                "Apply traceability label on complete pack",
                "Attached packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only",
                "",
                ""
            ],
            "REAR DOME": [
                "Pick up one part and apply bubble wrapping over it",
                "Apply tape and put 12 bubble wrapped part into a Polypropylene box",
                "Put Traceability label as per PMSPL standard guideline",
                "Prepare additional Polypropylene box in line with procurement schedule (multiple of secondary pack quantity -- 12)",
                "If procurement schedule is for less no. of parts, then load similar other parts in same Polypropylene box",
                "Apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attached packing list along with dispatch document and tag copy of same on pack (in case of multiple parts in same polypropylene box)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only",
                ""
            ],
            "FRONT DOME": [
                "Pick up 1 part and apply bubble wrapping over it",
                "Apply tape and put 12 bubble wrapped part into a Polypropylene box",
                "Put Traceability label as per PMSPL standard guideline",
                "Prepare additional Polypropylene box in line with procurement schedule (multiple of secondary pack quantity -- 12)",
                "If procurement schedule is for less no. of parts, then load similar other parts in same Polypropylene box",
                "Apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attached packing list along with dispatch document and tag copy of same on pack (in case of multiple parts in same polypropylene box)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only",
                ""
            ],
            "REAR WINDSHIELD": [
                "Pick 20 quantity of rear windshield glass",
                "Pack rear windshield it in metallic pallet with rubber cushioning separators between parts to arrest part movement during handling",
                "Seal metallic pallet with top rubber cushioning separators",
                "Prepare additional pallets inline with scheduled requirement",
                "Apply traceability label on complete pack",
                "Attached packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only",
                "",
                "",
                ""
            ],
            "FRONT HARNESS": [
                "Pick up 5 quantity of part and put it in a polybag",
                "Put 2 such polybags in a polypropylene box",
                "Seal polypropylene box and put Traceability label as per PMSPL standard guideline",
                "Prepare additional polypropylene boxes in line with procurement schedule (multiple of primary pack quantity -- 5)",
                "If procurement schedule is for less no. of boxes, then load similar boxes of other parts on same plastic pallet",
                "Load polypropylene boxes on base plastic pallet -- 4 boxes per layer & max 2 level",
                "Apply pet strap (2 times -- cross way)",
                "Apply traceability label on complete pack",
                "Attached packing list along with dispatch document and tag copy of same on pack (in case of multiple parts on same pallet)",
                "Ensure Loading/Unloading of palletize load using Hand pallet / stacker / forklift only"
            ]
        }
    
    def get_procedure_steps(self, packaging_type):
        """Get predefined procedure steps for selected packaging type"""
        return self.packaging_procedures.get(packaging_type, [""] * 10)
    
    def extract_images_from_excel(self, uploaded_file):
        """Extract images from Excel file"""
        images_data = {
            'Current Packaging': None,
            'Primary Packaging': None,
            'Secondary Packaging': None,
            'Label': None
        }
        tmp_file_path = None
        try:
            # Save uploaded file to temporary location
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            # Load workbook and extract images
            wb = load_workbook(tmp_file_path)
            ws = wb.active
    
            # First, let's find the header row and column positions dynamically
            header_positions = {}
    
            # Search for headers in the worksheet (typically in first few rows)
            for row_idx in range(1, 10):  # Check first 10 rows for headers
                for col_idx in range(1, ws.max_column + 1):
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    if cell_value:
                        cell_value = str(cell_value).strip().lower()
                        if "current packaging image" in cell_value:
                            header_positions['Current Packaging'] = col_idx - 1  # Convert to 0-based
                        elif "primary packaging image" in cell_value:
                            header_positions['Primary Packaging'] = col_idx - 1
                        elif "secondary packaging image" in cell_value:
                            header_positions['Secondary Packaging'] = col_idx - 1
                        elif "label image" in cell_value:
                            header_positions['Label'] = col_idx - 1
            st.info(f"Found header positions: {header_positions}")
    
            # Debug: Print total number of images found
            total_images = len(ws._images) if hasattr(ws, '_images') and ws._images else 0
            st.info(f"Found {total_images} images in the worksheet")
    
            # Extract all images from the worksheet
            if hasattr(ws, '_images') and ws._images:
                for idx, img in enumerate(ws._images):
                    try:
                        # Convert image to PIL Image
                        image_stream = io.BytesIO(img._data())
                        pil_image = PILImage.open(image_stream)
                
                        # Get anchor information to determine image location
                        anchor = img.anchor
                        col_idx = None
                        row_idx = None
                
                        # Get column and row from anchor
                        if hasattr(anchor, '_from') and anchor._from:
                            col_idx = anchor._from.col
                            row_idx = anchor._from.row
                        elif hasattr(anchor, 'col') and hasattr(anchor, 'row'):
                            col_idx = anchor.col
                            row_idx = anchor.row
                    
                        if col_idx is not None:
                            st.write(f"Image {idx+1}: Located at column {col_idx}, row {row_idx}")
                            # Map image to correct category based on header positions
                            assigned = False
                    
                            for category, expected_col in header_positions.items():
                                # Allow some flexibility in column matching (±1 column)
                                if abs(col_idx - expected_col) <= 1:
                                    images_data[category] = pil_image
                                    st.success(f"✅ {category} image found at col {col_idx}, row {row_idx}")
                                    assigned = True
                                    break
                            if not assigned:
                                st.warning(f"⚠️ Image {idx+1} found at unexpected location: col {col_idx}, row {row_idx}")
                                # Fallback: Try to map based on approximate column ranges
                                if 64 <= col_idx <= 65:  # BM-BN range
                                    if not images_data['Current Packaging']:
                                        images_data['Current Packaging'] = pil_image
                                        st.info("Assigned to Current Packaging (fallback)")
                                    elif not images_data['Primary Packaging']:
                                        images_data['Primary Packaging'] = pil_image
                                        st.info("Assigned to Primary Packaging (fallback)")
                                elif 66 <= col_idx <= 67:  # BO-BP range
                                    if not images_data['Secondary Packaging']:
                                        images_data['Secondary Packaging'] = pil_image
                                        st.info("Assigned to Secondary Packaging (fallback)")
                                    elif not images_data['Label']:
                                        images_data['Label'] = pil_image
                                        st.info("Assigned to Label (fallback)")
                                else:
                                    # Final fallback: assign to first empty slot
                                    for category in ['Current Packaging', 'Primary Packaging', 'Secondary Packaging', 'Label']:
                                        if not images_data[category]:
                                            images_data[category] = pil_image
                                            st.info(f"Assigned to {category} (final fallback)")
                                            break
                        else:
                            st.warning(f"Could not determine location for image {idx+1}")
                            # Final fallback: assign to first empty slot
                            for category in ['Current Packaging', 'Primary Packaging', 'Secondary Packaging', 'Label']:
                                if not images_data[category]:
                                    images_data[category] = pil_image
                                    st.info(f"Assigned to {category} (final fallback)")
                                    break
                    except Exception as img_error:
                        st.error(f"Error processing image {idx+1}: {str(img_error)}")
                        continue
            else:
                st.warning("No images found in the worksheet. Make sure images are properly embedded in the Excel file.")
                # Show final results
                st.info("Final image assignments:")
                for category, image in images_data.items():
                    if image:
                        st.success(f"✅ {category}: Image assigned")
                    else:
                        st.warning(f"❌ {category}: No image found")
                return images_data
        except Exception as e:
            st.error(f"Could not extract images: {str(e)}")
            return images_data
        finally:
            # Clean up temporary file
            if tmp_file_path and os.path.exists(tmp_file_path):
                try:
                    os.unlink(tmp_file_path)
                except Exception as cleanup_error:
                    st.warning(f"Could not clean up temporary file: {str(cleanup_error)}")

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
    
    def add_image_to_cell_range(self, ws, pil_image, start_cell, end_cell):
        """Add PIL image to specified cell range in worksheet"""
        try:
            # Convert PIL image to bytes
            img_buffer = io.BytesIO()
            pil_image.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            # Create openpyxl Image
            img = Image(img_buffer)
            
            # Calculate cell dimensions (approximate)
            start_col_letter = start_cell[0]
            start_row = int(start_cell[1:])
            end_col_letter = end_cell[0]
            end_row = int(end_cell[1:])
            
            # Estimate size based on merged cell area (adjust as needed)
            img.width = 120  # Adjust based on your needs
            img.height = 80  # Adjust based on your needs
            
            # Add image to worksheet
            ws.add_image(img, start_cell)
            
            return True
            
        except Exception as e:
            st.warning(f"Could not add image to worksheet: {str(e)}")
            return False

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
        bold_font = Font(color="000000", bold=True, size=12)
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
        ws['A2'].font = bold_font

        ws.merge_cells('B2:E2')
        ws['B2'] = "Revision 1"
        ws['B2'].border = border
        self.apply_border_to_range(ws, 'B2', 'E2')

        # Date field
        ws['F2'] = "Date"
        ws['F2'].border = border
        ws['F2'].alignment = left_alignment
        ws['F2'].font = bold_font

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
        ws['A5'].font = bold_font
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
        ws['F5'].font = bold_font

        ws.merge_cells('G5:K5')
        ws['G5'] = ""
        self.apply_border_to_range(ws, 'G5', 'K5')

        ws['L5'] = ""
        ws['L5'].border = border

        # Vendor Name Row
        ws['A6'] = "Name"
        ws['A6'].font = bold_font
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
        ws['F6'].font = bold_font

        ws.merge_cells('G6:K6')
        ws['G6'] = ""
        self.apply_border_to_range(ws, 'G6', 'K6')

        ws['L6'] = ""
        ws['L6'].border = border

        # Vendor Location Row
        ws['A7'] = "Location"
        ws['A7'].font = bold_font
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
        ws['F7'].font = bold_font

        ws.merge_cells('G7:K7')
        ws['G7'] = ""
        self.apply_border_to_range(ws, 'G7', 'K7')

        ws['L7'] = ""
        ws['L7'].border = border

        # Additional row after Unit Weight (Row 8) for L, W, H
        ws['F8'] = "L"
        ws['F8'].border = border
        ws['F8'].alignment = left_alignment
        ws['F8'].font = bold_font

        ws['G8'] = ""
        ws['G8'].border = border

        ws['H8'] = "W"
        ws['H8'].border = border
        ws['H8'].alignment = center_alignment
        ws['H8'].font = bold_font

        ws['I8'] = ""
        ws['I8'].border = border

        ws['J8'] = "H"
        ws['J8'].border = border
        ws['J8'].alignment = center_alignment
        ws['J8'].font = bold_font

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
            ws[f'{col}10'].font = bold_font

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
        ws['L10'].font = bold_font
        ws['L10'].alignment = left_alignment

        # Secondary packaging headers
        for i, header in enumerate(headers):
            col = chr(ord('A') + i)
            ws[f'{col}15'] = header
            ws[f'{col}15'].border = border
            ws[f'{col}15'].alignment = center_alignment
            ws[f'{col}15'].font = bold_font

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
            ws[f'A{row}'].font = bold_font

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
        ws['A31'].font = bold_font
        self.apply_border_to_range(ws, 'A31', 'C31')

        ws.merge_cells('D31:G31')
        ws['D31'] = "Secondary Packaging"
        ws['D31'].alignment = center_alignment
        ws['D31'].font = bold_font
        self.apply_border_to_range(ws, 'D31', 'G31')

        ws.merge_cells('H31:J31')
        ws['H31'] = "Label"
        ws['H31'].alignment = center_alignment
        ws['H31'].font = bold_font
        self.apply_border_to_range(ws, 'H31', 'J31')

        ws['K31'] = ""
        ws['K31'].border = border
        ws['L31'] = ""
        ws['L31'].border = border

        # Image placeholder areas (rows 32-37)
        ws.merge_cells('A32:C37')
        ws['A32'] = "Primary\nPackaging"
        ws['A32'].alignment = center_alignment
        ws['A2'].font = bold_font
        self.apply_border_to_range(ws, 'A32', 'C37')

        ws.merge_cells('D32:G37')
        ws['D32'] = "Secondary\nPackaging"
        ws['D32'].alignment = center_alignment
        self.apply_border_to_range(ws, 'D32', 'G37')

        ws.merge_cells('H32:J37')
        ws['H32'] = "Label"
        ws['H32'].alignment = center_alignment
        self.apply_border_to_range(ws, 'H32', 'J37')

        # K and L columns for rows 32-37
        for row in range(32, 38):
            ws[f'K{row}'] = ""
            ws[f'K{row}'].border = border
            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        # Approval section
        ws.merge_cells('A38:K38')
        ws['A38'] = "Approval"
        ws['A38'].fill = blue_fill
        ws['A38'].font = white_font
        ws['A38'].alignment = center_alignment
        self.apply_border_to_range(ws, 'A38', 'K38')

        ws['L38'] = ""
        ws['L38'].border = border

        # Approval rows
        approval_labels = ["Issued By", "Reviewed By", "Approved By"]
        for i, label in enumerate(approval_labels):
            row = 39 + i
            ws[f'A{row}'] = label
            ws[f'A{row}'].border = border
            ws[f'A{row}'].alignment = left_alignment
            ws[f'A{row}'].font = bold_font

            # Merge cells for name and signature
            ws.merge_cells(f'B{row}:E{row}')
            ws[f'B{row}'] = "Name & Sign:"
            ws[f'B{row}'].border = border
            ws[f'B{row}'].alignment = left_alignment
            self.apply_border_to_range(ws, f'B{row}', f'E{row}')

            ws[f'F{row}'] = "Date:"
            ws[f'F{row}'].border = border
            ws[f'F{row}'].alignment = left_alignment
            ws[f'F{row}'].font = bold_font

            ws.merge_cells(f'G{row}:K{row}')
            ws[f'G{row}'] = ""
            self.apply_border_to_range(ws, f'G{row}', f'K{row}')

            ws[f'L{row}'] = ""
            ws[f'L{row}'].border = border

        return wb

    def populate_template_with_data(self, wb, form_data, images_data=None):
        """Populate the template with form data and images"""
        ws = wb.active
        
        # Populate form data
        if form_data.get('Vendor Code'):
            ws['B5'] = form_data['Vendor Code']
        if form_data.get('Vendor Name'):
            ws['B6'] = form_data['Vendor Name']
        if form_data.get('Vendor Location'):
            ws['B7'] = form_data['Vendor Location']
            
        # Part information
        if form_data.get('Part No.'):
            ws['G5'] = form_data['Part No.']
        if form_data.get('Part Description'):
            ws['G6'] = form_data['Part Description']
        if form_data.get('Part Unit Weight'):
            weight_unit = form_data.get('Part Weight Unit', '')
            ws['G7'] = f"{form_data['Part Unit Weight']} {weight_unit}"
            
        # Part dimensions
        if form_data.get('Part L'):
            ws['G8'] = form_data['Part L']
        if form_data.get('Part W'):
            ws['I8'] = form_data['Part W']
        if form_data.get('Part H'):
            ws['K8'] = form_data['Part H']
            
        # Primary packaging
        if form_data.get('Primary Packaging Type'):
            ws['A11'] = form_data['Primary Packaging Type']
        if form_data.get('Primary L-mm'):
            ws['B11'] = form_data['Primary L-mm']
        if form_data.get('Primary W-mm'):
            ws['C11'] = form_data['Primary W-mm']
        if form_data.get('Primary H-mm'):
            ws['D11'] = form_data['Primary H-mm']
        if form_data.get('Primary Qty/Pack'):
            ws['E11'] = form_data['Primary Qty/Pack']
        if form_data.get('Primary Empty Weight'):
            ws['F11'] = form_data['Primary Empty Weight']
        if form_data.get('Primary Pack Weight'):
            ws['G11'] = form_data['Primary Pack Weight']
            
        # Secondary packaging
        if form_data.get('Secondary Packaging Type'):
            ws['A16'] = form_data['Secondary Packaging Type']
        if form_data.get('Secondary L-mm'):
            ws['B16'] = form_data['Secondary L-mm']
        if form_data.get('Secondary W-mm'):
            ws['C16'] = form_data['Secondary W-mm']
        if form_data.get('Secondary H-mm'):
            ws['D16'] = form_data['Secondary H-mm']
        if form_data.get('Secondary Qty/Pack'):
            ws['E16'] = form_data['Secondary Qty/Pack']
        if form_data.get('Secondary Empty Weight'):
            ws['F16'] = form_data['Secondary Empty Weight']
        if form_data.get('Secondary Pack Weight'):
            ws['G16'] = form_data['Secondary Pack Weight']
            
        # Procedure steps
        for i in range(1, 11):
            step_key = f'Procedure Step {i}'
            if form_data.get(step_key):
                row = 19 + i
                ws[f'B{row}'] = form_data[step_key]
                
        # Approval fields
        if form_data.get('Issued By'):
            ws['B39'] = f"Name & Sign: {form_data['Issued By']}"
        if form_data.get('Reviewed By'):
            ws['B40'] = f"Name & Sign: {form_data['Reviewed By']}"
        if form_data.get('Approved By'):
            ws['B41'] = f"Name & Sign: {form_data['Approved By']}"
            
        # Add date if provided
        if form_data.get('Date'):
            ws['G2'] = form_data['Date']
            
        # Problem and Caution fields
        if form_data.get('Problem If Any'):
            ws['L12'] = form_data['Problem If Any']
        if form_data.get('Caution'):
            ws['L13'] = form_data['Caution']
            
        # Add images if provided
        if images_data:
            # Current Packaging image (L2:L8 area)
            if images_data.get('Current Packaging'):
                self.add_image_to_cell_range(ws, images_data['Current Packaging'], 'L2', 'L8')
                
            # Primary Packaging image (A32:C37)
            if images_data.get('Primary Packaging'):
                self.add_image_to_cell_range(ws, images_data['Primary Packaging'], 'A32', 'C37')
                
            # Secondary Packaging image (D32:G37)
            if images_data.get('Secondary Packaging'):
                self.add_image_to_cell_range(ws, images_data['Secondary Packaging'], 'D32', 'G37')
                
            # Label image (H32:J37)
            if images_data.get('Label'):
                self.add_image_to_cell_range(ws, images_data['Label'], 'H32', 'J37')
                
        return wb

def main():
    st.set_page_config(page_title="Exact Packaging Template Generator", layout="wide")
    st.title("🏭 Packaging Instruction Template Generator")
    st.markdown("Generate packaging instruction templates that match your exact specifications")
    
    # Initialize template manager
    template_manager = ExactPackagingTemplateManager()
    
    # Sidebar for navigation
    st.sidebar.title("Navigation")
    mode = st.sidebar.selectbox("Select Mode", ["Create New Template", "Upload & Modify Existing"])
    
    if mode == "Create New Template":
        st.header("📝 Create New Packaging Template")
        
        # Create form in columns
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📋 Basic Information")
            revision_no = st.text_input("Revision No.", value="Revision 1")
            date = st.date_input("Date")
            
            st.subheader("🏢 Vendor Information")
            vendor_code = st.text_input("Vendor Code")
            vendor_name = st.text_input("Vendor Name")
            vendor_location = st.text_input("Vendor Location")
            
        with col2:
            st.subheader("🔧 Part Information")
            part_no = st.text_input("Part No.")
            part_description = st.text_area("Part Description", height=100)
            
            col2a, col2b = st.columns(2)
            with col2a:
                part_unit_weight = st.number_input("Part Unit Weight", min_value=0.0, format="%.2f")
            with col2b:
                part_weight_unit = st.selectbox("Weight Unit", ["kg", "g", "lbs"])
                
            st.write("**Part Dimensions**")
            col2c, col2d, col2e = st.columns(3)
            with col2c:
                part_l = st.number_input("Length (L)", min_value=0.0, format="%.2f")
            with col2d:
                part_w = st.number_input("Width (W)", min_value=0.0, format="%.2f")
            with col2e:
                part_h = st.number_input("Height (H)", min_value=0.0, format="%.2f")
        
        # Packaging sections
        st.subheader("📦 Primary Packaging")
        col3, col4 = st.columns(2)
        
        with col3:
            primary_packaging_type = st.selectbox("Primary Packaging Type", 
                                                 ["", "RIM (R-1)", "REAR DOME", "FRONT DOME", "REAR WINDSHIELD", "FRONT HARNESS", "Custom"])
            if primary_packaging_type == "Custom":
                primary_packaging_type = st.text_input("Enter Custom Primary Packaging Type")
                
            col3a, col3b, col3c = st.columns(3)
            with col3a:
                primary_l = st.number_input("Primary L (mm)", min_value=0.0, format="%.2f", key="prim_l")
            with col3b:
                primary_w = st.number_input("Primary W (mm)", min_value=0.0, format="%.2f", key="prim_w")
            with col3c:
                primary_h = st.number_input("Primary H (mm)", min_value=0.0, format="%.2f", key="prim_h")
                
        with col4:
            col4a, col4b = st.columns(2)
            with col4a:
                primary_qty = st.number_input("Primary Qty/Pack", min_value=0, key="prim_qty")
                primary_empty_weight = st.number_input("Primary Empty Weight", min_value=0.0, format="%.2f", key="prim_empty")
            with col4b:
                primary_pack_weight = st.number_input("Primary Pack Weight", min_value=0.0, format="%.2f", key="prim_pack")
                
        st.subheader("📦 Secondary Packaging")
        col5, col6 = st.columns(2)
        
        with col5:
            secondary_packaging_type = st.text_input("Secondary Packaging Type")
            col5a, col5b, col5c = st.columns(3)
            with col5a:
                secondary_l = st.number_input("Secondary L (mm)", min_value=0.0, format="%.2f", key="sec_l")
            with col5b:
                secondary_w = st.number_input("Secondary W (mm)", min_value=0.0, format="%.2f", key="sec_w")
            with col5c:
                secondary_h = st.number_input("Secondary H (mm)", min_value=0.0, format="%.2f", key="sec_h")
                
        with col6:
            col6a, col6b = st.columns(2)
            with col6a:
                secondary_qty = st.number_input("Secondary Qty/Pack", min_value=0, key="sec_qty")
                secondary_empty_weight = st.number_input("Secondary Empty Weight", min_value=0.0, format="%.2f", key="sec_empty")
            with col6b:
                secondary_pack_weight = st.number_input("Secondary Pack Weight", min_value=0.0, format="%.2f", key="sec_pack")
        
        # Auto-populate procedure steps if predefined packaging type is selected
        procedure_steps = [""] * 10
        if primary_packaging_type and primary_packaging_type in template_manager.packaging_procedures:
            procedure_steps = template_manager.get_procedure_steps(primary_packaging_type)
            
        st.subheader("📋 Packaging Procedure Steps")
        st.info("Procedure steps are auto-populated based on the selected primary packaging type. You can modify them as needed.")
        
        # Create procedure steps in a more compact layout
        col7, col8 = st.columns(2)
        
        with col7:
            step_1 = st.text_area("Step 1", value=procedure_steps[0], height=60)
            step_2 = st.text_area("Step 2", value=procedure_steps[1], height=60)
            step_3 = st.text_area("Step 3", value=procedure_steps[2], height=60)
            step_4 = st.text_area("Step 4", value=procedure_steps[3], height=60)
            step_5 = st.text_area("Step 5", value=procedure_steps[4], height=60)
            
        with col8:
            step_6 = st.text_area("Step 6", value=procedure_steps[5], height=60)
            step_7 = st.text_area("Step 7", value=procedure_steps[6], height=60)
            step_8 = st.text_area("Step 8", value=procedure_steps[7], height=60)
            step_9 = st.text_area("Step 9", value=procedure_steps[8], height=60)
            step_10 = st.text_area("Step 10", value=procedure_steps[9], height=60)
        
        st.subheader("✅ Approval Information")
        col9, col10, col11 = st.columns(3)
        
        with col9:
            issued_by = st.text_input("Issued By")
        with col10:
            reviewed_by = st.text_input("Reviewed By")
        with col11:
            approved_by = st.text_input("Approved By")
            
        st.subheader("⚠️ Additional Information")
        col12, col13 = st.columns(2)
        with col12:
            problem_if_any = st.text_area("Problem If Any", height=80)
        with col13:
            caution = st.text_area("Caution", value="REVISED DESIGN", height=80)
        
        # Image upload section
        st.subheader("🖼️ Upload Images")
        col14, col15 = st.columns(2)
        
        with col14:
            current_packaging_img = st.file_uploader("Current Packaging Image", type=['png', 'jpg', 'jpeg'])
            primary_packaging_img = st.file_uploader("Primary Packaging Image", type=['png', 'jpg', 'jpeg'])
            
        with col15:
            secondary_packaging_img = st.file_uploader("Secondary Packaging Image", type=['png', 'jpg', 'jpeg'])
            label_img = st.file_uploader("Label Image", type=['png', 'jpg', 'jpeg'])
        
        # Generate template button
        if st.button("🚀 Generate Template", type="primary", use_container_width=True):
            # Prepare form data
            form_data = {
                'Revision No.': revision_no,
                'Date': str(date) if date else '',
                'Vendor Code': vendor_code,
                'Vendor Name': vendor_name,
                'Vendor Location': vendor_location,
                'Part No.': part_no,
                'Part Description': part_description,
                'Part Unit Weight': part_unit_weight,
                'Part Weight Unit': part_weight_unit,
                'Part L': part_l,
                'Part W': part_w,
                'Part H': part_h,
                'Primary Packaging Type': primary_packaging_type,
                'Primary L-mm': primary_l,
                'Primary W-mm': primary_w,
                'Primary H-mm': primary_h,
                'Primary Qty/Pack': primary_qty,
                'Primary Empty Weight': primary_empty_weight,
                'Primary Pack Weight': primary_pack_weight,
                'Secondary Packaging Type': secondary_packaging_type,
                'Secondary L-mm': secondary_l,
                'Secondary W-mm': secondary_w,
                'Secondary H-mm': secondary_h,
                'Secondary Qty/Pack': secondary_qty,
                'Secondary Empty Weight': secondary_empty_weight,
                'Secondary Pack Weight': secondary_pack_weight,
                'Procedure Step 1': step_1,
                'Procedure Step 2': step_2,
                'Procedure Step 3': step_3,
                'Procedure Step 4': step_4,
                'Procedure Step 5': step_5,
                'Procedure Step 6': step_6,
                'Procedure Step 7': step_7,
                'Procedure Step 8': step_8,
                'Procedure Step 9': step_9,
                'Procedure Step 10': step_10,
                'Issued By': issued_by,
                'Reviewed By': reviewed_by,
                'Approved By': approved_by,
                'Problem If Any': problem_if_any,
                'Caution': caution
            }
            
            # Prepare images
            images_data = {}
            if current_packaging_img:
                images_data['Current Packaging'] = PILImage.open(current_packaging_img)
            if primary_packaging_img:
                images_data['Primary Packaging'] = PILImage.open(primary_packaging_img)
            if secondary_packaging_img:
                images_data['Secondary Packaging'] = PILImage.open(secondary_packaging_img)
            if label_img:
                images_data['Label'] = PILImage.open(label_img)
            
            try:
                # Create and populate template
                wb = template_manager.create_exact_template_excel()
                wb = template_manager.populate_template_with_data(wb, form_data, images_data)
                
                # Save to bytes
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                # Download button
                filename = f"packaging_instruction_{part_no}_{str(date)}.xlsx".replace(" ", "_")
                st.download_button(
                    label="📥 Download Template",
                    data=output.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.success("✅ Template generated successfully!")
                
            except Exception as e:
                st.error(f"❌ Error generating template: {str(e)}")
    
    elif mode == "Upload & Modify Existing":
        st.header("📤 Upload & Modify Existing Template")
        
        uploaded_file = st.file_uploader("Upload Excel Template", type=['xlsx', 'xls'])
        
        if uploaded_file is not None:
            try:
                # Read the existing file
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                st.success("✅ File uploaded successfully!")
                
                # Try to extract images
                with st.spinner("🔍 Extracting images from Excel file..."):
                    images_data = template_manager.extract_images_from_excel(uploaded_file)
                
                # Display extracted images
                if any(images_data.values()):
                    st.subheader("🖼️ Extracted Images")
                    cols = st.columns(4)
                    
                    for idx, (category, image) in enumerate(images_data.items()):
                        with cols[idx % 4]:
                            if image:
                                st.image(image, caption=category, use_column_width=True)
                            else:
                                st.info(f"No {category} image found")
                
                # Allow modification of extracted data
                st.subheader("✏️ Modify Template Data")
                st.info("Modify the fields below and regenerate the template")
                
                # Here you can add form fields similar to the "Create New Template" section
                # but pre-populated with data extracted from the uploaded file
                
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")

if __name__ == "__main__":
    main()
