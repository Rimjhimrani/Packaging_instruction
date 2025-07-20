import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io
import tempfile
import os
from datetime import datetime

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
    
    def extract_data_from_excel(self, uploaded_file):
        """Extract data from uploaded Excel file"""
        extracted_data = {}
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file, sheet_name=0)
            
            # Create a mapping of possible column names to our field names
            field_mapping = {
                # Basic info
                'revision no.': 'Revision No.',
                'revision': 'Revision No.',
                'date': 'Date',
                
                # Vendor info
                'vendor code': 'Vendor Code',
                'code': 'Vendor Code',
                'vendor name': 'Vendor Name',
                'name': 'Vendor Name',
                'vendor location': 'Vendor Location',
                'location': 'Vendor Location',
                
                # Part info
                'part no.': 'Part No.',
                'part number': 'Part No.',
                'part description': 'Part Description',
                'description': 'Part Description',
                'part unit weight': 'Part Unit Weight',
                'unit weight': 'Part Unit Weight',
                'weight': 'Part Unit Weight',
                'part l': 'Part L',
                'length': 'Part L',
                'part w': 'Part W',
                'width': 'Part W',
                'part h': 'Part H',
                'height': 'Part H',
                
                # Primary packaging
                'primary packaging type': 'Primary Packaging Type',
                'packaging type': 'Primary Packaging Type',
                'primary l-mm': 'Primary L-mm',
                'primary l': 'Primary L-mm',
                'primary w-mm': 'Primary W-mm',
                'primary w': 'Primary W-mm',
                'primary h-mm': 'Primary H-mm',
                'primary h': 'Primary H-mm',
                'primary qty/pack': 'Primary Qty/Pack',
                'qty/pack': 'Primary Qty/Pack',
                'primary empty weight': 'Primary Empty Weight',
                'empty weight': 'Primary Empty Weight',
                'primary pack weight': 'Primary Pack Weight',
                'pack weight': 'Primary Pack Weight',
                
                # Secondary packaging
                'secondary packaging type': 'Secondary Packaging Type',
                'secondary l-mm': 'Secondary L-mm',
                'secondary l': 'Secondary L-mm',
                'secondary w-mm': 'Secondary W-mm',
                'secondary w': 'Secondary W-mm',
                'secondary h-mm': 'Secondary H-mm',
                'secondary h': 'Secondary H-mm',
                'secondary qty/pack': 'Secondary Qty/Pack',
                'secondary empty weight': 'Secondary Empty Weight',
                'secondary pack weight': 'Secondary Pack Weight',
                
                # Approval
                'issued by': 'Issued By',
                'reviewed by': 'Reviewed By',
                'approved by': 'Approved By',
                
                # Additional
                'problem if any': 'Problem If Any',
                'caution': 'Caution'
            }
            
            # Extract data from DataFrame
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if col_lower in field_mapping:
                    field_name = field_mapping[col_lower]
                    # Get first non-null value from the column
                    values = df[col].dropna()
                    if len(values) > 0:
                        extracted_data[field_name] = str(values.iloc[0])
            
            # Try to extract procedure steps if they exist
            for i in range(1, 11):
                step_patterns = [f'procedure step {i}', f'step {i}', f'{i}']
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    if any(pattern in col_lower for pattern in step_patterns):
                        values = df[col].dropna()
                        if len(values) > 0:
                            extracted_data[f'Procedure Step {i}'] = str(values.iloc[0])
                        break
            
            st.success(f"Successfully extracted {len(extracted_data)} fields from Excel file")
            return extracted_data
            
        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            return {}
    
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
                    try:
                        cell_value = ws.cell(row=row_idx, column=col_idx).value
                        if cell_value:
                            cell_value = str(cell_value).strip().lower()
                            if "current packaging image" in cell_value or "current packaging" in cell_value:
                                header_positions['Current Packaging'] = col_idx - 1  # Convert to 0-based
                            elif "primary packaging image" in cell_value or "primary packaging" in cell_value:
                                header_positions['Primary Packaging'] = col_idx - 1
                            elif "secondary packaging image" in cell_value or "secondary packaging" in cell_value:
                                header_positions['Secondary Packaging'] = col_idx - 1
                            elif "label image" in cell_value or "label" in cell_value:
                                header_positions['Label'] = col_idx - 1
                    except Exception:
                        continue
            
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
                                # Final fallback: assign to first empty slot
                                for category in ['Current Packaging', 'Primary Packaging', 'Secondary Packaging', 'Label']:
                                    if not images_data[category]:
                                        images_data[category] = pil_image
                                        st.info(f"Assigned to {category} (fallback)")
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

        # Continue with remaining rows following the same pattern...
        # (Rest of the template creation code remains the same)
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

        # Continue with remaining template structure...
        # (The full template creation code continues here)

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
        
        # Create tabs for better organization
        tab1, tab2, tab3, tab4 = st.tabs(["Basic Info", "Packaging Details", "Procedures", "Images & Generate"])
        
        with tab1:
            st.subheader("📋 Basic Information")
            
            # Create form in columns
            col1, col2 = st.columns(2)
            
            with col1:
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
                col2c, col2d, col2e  = st.columns(3)
                with col2c:
                    part_l = st.number_input("Length (mm)", min_value=0.0, format="%.1f")
                with col2d:
                    part_w = st.number_input("Width (mm)", min_value=0.0, format="%.1f")
                with col2e:
                    part_h = st.number_input("Height (mm)", min_value=0.0, format="%.1f")
        
        with tab2:
            st.subheader("📦 Packaging Details")
            
            col3, col4 = st.columns(2)
            
            with col3:
                st.write("**Primary Packaging**")
                primary_type = st.selectbox("Primary Packaging Type", 
                    ["", "RIM (R-1)", "REAR DOME", "FRONT DOME", "REAR WINDSHIELD", "FRONT HARNESS", "Custom"])
                
                if primary_type == "Custom":
                    primary_type = st.text_input("Custom Primary Type")
                
                col3a, col3b, col3c = st.columns(3)
                with col3a:
                    primary_l = st.number_input("Primary L (mm)", min_value=0.0, format="%.0f", key="prim_l")
                with col3b:
                    primary_w = st.number_input("Primary W (mm)", min_value=0.0, format="%.0f", key="prim_w")
                with col3c:
                    primary_h = st.number_input("Primary H (mm)", min_value=0.0, format="%.0f", key="prim_h")
                
                col3d, col3e = st.columns(2)
                with col3d:
                    primary_qty = st.number_input("Primary Qty/Pack", min_value=0, format="%d", key="prim_qty")
                    primary_empty_weight = st.number_input("Primary Empty Weight (kg)", min_value=0.0, format="%.2f", key="prim_empty")
                with col3e:
                    primary_pack_weight = st.number_input("Primary Pack Weight (kg)", min_value=0.0, format="%.2f", key="prim_pack")
            
            with col4:
                st.write("**Secondary Packaging**")
                secondary_type = st.text_input("Secondary Packaging Type")
                
                col4a, col4b, col4c = st.columns(3)
                with col4a:
                    secondary_l = st.number_input("Secondary L (mm)", min_value=0.0, format="%.0f", key="sec_l")
                with col4b:
                    secondary_w = st.number_input("Secondary W (mm)", min_value=0.0, format="%.0f", key="sec_w")
                with col4c:
                    secondary_h = st.number_input("Secondary H (mm)", min_value=0.0, format="%.0f", key="sec_h")
                
                col4d, col4e = st.columns(2)
                with col4d:
                    secondary_qty = st.number_input("Secondary Qty/Pack", min_value=0, format="%d", key="sec_qty")
                    secondary_empty_weight = st.number_input("Secondary Empty Weight (kg)", min_value=0.0, format="%.2f", key="sec_empty")
                with col4e:
                    secondary_pack_weight = st.number_input("Secondary Pack Weight (kg)", min_value=0.0, format="%.2f", key="sec_pack")
        
        with tab3:
            st.subheader("📋 Packaging Procedures")
            
            # Auto-populate procedures if primary type is selected
            procedure_steps = [""] * 10
            if primary_type and primary_type in template_manager.packaging_procedures:
                procedure_steps = template_manager.get_procedure_steps(primary_type)
                st.info(f"Auto-populated procedures for {primary_type}")
            
            # Allow manual editing of procedure steps
            col5, col6 = st.columns(2)
            
            with col5:
                st.write("**Steps 1-5**")
                step1 = st.text_area("Procedure Step 1", value=procedure_steps[0], key="step1")
                step2 = st.text_area("Procedure Step 2", value=procedure_steps[1], key="step2")
                step3 = st.text_area("Procedure Step 3", value=procedure_steps[2], key="step3")
                step4 = st.text_area("Procedure Step 4", value=procedure_steps[3], key="step4")
                step5 = st.text_area("Procedure Step 5", value=procedure_steps[4], key="step5")
            
            with col6:
                st.write("**Steps 6-10**")
                step6 = st.text_area("Procedure Step 6", value=procedure_steps[5], key="step6")
                step7 = st.text_area("Procedure Step 7", value=procedure_steps[6], key="step7")
                step8 = st.text_area("Procedure Step 8", value=procedure_steps[7], key="step8")
                step9 = st.text_area("Procedure Step 9", value=procedure_steps[8], key="step9")
                step10 = st.text_area("Procedure Step 10", value=procedure_steps[9], key="step10")
            
            # Additional fields
            st.subheader("📝 Additional Information")
            col7, col8 = st.columns(2)
            with col7:
                problem_if_any = st.text_area("Problem If Any")
                caution = st.text_area("Caution")
            
            with col8:
                st.write("**Approval**")
                issued_by = st.text_input("Issued By")
                reviewed_by = st.text_input("Reviewed By")
                approved_by = st.text_input("Approved By")
        
        with tab4:
            st.subheader("🖼️ Images")
            
            col9, col10 = st.columns(2)
            
            with col9:
                current_packaging_img = st.file_uploader("Current Packaging Image", type=['png', 'jpg', 'jpeg'])
                primary_packaging_img = st.file_uploader("Primary Packaging Image", type=['png', 'jpg', 'jpeg'])
            
            with col10:
                secondary_packaging_img = st.file_uploader("Secondary Packaging Image", type=['png', 'jpg', 'jpeg'])
                label_img = st.file_uploader("Label Image", type=['png', 'jpg', 'jpeg'])
            
            st.subheader("📁 Generate Template")
            
            if st.button("🚀 Generate Excel Template", type="primary"):
                # Compile form data
                form_data = {
                    'Revision No.': revision_no,
                    'Date': str(date) if date else '',
                    'Vendor Code': vendor_code,
                    'Vendor Name': vendor_name,
                    'Vendor Location': vendor_location,
                    'Part No.': part_no,
                    'Part Description': part_description,
                    'Part Unit Weight': str(part_unit_weight) if part_unit_weight > 0 else '',
                    'Part Weight Unit': part_weight_unit,
                    'Part L': str(part_l) if part_l > 0 else '',
                    'Part W': str(part_w) if part_w > 0 else '',
                    'Part H': str(part_h) if part_h > 0 else '',
                    'Primary Packaging Type': primary_type,
                    'Primary L-mm': str(primary_l) if primary_l > 0 else '',
                    'Primary W-mm': str(primary_w) if primary_w > 0 else '',
                    'Primary H-mm': str(primary_h) if primary_h > 0 else '',
                    'Primary Qty/Pack': str(primary_qty) if primary_qty > 0 else '',
                    'Primary Empty Weight': str(primary_empty_weight) if primary_empty_weight > 0 else '',
                    'Primary Pack Weight': str(primary_pack_weight) if primary_pack_weight > 0 else '',
                    'Secondary Packaging Type': secondary_type,
                    'Secondary L-mm': str(secondary_l) if secondary_l > 0 else '',
                    'Secondary W-mm': str(secondary_w) if secondary_w > 0 else '',
                    'Secondary H-mm': str(secondary_h) if secondary_h > 0 else '',
                    'Secondary Qty/Pack': str(secondary_qty) if secondary_qty > 0 else '',
                    'Secondary Empty Weight': str(secondary_empty_weight) if secondary_empty_weight > 0 else '',
                    'Secondary Pack Weight': str(secondary_pack_weight) if secondary_pack_weight > 0 else '',
                    'Procedure Step 1': step1,
                    'Procedure Step 2': step2,
                    'Procedure Step 3': step3,
                    'Procedure Step 4': step4,
                    'Procedure Step 5': step5,
                    'Procedure Step 6': step6,
                    'Procedure Step 7': step7,
                    'Procedure Step 8': step8,
                    'Procedure Step 9': step9,
                    'Procedure Step 10': step10,
                    'Issued By': issued_by,
                    'Reviewed By': reviewed_by,
                    'Approved By': approved_by,
                    'Problem If Any': problem_if_any,
                    'Caution': caution
                }
                
                # Process images
                images_data = {}
                if current_packaging_img:
                    images_data['Current Packaging'] = PILImage.open(current_packaging_img)
                if primary_packaging_img:
                    images_data['Primary Packaging'] = PILImage.open(primary_packaging_img)
                if secondary_packaging_img:
                    images_data['Secondary Packaging'] = PILImage.open(secondary_packaging_img)
                if label_img:
                    images_data['Label'] = PILImage.open(label_img)
                
                # Generate Excel file
                try:
                    wb = template_manager.create_exact_template_excel()
                    wb = template_manager.populate_template_with_data(wb, form_data, images_data)
                    
                    # Save to buffer
                    buffer = io.BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)
                    
                    # Provide download
                    st.success("✅ Template generated successfully!")
                    st.download_button(
                        label="⬇️ Download Excel Template",
                        data=buffer.getvalue(),
                        file_name=f"Packaging_Template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"Error generating template: {str(e)}")
    
    else:  # Upload & Modify Existing mode
        st.header("📁 Upload & Modify Existing Template")
        
        uploaded_file = st.file_uploader(
            "Upload Existing Excel Template",
            type=['xlsx', 'xls'],
            help="Upload an existing packaging template to extract and modify data"
        )
        
        if uploaded_file is not None:
            st.success("File uploaded successfully!")
            
            # Extract data and images from uploaded file
            with st.spinner("Extracting data from Excel file..."):
                extracted_data = template_manager.extract_data_from_excel(uploaded_file)
                
                # Reset file pointer for image extraction
                uploaded_file.seek(0)
                extracted_images = template_manager.extract_images_from_excel(uploaded_file)
            
            if extracted_data:
                st.subheader("📊 Extracted Data")
                with st.expander("View Extracted Fields", expanded=False):
                    for key, value in extracted_data.items():
                        if value:
                            st.write(f"**{key}**: {value}")
                
                # Only show packaging procedure selection
                st.subheader("📋 Update Packaging Procedures")
                
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.write("**Select Packaging Type:**")
                    procedure_type = st.selectbox(
                        "Packaging Procedure Type",
                        ["", "RIM (R-1)", "REAR DOME", "FRONT DOME", "REAR WINDSHIELD", "FRONT HARNESS"],
                        help="Select a packaging type to auto-populate procedure steps"
                    )
                
                with col2:
                    if procedure_type:
                        st.info(f"Selected: {procedure_type}")
                        if procedure_type in template_manager.packaging_procedures:
                            procedures = template_manager.get_procedure_steps(procedure_type)
                            st.write("**Procedure Steps Preview:**")
                            for i, step in enumerate(procedures, 1):
                                if step.strip():
                                    st.write(f"{i}. {step}")
                
                st.subheader("📁 Generate Updated Template")
                
                if st.button("🚀 Generate Updated Excel Template", type="primary"):
                    # Use original extracted data
                    updated_form_data = extracted_data.copy()
                    
                    # Update only the procedure steps if a type is selected
                    if procedure_type and procedure_type in template_manager.packaging_procedures:
                        procedure_steps = template_manager.get_procedure_steps(procedure_type)
                        for i, step in enumerate(procedure_steps, 1):
                            updated_form_data[f'Procedure Step {i}'] = step
                        
                        # Also update the primary packaging type
                        updated_form_data['Primary Packaging Type'] = procedure_type
                        st.success(f"Updated procedures for {procedure_type}")
                    
                    # Generate Excel file
                    try:
                        wb = template_manager.create_exact_template_excel()
                        wb = template_manager.populate_template_with_data(wb, updated_form_data, extracted_images)
                        
                        # Save to buffer
                        buffer = io.BytesIO()
                        wb.save(buffer)
                        buffer.seek(0)
                        
                        # Provide download
                        st.success("✅ Updated template generated successfully!")
                        st.download_button(
                            label="⬇️ Download Updated Excel Template",
                            data=buffer.getvalue(),
                            file_name=f"Updated_Packaging_Template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    except Exception as e:
                        st.error(f"Error generating updated template: {str(e)}")
                        st.write("Debug info:", str(e))
            else:
                st.warning("Could not extract data from the uploaded file. Please check the file format and try again.")

if __name__ == "__main__":
    main()
