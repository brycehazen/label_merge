import os
import pandas as pd
from chardet.universaldetector import UniversalDetector
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.graphics.barcode import code39
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.units import inch
from conscode_mapping import get_parish_name
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfdoc import PDFCatalog, PDFDictionary, PDFString
from reportlab.pdfbase import pdfdoc
from reportlab.lib.colors import Color

# Required columns
required_columns = [
    'ConsID',
    'PrimAddText',
    'AddrLines',
    'AddrCity',
    'AddrState',
    'AddrZIP',
    'Year_2024'
]

def create_cover_sheet(c, cons_code, record_count):
    """Create editable cover sheet as first page"""
    # Get page size
    PAGE_WIDTH, PAGE_HEIGHT = LETTER
    
    # Calculate center of page
    center_x = PAGE_WIDTH / 2
    
    def center_text(text, y_position, font_name="Times-Roman", font_size=32):
        """Helper function to center text and adjust font size if needed"""
        c.setFont(font_name, font_size)
        text_width = stringWidth(text, font_name, font_size)
        
        # If text is too wide, reduce font size until it fits
        while text_width > (PAGE_WIDTH - 2*inch) and font_size > 16:
            font_size -= 2
            c.setFont(font_name, font_size)
            text_width = stringWidth(text, font_name, font_size)
        
        # Calculate x position to center text
        x_position = center_x - (text_width / 2)
        c.drawString(x_position, y_position, text)
        
        return font_size  # Return the final font size used
    
    # Add the logo image
    try:
        image_width = PAGE_WIDTH - inch  # Increased width (reduced margin from 2*inch to 1*inch)
        image_height = 3*inch  # Increased height from 2*inch to 3*inch
        image_x = center_x - (image_width / 2)
        image_y = PAGE_HEIGHT - 4*inch  # Adjusted y position to accommodate larger height
        
        c.drawImage('Untitled.png',
                    image_x,
                    image_y,
                    width=image_width,
                    height=image_height,
                    preserveAspectRatio=True,
                    mask='auto')
    except:
        print("Warning: Could not load logo image 'Untitled.png'")

    # Get parish name
    parish_name = get_parish_name(cons_code)

    # Add title text
    center_text("2025 Labels", PAGE_HEIGHT - 4*inch)
    
    # Add parish name (was parish code)
    used_font_size = center_text(f"{cons_code} {parish_name}", PAGE_HEIGHT - 5*inch)
    
    # Add record count (in smaller font)
    center_text(f"{record_count:,} Records", PAGE_HEIGHT - 7*inch, font_size=26)
    
    c.showPage()
    
def detect_file_encoding(file_path):
    """Detect the encoding of a file"""
    detector = UniversalDetector()
    try:
        with open(file_path, 'rb') as file:
            for line in file:
                detector.feed(line)
                if detector.done:
                    break
        detector.close()
        encoding = detector.result['encoding']
        return encoding or 'utf-8'
    except Exception as e:
        print(f"Error detecting file encoding for {file_path}: {e}")
        raise

def clean_text_field(value):
    """Clean text fields by converting null-like values to empty strings"""
    if pd.isna(value) or str(value).strip().lower() in ['nan', 'null', '']:
        return ''
    return str(value).strip()

def clean_currency(amount):
    """Format currency values consistently"""
    if pd.isna(amount) or str(amount).strip() == '':
        return '$0.00'
    # If it's numeric, just format
    if isinstance(amount, (int, float)):
        return f'${float(amount):,.2f}'
    # If it's a string that might contain $, commas, etc.
    val = str(amount).replace('$','').replace(',','')
    try:
        return f'${float(val):,.2f}'
    except:
        return '$0.00'

def validate_record(record):
    """
    Validate if a record should be included in the output
    Returns: bool indicating if record should be included
    """
    # Check if ConsID contains asterisk
    if '*' in str(record['ConsID']):
        return False
        
    # Check if AddrLines is empty/null/nan
    addr_lines = clean_text_field(record['AddrLines'])
    if not addr_lines:
        return False
        
    return True

def read_csv_data(csv_path):
    """Read and process CSV data"""
    try:
        encoding = detect_file_encoding(csv_path)
        df = pd.read_csv(
            csv_path,
            encoding=encoding,
            usecols=required_columns
        )
        # Validate required columns
        missing_cols = set(required_columns) - set(df.columns)
        if missing_cols:
            raise ValueError(f"Missing required columns in {csv_path}: {missing_cols}")
            
        # Filter out invalid records
        df = df[df.apply(validate_record, axis=1)]
        
        # Clean currency values for Year_2024
        df['Year_2024'] = df['Year_2024'].apply(clean_currency)
        
        return df
    except Exception as e:
        print(f"Error reading CSV file {csv_path}: {e}")
        raise

def wrap_text_to_width(text, font_name, font_size, max_width, canvas_obj):
    """
    Wrap text to fit within a specified width
    Returns list of wrapped lines
    """
    words = text.split()
    lines = []
    current_line = ""
    
    for w in words:
        # Try adding the next word
        test_line = (current_line + " " + w).strip() if current_line else w
        
        # If adding this word exceeds width, start a new line
        if stringWidth(test_line, font_name, font_size) > max_width:
            if current_line:
                lines.append(current_line)
            current_line = w
        else:
            current_line = test_line
            
    if current_line:
        lines.append(current_line)
    
    return lines

def fit_text_into_label(c, text_lines, label_width, label_height, font_name, start_font_size=14, min_font_size=12):
    """
    Wrap text to fit within label space, adjusting font size if needed
    Returns wrapped lines and final font size
    """
    label_width = label_width - 4  # Account for margins
    current_font_size = start_font_size
    
    # First try wrapping at current font size
    wrapped_lines = []
    for line in text_lines:
        wrapped_lines.extend(wrap_text_to_width(line, font_name, current_font_size, label_width, c))
    
    # Calculate if wrapped lines fit in height
    line_height = current_font_size * 1.2
    total_height = len(wrapped_lines) * line_height
    
    # If doesn't fit height, reduce font size until it does
    while total_height > label_height and current_font_size > min_font_size:
        current_font_size -= 1
        line_height = current_font_size * 1.2
        total_height = len(wrapped_lines) * line_height
        
        # Re-wrap at new font size if needed
        if total_height > label_height:
            wrapped_lines = []
            for line in text_lines:
                wrapped_lines.extend(wrap_text_to_width(line, font_name, current_font_size, label_width, c))
    
    return wrapped_lines, current_font_size

def create_labels_pdf(dataframe, output_path):
    """Create PDF with labels from DataFrame"""
    # PDF page size and margins
    PAGE_WIDTH, PAGE_HEIGHT = LETTER
    top_margin = 0.4 * inch
    bottom_margin = 0.5 * inch
    
    # Label dimensions
    label_width = 2.625 * inch
    label_height = 1.0 * inch
    
    # 3 columns, 10 rows = 30 labels per sheet
    left_margin = 0.3 * inch
    col_positions = [left_margin, left_margin + label_width, left_margin + 2*label_width]
    
    # Vertical positions (10 rows)
    start_y = PAGE_HEIGHT - top_margin - label_height
    row_positions = [start_y - r * label_height for r in range(10)]
    
    # Create PDF with editing enabled
    c = canvas.Canvas(output_path, pagesize=LETTER)
    c.setTitle("Editable Labels")
    c.setKeywords('Tagged PDF')
    c.setAuthor("Label Generator")
    PDFCatalog.OpenAction = None

   # Get ConsCode and create cover sheet
    filename = os.path.basename(output_path)
    cons_code = filename.split(' ')[0]  # Get code before first space
    record_count = len(dataframe)
    create_cover_sheet(c, cons_code, record_count)

    font_name = "Times-Roman"
    records = dataframe.to_dict('records')
    labels_per_page = 10
    row_count = 0

    for idx, rec in enumerate(records):
        # New page check
        if idx > 0 and idx % labels_per_page == 0:
            c.showPage()
            row_count = 0

        # Extract and clean data
        ConsID = clean_text_field(rec['ConsID'])
        PrimAddText = clean_text_field(rec['PrimAddText'])
        AddrLines = clean_text_field(rec['AddrLines'])
        AddrCity = clean_text_field(rec['AddrCity'])
        AddrState = clean_text_field(rec['AddrState'])
        AddrZIP = clean_text_field(rec['AddrZIP'])
        Year_2024 = str(rec['Year_2024']).strip()

        # Format address
        city_state_zip = f"{AddrCity}, {AddrState} {AddrZIP}" if AddrCity and AddrState else ""

        # Build lines list
        common_lines = [line for line in [PrimAddText, AddrLines, city_state_zip] if line]
        other_label_lines = [ConsID] + common_lines if ConsID else common_lines
        y_pos = row_positions[row_count % 10]

        for col_idx in range(3):
            x_pos = col_positions[col_idx]
            
            # Draw Barcode
            if ConsID:
                barcode = code39.Standard39(ConsID, barHeight=12, quiet=False, checksum=0)
                barcode.drawOn(c, x_pos, y_pos + label_height-22)

            text_area_height = label_height - (label_height*0.3) - 10

            # Handle first column differently
            if col_idx == 0:
                fitted_lines, fs = fit_text_into_label(
                    c,
                    [ConsID] + common_lines if ConsID else common_lines,
                    label_width - 10,
                    text_area_height,
                    font_name
                )
                c.setFont(font_name, fs)
                line_height = fs * .9
                current_y = y_pos + text_area_height
                
                # Draw ConsID and Year_2024 on same line
                if ConsID:
                    consid_width = stringWidth(ConsID, font_name, fs)
                    c.drawString(x_pos+2, current_y, ConsID)
                    c.drawString(x_pos+2+consid_width+30, current_y, Year_2024)
                    current_y -= line_height
                
                # Draw remaining lines
                for line in fitted_lines[1 if ConsID else 0:]:
                    c.drawString(x_pos+2, current_y, line)
                    current_y -= line_height
            else:
                # Other columns
                fitted_lines, fs = fit_text_into_label(
                    c,
                    other_label_lines,
                    label_width - 10,
                    text_area_height,
                    font_name
                )
                c.setFont(font_name, fs)
                line_height = fs * .9
                current_y = y_pos + text_area_height
                for line in fitted_lines:
                    c.drawString(x_pos+2, current_y, line)
                    current_y -= line_height

        row_count += 1

    c.showPage()
    c.save()

def process_csv_files():
    """Process all CSV files in current directory"""
    current_dir = os.getcwd()
    csv_files = [f for f in os.listdir(current_dir) if f.endswith('.csv')]
    
    if not csv_files:
        print("No CSV files found in the current directory.")
        return
    
    print(f"Found {len(csv_files)} CSV file(s) to process.")
    
    for csv_file in csv_files:
        try:
            print(f"\nProcessing {csv_file}...")
            
            # Extract ConsCode from filename
            cons_code = csv_file.split('_')[1] if '_' in csv_file else "Unknown"
            
            # Get parish name and create filename
            parish_name = get_parish_name(cons_code)
            output_pdf = f"{cons_code} {parish_name}_.pdf"
            
            # Remove any characters that might cause issues in filenames
            output_pdf = output_pdf.replace('/', '-').replace('\\', '-')
            
            # Read and process the CSV file
            df = read_csv_data(csv_file)
            create_labels_pdf(df, output_pdf)
            print(f"Successfully generated: {output_pdf}")
            
        except Exception as e:
            print(f"Error processing {csv_file}: {e}")
            continue

if __name__ == "__main__":
    process_csv_files()
