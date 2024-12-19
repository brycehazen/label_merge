
# Label Merge PDF Script

## Overview
This script processes CSV files to generate PDF files containing labels. It reads CSV files from the current directory, validates the data, formats it, and outputs a PDF with formatted labels, including a cover page.

## Features
- Automatically detects encoding of input CSV files.
- Validates and cleans data from CSV files.
- Generates labels in a standardized format.
- Creates a cover sheet for each PDF summarizing the label data.
- Supports barcodes and formatted address labels.

## Requirements
- Python 3.x
- Required Python Libraries:
  - pandas
  - chardet
  - reportlab

## Usage
1. Place your CSV files in the same directory as the script.
2. Ensure the CSV files include the required columns:
   - `ConsID`
   - `PrimAddText`
   - `AddrLines`
   - `AddrCity`
   - `AddrState`
   - `AddrZIP`
   - `Year_2024`
3. Run the script using:
   ```bash
   python label_merge_pdf.py
   ```
4. Output PDF files will be generated in the same directory.

## Notes
- The script attempts to load a logo image named `Untitled.png` for the cover page. If the image is not available, a warning is displayed.
- Barcodes are generated using the `ConsID` field.
- The script adjusts font sizes dynamically to fit label content.

## Troubleshooting
- Ensure that all required columns are present in the CSV files.
- If encoding errors occur, verify the encoding of the CSV files and adjust as necessary.
- Check for any missing or invalid data in the CSV files.

## Customization
- To change the logo on the cover page, replace `Untitled.png` with a different image of your choice.
- Modify font styles and sizes by editing the `create_labels_pdf` and `create_cover_sheet` functions.

## License
This script is provided as-is, without warranty of any kind.
