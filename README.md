# KRA Data Extraction System

A Streamlit-based application for extracting structured data from KRA (Kenya Revenue Authority) tax notice letters.

## Features

- **PDF Processing**: Handles both digital and scanned PDFs
- **OCR Integration**: Uses Tesseract for text extraction from images
- **Data Extraction**: Extracts 8 key fields:
  - Date
  - PIN (Tax ID)
  - Taxpayer Name
  - Notice Title
  - Total Tax Amount
  - Year
  - KRA Station
  - Officer Name
- **Excel Export**: Download results as formatted Excel files
- **Debug Mode**: Detailed processing information

## Quick Start

1. **Double-click** `start_kra_extractor.bat` to launch the application
2. **Open browser** to http://localhost:8501
3. **Upload PDF** files using the file uploader
4. **Enable Debug Mode** (optional) for detailed processing info
5. **Download results** as Excel files

## System Requirements

- Windows 10/11
- Python 3.8+ (included in virtual environment)
- Tesseract OCR (installed at C:\Program Files\Tesseract-OCR\)
- Poppler utilities (installed at C:\poppler\)

## File Structure

- `app.py` - Main Streamlit application
- `complete_kra_extractor.py` - Alternative bulk processor
- `requirements.txt` - Python dependencies
- `test_*.py` - Testing utilities
- `.venv/` - Python virtual environment
- `start_kra_extractor.bat` - Quick launcher

## Troubleshooting

If the application doesn't start:
1. Ensure Tesseract is installed: `C:\Program Files\Tesseract-OCR\tesseract.exe`
2. Ensure Poppler is installed: `C:\poppler\poppler-24.02.0\Library\bin\`
3. Run `test_full_setup.py` to verify dependencies

## Tips for Best Results

- Use clear, high-quality PDF scans
- Ensure PDFs contain actual KRA tax notices
- Enable Debug Mode to see extraction details
- Try different PDFs if results are poor

Created: September 2025

## Features

- **PDF Upload**: Upload 2-page scanned KRA tax notice PDFs
- **OCR Processing**: Convert PDF pages to images and extract text using Tesseract
- **Data Extraction**: Extract structured fields using regex patterns:
  - **Page 1**: Date, PIN, Taxpayer Name, Notice Title, Year, Total Tax, Station
  - **Page 2**: Officer Name
- **DataFrame Display**: View extracted data in a formatted table
- **Excel Export**: Download results as an Excel file

## Installation

### Prerequisites

1. **Python 3.8+**
2. **Tesseract OCR** - Install from [GitHub releases](https://github.com/UB-Mannheim/tesseract/wiki)
   - Windows: Download and install the latest version
   - Add Tesseract to your system PATH

### Setup

1. **Clone or download this project**
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

### For Windows Users

If you encounter issues with Tesseract, you may need to specify the path in your code:
```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

## Usage

1. **Start the application**:
   ```bash
   streamlit run app.py
   ```

2. **Access the web interface**:
   - Open your browser to `http://localhost:8501`

3. **Process a PDF**:
   - Upload a 2-page KRA tax notice PDF
   - Click "Process PDF" 
   - Review extracted data
   - Download Excel file with results

## File Structure

```
KRA DATA EXTRACTION/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md          # This file
```

## Extracted Fields

### Page 1 Fields:
- **Date**: Notice date (e.g., "26TH AUGUST, 2025")
- **PIN**: Tax identification number (e.g., "P052148271F")
- **Taxpayer Name**: Company or individual name
- **Notice Title**: Full notice title under Tax Procedures Act
- **Year**: Tax year (e.g., "2023-2024")
- **Total Tax**: Tax amount from Total Tax row
- **Station**: KRA station (e.g., "LODWAR")

### Page 2 Fields:
- **Officer Name**: Signing officer (e.g., "MR LOMUKE EKUTAN")

## Error Handling

- Missing fields are left blank instead of causing crashes
- PDF processing errors are displayed to the user
- OCR errors are handled gracefully

## Dependencies

- `streamlit`: Web application framework
- `pandas`: Data manipulation and analysis
- `openpyxl`: Excel file handling
- `pdf2image`: PDF to image conversion
- `pytesseract`: OCR text extraction
- `pillow`: Image processing

## Troubleshooting

### Common Issues:

1. **Tesseract not found**: Ensure Tesseract is installed and in PATH
2. **Poor OCR results**: Ensure PDF quality is good and pages are clear
3. **Missing fields**: Regex patterns may need adjustment for different document formats

### Performance Tips:

- Use high-quality scanned PDFs for better OCR results
- Ensure text is clear and readable in the original document
- Check that the PDF has exactly 2 pages

## License

This project is provided as-is for educational and practical use.