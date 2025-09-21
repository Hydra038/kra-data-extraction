# KRA Data Extraction System - Multi-Format Processor

A comprehensive Streamlit web application that processes KRA (Kenya Revenue Authority) tax notice letters and extracts structured data from both **PDF and Word documents** (.pdf, .docx, .doc).

## 🆕 New Features

- **Multi-Format Support**: Process both PDF and Word documents
- **Folder Batch Processing**: Process entire folders of documents at once
- **Enhanced Word Processing**: Support for .docx and .doc files
- **Improved Extraction**: Better regex patterns for higher accuracy
- **Smart File Detection**: Automatic file type detection and processing

## Features

- **Multi-Format Processing**: Handles PDF (.pdf) and Word (.docx, .doc) documents
- **Intelligent Text Extraction**: 
  - Digital PDF text extraction
  - OCR for scanned PDFs
  - Native Word document text extraction
- **Batch Processing Options**:
  - Upload multiple individual files
  - Process entire folders of documents
- **Advanced Field Extraction**: Extracts 8 key fields using optimized regex patterns
- **Excel Export**: Download comprehensive results as formatted Excel files
- **Debug Mode**: Detailed logging and processing information
- **User-Friendly Interface**: Clean Streamlit web interface with progress tracking

## Extracted Fields

The system extracts the following 8 fields from KRA tax notices:

1. **Date** - Notice date (multiple formats supported)
2. **PIN** - Tax Identification Number (format: A123456789B)
3. **Taxpayer Name** - Company or individual name
4. **Notice Title** - Subject/RE line of the notice
5. **Total Tax** - Amount due (with proper number formatting)
6. **Year** - Tax year or period
7. **Officer Name** - KRA officer name
8. **Station** - KRA station/office

## Installation & Setup

### Prerequisites

- Python 3.7 or higher
- Windows 10/11 (for current setup)

### Quick Start

1. **Clone or download** this repository to your computer

2. **Run the multi-format setup script**:
   ```
   Double-click: start_multi_format_extractor.bat
   ```

   This will automatically:
   - Install required Python packages (including Word processing libraries)
   - Download and configure Tesseract OCR
   - Download and configure Poppler utilities
   - Launch the multi-format web application

3. **Open your browser** to: http://localhost:8501

### Manual Installation

If the automatic setup doesn't work, you can install manually:

```bash
# Install Python packages (including Word processing)
pip install -r requirements.txt

# Install Tesseract OCR
# Download from: https://github.com/UB-Mannheim/tesseract/wiki

# Install Poppler (for PDF to image conversion)
python install_poppler.py
```

## Usage

### Option 1: Individual File Processing

1. **Launch the application**:
   - Run `start_multi_format_extractor.bat` or
   - Run `streamlit run multi_format_extractor.py` in terminal

2. **Select "Individual Files" mode**

3. **Upload documents**:
   - Click "Browse files" button
   - Select one or more files
   - Supported formats: .pdf, .docx, .doc

4. **Process documents**:
   - Click "Process Uploaded Files" button
   - Monitor progress bar
   - View extracted data in results table

### Option 2: Folder Batch Processing

1. **Select "Folder Batch Processing" mode**

2. **Enter folder path**:
   - Type the full path to your documents folder
   - Example: `C:\Documents\KRA_Files`

3. **Process folder**:
   - Click "Process Folder" button
   - System will find all supported documents
   - Process all files automatically

4. **Export results**:
   - Click "Download Excel Report" to save results
   - Excel file includes summary statistics and detailed data

## File Structure

```
KRA DATA EXTRACTION/
├── multi_format_extractor.py       # NEW: Multi-format processor (PDF + Word)
├── start_multi_format_extractor.bat # NEW: Multi-format launcher
├── app.py                          # Original PDF-only processor
├── complete_kra_extractor.py       # Alternative bulk processor
├── requirements.txt                # Updated with Word processing libs
├── start_kra_extractor.bat        # Original launcher
├── install_poppler.py             # Poppler installation script
├── test_full_setup.py             # System test script
├── test_ocr_setup.py              # OCR test script
├── create_test_word_doc.py         # Word document test creator
├── test_kra_document.docx          # Sample Word document
├── README.md                      # This file
└── .venv/                         # Virtual environment
```

## Supported File Types

| Format | Extension | Processing Method | Notes |
|--------|-----------|-------------------|-------|
| PDF (Digital) | .pdf | Direct text extraction | Fastest processing |
| PDF (Scanned) | .pdf | OCR via Tesseract | Requires good image quality |
| Word Document | .docx | Native text extraction | Modern Word format |
| Word Document | .doc | Native text extraction | Legacy Word format |

## Processing Modes

### Individual Files Mode
- Upload specific documents via web interface
- Preview files before processing
- Ideal for occasional processing

### Folder Batch Processing Mode
- Process all documents in a folder automatically
- Recursively finds supported file types
- Perfect for bulk processing workflows

## Troubleshooting

### Common Issues

1. **Word processing libraries missing**:
   - Click "Install Word Processing Dependencies" in sidebar
   - Or run: `pip install python-docx docx2txt`
   - Restart application after installation

2. **Tesseract not found**:
   - Make sure Tesseract is installed in `C:\Program Files\Tesseract-OCR`
   - Download from: https://github.com/UB-Mannheim/tesseract/wiki

3. **Poppler utilities missing**:
   - Run `python install_poppler.py` to install automatically
   - Or download manually from: https://poppler.freedesktop.org/

4. **Folder not found**:
   - Ensure folder path is correct and exists
   - Use full absolute paths (e.g., `C:\Documents\KRA_Files`)
   - Check folder permissions

5. **Poor extraction results**:
   - Enable "Debug Mode" to see processing details
   - Ensure document quality is good (clear text, good contrast)
   - Check supported file formats

### Test Your Setup

Create and test with a sample Word document:

```bash
python create_test_word_doc.py
```

Run the full system test:

```bash
python test_full_setup.py
```

## Technical Details

### Processing Pipeline

1. **File Detection**: Automatically identify file types (.pdf, .docx, .doc)
2. **Text Extraction**: 
   - PDF: Digital text extraction → OCR fallback
   - Word: Native document parsing via python-docx/docx2txt
3. **Field Extraction**: Apply improved regex patterns
4. **Results Compilation**: Aggregate data from all documents
5. **Export**: Generate comprehensive Excel reports

### Word Document Processing

- **python-docx**: Primary library for .docx files (paragraph extraction)
- **docx2txt**: Fallback library for broader compatibility
- **Automatic fallback**: If one method fails, try the other
- **Error handling**: Graceful degradation with detailed error messages

### Performance Comparison

| Document Type | Processing Time | Accuracy |
|---------------|----------------|----------|
| Digital PDF | 0.5-2 seconds | Very High |
| Scanned PDF | 5-15 seconds | High (depends on quality) |
| Word Document | 0.2-1 second | Very High |

### Enhanced Extraction Patterns

- **Improved PIN validation**: Better format checking
- **Multi-format date parsing**: Handles various date formats
- **Better name extraction**: Enhanced company name detection
- **Robust amount parsing**: Handles different number formats
- **Station identification**: Expanded location database

## Migration from Original System

If you're using the original PDF-only system:

1. **Keep existing files**: Original `app.py` remains functional
2. **Try new system**: Use `multi_format_extractor.py` for enhanced features
3. **Gradual transition**: Both systems can coexist
4. **Same data format**: Excel exports are compatible

## License

This project is for educational and business use. Please ensure compliance with data protection regulations when processing sensitive tax documents.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Use the built-in dependency installer in the sidebar
3. Run test scripts to diagnose problems
4. Enable debug mode for detailed processing logs
5. Check the Dependencies Status in the sidebar