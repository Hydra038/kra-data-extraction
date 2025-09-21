"""
KRA Tax Notice Letter Processing App with Debugging
====================================================

This Streamlit app processes scanned KRA tax notice letters (2-page PDFs) and extracts 
structured data using OCR and regex patterns.

Author: GitHub Copilot
Date: September 19, 2025
"""

import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
import re
import io
from pathlib import Path
import tempfile
import os
import fitz  # PyMuPDF for efficient PDF handling
import logging
import traceback
from datetime import datetime
import sys
import subprocess

# Import deduplication utilities
from deduplication_utils import deduplicate_dataframe
# Import database utilities
from database_utils import save_to_database, get_database_stats, export_database_to_excel

# Configure Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Configure page layout
st.set_page_config(
    page_title="KRA Tax Notice Processor",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS inspired by KRA website (same as multi_format_extractor.py)
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Root variables inspired by KRA colors */
    :root {
        --kra-primary: #1e3a8a;
        --kra-secondary: #3b82f6;
        --kra-accent: #f59e0b;
        --kra-success: #10b981;
        --kra-danger: #ef4444;
        --kra-light: #f8fafc;
        --kra-dark: #1e293b;
        --kra-border: #e2e8f0;
    }
    
    /* Main app styling */
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        font-family: 'Inter', sans-serif;
    }
    
    /* Header styling */
    .kra-header {
        background: linear-gradient(135deg, var(--kra-primary) 0%, var(--kra-secondary) 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.2);
    }
    
    .kra-header h1 {
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        text-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .kra-header p {
        font-size: 1.2rem;
        opacity: 0.9;
        margin-bottom: 0;
    }
    
    /* Sidebar styling */
    .css-1d391kg, .css-1544g2n {
        background: linear-gradient(180deg, var(--kra-light) 0%, #ffffff 100%);
        border-right: 2px solid var(--kra-border);
    }
    
    /* Card styling */
    .kra-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        border: 1px solid var(--kra-border);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
        margin-bottom: 1.5rem;
    }
    
    /* Stats cards */
    .stat-card {
        background: linear-gradient(135deg, var(--kra-primary) 0%, var(--kra-secondary) 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 4px 15px rgba(30, 58, 138, 0.2);
        margin-bottom: 1rem;
    }
    
    .stat-card h3 {
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        color: white;
    }
    
    .stat-card p {
        opacity: 0.9;
        font-weight: 500;
        margin-bottom: 0;
    }
    
    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, var(--kra-accent) 0%, #f59e0b 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(245, 158, 11, 0.3);
        width: 100%;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(245, 158, 11, 0.4);
    }
    
    /* File uploader styling */
    .stFileUploader {
        border: 2px dashed var(--kra-accent);
        border-radius: 12px;
        padding: 2rem;
        background: var(--kra-light);
        text-align: center;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, var(--kra-success) 0%, var(--kra-accent) 100%);
    }
    
    /* Table styling */
    .stDataFrame {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
    }
    
    /* Success/Error messages */
    .stSuccess {
        background: linear-gradient(135deg, var(--kra-success) 0%, #065f46 100%);
        border-radius: 8px;
    }
    
    .stError {
        background: linear-gradient(135deg, var(--kra-danger) 0%, #dc2626 100%);
        border-radius: 8px;
    }
    
    /* Hide Streamlit style */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display:none;}
    header[data-testid="stHeader"] {display:none;}
    
    /* Info boxes */
    .stInfo {
        background: linear-gradient(135deg, var(--kra-secondary) 0%, #3b82f6 100%);
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Debug mode toggle (at the top of sidebar)
DEBUG_MODE = st.sidebar.toggle("🐛 Debug Mode", value=False, help="Enable detailed debugging information")

def log_debug(message, data=None):
    """Log debug information if debug mode is enabled"""
    if DEBUG_MODE:
        timestamp = datetime.now().strftime("%H:%M:%S")
        st.sidebar.write(f"🐛 **{timestamp}**: {message}")
        if data:
            st.sidebar.json(data)
        logger.info(f"DEBUG: {message}")

def log_error(message, error=None):
    """Log error information"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.error(f"❌ **{timestamp}**: {message}")
    if error and DEBUG_MODE:
        st.sidebar.error(f"Error Details: {str(error)}")
        st.sidebar.text(traceback.format_exc())
    logger.error(f"ERROR: {message} - {str(error) if error else ''}")

# Custom CSS for better UI
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, rgb(30,60,114) 0%, rgb(42,82,152) 100%);
        padding: 2rem 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
        color: white;
    }
    
    .upload-section {
        background: #f8f9fa;
        padding: 2rem;
        border-radius: 10px;
        border: 2px dashed #dee2e6;
        margin: 1rem 0;
        text-align: center;
    }
    
    .success-box {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .info-box {
        background: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    
    .stMetric {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .feature-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        margin: 1rem 0;
        border-left: 4px solid rgb(30,60,114);
    }
</style>
""", unsafe_allow_html=True)

def extract_text_from_image(image):
    """
    Extract text from a PIL Image using OCR.
    
    Args:
        image: PIL Image object
        
    Returns:
        str: Extracted text from the image
    """
    try:
        log_debug(f"Extracting text from image", {"image_size": image.size, "mode": image.mode})
        
        # Configure Tesseract for better OCR results
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,()/:&-"'
        
        # Extract text using Tesseract
        text = pytesseract.image_to_string(image, config=custom_config)
        log_debug(f"Text extracted successfully", {"text_length": len(text)})
        
        return text.strip()
        
    except Exception as e:
        log_error("Error extracting text from image", e)
        return ""

def detect_pdf_type(pdf_bytes):
    """
    Detect if PDF contains mostly images (scanned) or text.
    
    Args:
        pdf_bytes: PDF file content as bytes
        
    Returns:
        tuple: (is_scanned, page_count)
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_count = doc.page_count
        
        # Check first page for content type
        page = doc[0]
        text_content = page.get_text().strip()
        image_list = page.get_images()
        
        # If page has minimal text but images, it's likely scanned
        is_scanned = len(text_content) < 100 and len(image_list) > 0
        
        doc.close()
        log_debug(f"PDF analysis", {
            "page_count": page_count,
            "is_scanned": is_scanned,
            "text_length": len(text_content),
            "image_count": len(image_list)
        })
        
        return is_scanned, page_count
        
    except Exception as e:
        log_error("Error analyzing PDF", e)
        return True, 0  # Assume scanned if analysis fails

def extract_text_from_pdf(pdf_file):
    """
    Extract text from PDF with optimized processing based on content type.
    
    Args:
        pdf_file: Uploaded PDF file from Streamlit
        
    Returns:
        tuple: (page1_text, page2_text) or (None, None) if error
    """
    try:
        log_debug("Starting optimized PDF processing", {"file_name": pdf_file.name, "file_size": pdf_file.size})
        
        pdf_bytes = pdf_file.read()
        is_scanned, page_count = detect_pdf_type(pdf_bytes)
        
        if page_count < 2:
            error_msg = f"PDF has {page_count} pages, but 2 pages are required"
            log_error(error_msg)
            st.error(error_msg)
            return None, None
        
        if is_scanned:
            log_debug("PDF detected as scanned document - using OCR conversion")
            # Use pdf2image for scanned documents
            with st.spinner("Processing scanned document with OCR..."):
                log_debug("Converting scanned PDF to images with 300 DPI")
                images = convert_from_bytes(pdf_bytes, dpi=300)
                
                # Display image info in debug mode
                if DEBUG_MODE:
                    for i, img in enumerate(images[:2]):
                        st.sidebar.write(f"📄 Page {i+1}: {img.size[0]}x{img.size[1]} pixels")
                
                with st.spinner("Running OCR on page 1..."):
                    log_debug("Starting OCR on page 1")
                    page1_text = extract_text_from_image(images[0])
                    log_debug(f"Page 1 OCR completed. Text length: {len(page1_text)} characters")
                
                with st.spinner("Running OCR on page 2..."):
                    log_debug("Starting OCR on page 2")
                    page2_text = extract_text_from_image(images[1])
                    log_debug(f"Page 2 OCR completed. Text length: {len(page2_text)} characters")
        else:
            log_debug("PDF detected as text-based document - using direct text extraction (faster)")
            # Use PyMuPDF for text-based documents (much faster)
            with st.spinner("Extracting text directly from PDF..."):
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                
                page1_text = doc[0].get_text().strip()
                page2_text = doc[1].get_text().strip()
                
                doc.close()
                
                log_debug(f"Direct text extraction completed", {
                    "page1_length": len(page1_text),
                    "page2_length": len(page2_text)
                })
        
        # Show text preview in debug mode
        if DEBUG_MODE:
            processing_method = "OCR" if is_scanned else "Direct extraction"
            st.sidebar.write(f"⚡ **Processing Method:** {processing_method}")
            st.sidebar.write("📄 **Page 1 Preview:**")
            st.sidebar.text(page1_text[:200] + "..." if len(page1_text) > 200 else page1_text)
            st.sidebar.write("📄 **Page 2 Preview:**")
            st.sidebar.text(page2_text[:200] + "..." if len(page2_text) > 200 else page2_text)
        
        return page1_text, page2_text
    
    except Exception as e:
        log_error("Error processing PDF", e)
        return None, None

def process_uploaded_file(uploaded_file):
    """
    Process uploaded file (PDF or image) and extract text.
    
    Args:
        uploaded_file: Streamlit uploaded file object
        
    Returns:
        tuple: (page1_text, page2_text) or (None, None) if error
    """
    try:
        file_type = uploaded_file.type
        log_debug(f"Processing uploaded file", {"file_name": uploaded_file.name, "file_type": file_type})
        
        if file_type == "application/pdf":
            return extract_text_from_pdf(uploaded_file)
        
        elif file_type.startswith("image/"):
            # Handle direct image upload
            with st.spinner("Processing image with OCR..."):
                image = Image.open(uploaded_file)
                text = extract_text_from_image(image)
                
                # For single image, treat as page 1, leave page 2 empty
                log_debug(f"Image processed", {"text_length": len(text)})
                return text, ""
        
        else:
            st.error(f"❌ Unsupported file type: {file_type}")
            return None, None
            
    except Exception as e:
        log_error("Error processing uploaded file", e)
        st.error(f"❌ Error processing file: {str(e)}")
        return None, None

def extract_data_from_page2(text):
    """
    Extract structured fields from page 2 text using improved regex patterns.
    
    Args:
        text: OCR extracted text from page 2
        
    Returns:
        dict: Dictionary containing extracted fields
    """
    data = {
        'Date': '',
        'PIN': '',
        'Taxpayer_Name': '',
        'Year': '',
        'Officer_Name': '',
        'Station': ''
    }
    extraction_log = []
    
    try:
        log_debug("Starting unified data extraction from all pages")
        
        # Extract Date (existing patterns work well)
        date_patterns = [
            r'(\d{1,2}(?:ST|ND|RD|TH)\s+[A-Z]+,?\s+\d{4})',
            r'(\d{1,2}[A-Z]{2}\s+[A-Z]{3,9},?\s*\d{4})',
            r'(\d{1,2}\s+[A-Z]{3,9}\s+\d{4})',
            r'(\d{1,2}[-/]\d{1,2}[-/]\d{4})',
        ]
        
        for pattern in date_patterns:
            date_match = re.search(pattern, text, re.IGNORECASE)
            if date_match:
                data['Date'] = date_match.group(1).strip()
                break
        
        # Extract PIN (existing patterns work well)
        pin_patterns = [
            r'PIN[:\s]*([A-Z]\d{9}[A-Z])',
            r'P\.?I\.?N\.?[:\s]*([A-Z]\d{9}[A-Z])',
            r'([A-Z]\d{9}[A-Z])',
        ]
        
        for pattern in pin_patterns:
            pin_match = re.search(pattern, text, re.IGNORECASE)
            if pin_match:
                pin = pin_match.group(1).upper()
                if re.match(r'^[A-Z]\d{9}[A-Z]$', pin):
                    data['PIN'] = pin
                    break
        
        # IMPROVED: Extract Taxpayer Name (handles both companies and individuals)
        improved_taxpayer_patterns = [
            # Pattern 1: Individual name after PIN with comma (like "Peter Kimutai Telengech,")
            r'PIN[:\s]*[A-Z]\d{9}[A-Z][^\n]*\n\s*([A-Za-z][A-Za-z\s]+?),',
            # Pattern 2: Company name with business suffixes
            r'([A-Z][A-Z\s&.,()-]+?(?:LIMITED|LTD|COMPANY|GROUP|CORPORATION|CORP|INC|ENTERPRISES|SERVICES))\s*(?:\n|$|P\.O\.)',
            # Pattern 3: General pattern between PIN and P.O BOX
            r'PIN[:\s]*[A-Z]\d{9}[A-Z]\s*\n\s*([A-Z][A-Z\s&.,()-]{5,}?)\s*\n\s*P\.\s*O\.',
            # Pattern 4: Before P.O. BOX (more general)
            r'([A-Z][A-Z\s&.,()LTD-]{10,}?)\s*\n\s*P\.\s*O\.\s*BOX',
        ]
        
        for pattern in improved_taxpayer_patterns:
            name_match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if name_match:
                name = name_match.group(1).strip()
                name = re.sub(r'\s+', ' ', name)
                name = name.replace('\n', ' ').replace('\r', ' ')
                name = re.sub(r'\s+', ' ', name)
                
                # Enhanced validation for both individual and business names
                words = name.split()
                valid_keywords = ['LIMITED', 'LTD', 'COMPANY', 'GROUP', 'CORPORATION', 'CORP', 'INC', 'ENTERPRISES', 'SERVICES']
                
                is_valid = (
                    len(name) >= 5 and len(name) <= 100 and
                    (
                        # Has business keywords OR
                        any(keyword in name.upper() for keyword in valid_keywords) or
                        # Is likely an individual name (2-4 words, mostly letters)
                        (len(words) >= 2 and len(words) <= 4 and 
                         all(word.isalpha() for word in words))
                    ) and
                    # Contains mostly letters
                    sum(c.isalpha() or c.isspace() for c in name) / len(name) > 0.7
                )
                
                if is_valid:
                    data['Taxpayer_Name'] = name
                    break
        
        # FIXED: Extract Year with proper patterns and business logic
        year_found = False
        
        # Try ONLY explicit tax year mentions (no P.O. BOX ranges)
        explicit_year_patterns = [
            r'(?:tax\s+year|year\s+of\s+income|for\s+the\s+year)[:\s]*(\d{4})',
            # Only year ranges in tax contexts (not P.O. BOX)
            r'(?:tax\s+year|income\s+year|assessment).*?(\d{4}[-–]\d{4})',
        ]
        
        for pattern in explicit_year_patterns:
            year_match = re.search(pattern, text, re.IGNORECASE)
            if year_match:
                year = year_match.group(1).strip()
                if year.isdigit() and 2015 <= int(year) <= 2030:
                    data['Year'] = year
                    year_found = True
                    break
                elif '-' in year and len(year.split('-')[0]) == 4:  # Valid year range
                    data['Year'] = year
                    year_found = True
                    break
        
        # If no explicit year found, use business logic: document year - 1
        if not year_found and data['Date']:
            # Extract year from document date
            doc_year_match = re.search(r'\d{4}', data['Date'])
            if doc_year_match:
                doc_year = int(doc_year_match.group(0))
                # Tax assessments are typically for the previous year
                tax_year = doc_year - 1
                data['Year'] = str(tax_year)
                year_found = True
        
        # IMPROVED: Extract Officer Name (from contact information)
        improved_officer_patterns = [
            # Pattern 1: Direct Officer mention (most direct)
            r'Officer[:\s]*([A-Z][a-zA-Z\s]+?)(?:\n|Contact|Tel|Phone|Email)',
            # Pattern 2: Contact name in "contact X or Y" phrase (most reliable for this doc type)
            r'contact\s+([A-Z][a-z]+\s+[A-Z][a-z]+)\s+or',
            # Pattern 3: After "hesitate to contact"
            r'hesitate\s+to\s+contact\s+([A-Z][a-z]+\s+[A-Z][a-z]+)',
            # Pattern 4: After "contact" before phone number
            r'contact\s+([A-Z][a-z]+\s+[A-Z][a-z]+).*?phone',
            # Pattern 5: Signature name (fallback)
            r'Yours\s+faithfully,.*?\n\s*([A-Z][a-z]+\s+[A-Z][a-z]+)',
            # Pattern 6: General contact pattern
            r'contact\s+([A-Z][a-z]+\s+[A-Z][a-z]+)',
        ]
        
        for pattern in improved_officer_patterns:
            officer_match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if officer_match:
                officer = officer_match.group(1).strip()
                officer = re.sub(r'\s+', ' ', officer)
                
                # Enhanced validation - should be 2-4 words, all letters/spaces
                words = officer.split()
                if (len(words) >= 2 and len(words) <= 4 and 
                    all(word.isalpha() for word in words) and 
                    len(officer) >= 5 and len(officer) <= 50):
                    data['Officer_Name'] = officer
                    break
        
        # Extract Station (existing patterns work well)
        station_patterns = [
            # Priority 1: Station mentioned with "STATION" or "OFFICE" (most specific)
            r'([A-Z]{4,})\s+(?:STATION|OFFICE)',
            # Priority 2: Known KRA stations in document (avoid KRA headquarters in header)
            r'(?:P\.?\s*O\.?\s*BOX\s+\d+[^,\n]*,?\s*)?([A-Z]{4,}(?:LODWAR|MOMBASA|KISUMU|NAKURU|ELDORET|NYERI|MERU|MACHAKOS|KITALE|GARISSA|ISIOLO|MALINDI|KILIFI|EMBU|THIKA|KIAMBU|KAKAMEGA|KERICHO|BOMET|BUNGOMA|WEBUYE|MIGORI|HOMABAY|SIAYA|BUSIA|MARSABIT|MANDERA|WAJIR|MOYALE|KAPENGURIA|MARALAL))',
            # Priority 3: Specific station names (excluding NAIROBI from KRA header)
            r'\b(LODWAR|MOMBASA|KISUMU|NAKURU|ELDORET|NYERI|MERU|MACHAKOS|KITALE|GARISSA|ISIOLO|MALINDI|KILIFI|EMBU|THIKA|KIAMBU|KAKAMEGA|KERICHO|BOMET|BUNGOMA|WEBUYE|MIGORI|HOMABAY|SIAYA|BUSIA|MARSABIT|MANDERA|WAJIR|MOYALE|KAPENGURIA|MARALAL)\b',
            # Priority 4: General location after P.O. BOX (fallback)
            r'P\.?\s*O\.?\s*BOX\s+\d+[-–\s]*\d*[,\s]*([A-Z]{3,})',
        ]
        
        for pattern in station_patterns:
            station_match = re.search(pattern, text, re.IGNORECASE)
            if station_match:
                station = station_match.group(1).strip().upper()
                if len(station) >= 3:
                    data['Station'] = station
                    break
        
        # Log results for each field
        for field in data.keys():
            status = "✅ Found" if data[field] else "❌ Not found"
            extraction_log.append(f"{field}: {status}")
        
        # Log extraction results
        if DEBUG_MODE:
            st.sidebar.write("**📄 Combined Extraction Results:**")
            for log_entry in extraction_log:
                st.sidebar.write(f"• {log_entry}")
        
        fields_found = sum(1 for v in data.values() if v)
        log_debug(f"Combined extraction completed: {fields_found}/6 fields found")
    
    except Exception as e:
        log_error("Error extracting data from pages", e)
    
    return data

def create_dataframe(page1_data, page2_data):
    """
    Combine extracted data from both pages into a pandas DataFrame.
    
    Args:
        page1_data: Dictionary with data from page 1
        page2_data: Dictionary with data from page 2
        
    Returns:
        pandas.DataFrame: Combined data in a single row with only core fields
    """
    try:
        log_debug("Creating DataFrame from extracted data")
        
        # Combine data from both pages, keeping only core fields
        combined_data = {
            'Date': page1_data.get('Date', '') or page2_data.get('Date', ''),
            'PIN': page1_data.get('PIN', '') or page2_data.get('PIN', ''),
            'Taxpayer_Name': page1_data.get('Taxpayer_Name', '') or page2_data.get('Taxpayer_Name', ''),
            'Year': page1_data.get('Year', '') or page2_data.get('Year', ''),
            'Officer_Name': page1_data.get('Officer_Name', '') or page2_data.get('Officer_Name', ''),
            'Station': page1_data.get('Station', '') or page2_data.get('Station', '')
        }
        
        # Log what data was combined
        if DEBUG_MODE:
            st.sidebar.write("**🔗 Combined Data:**")
            for key, value in combined_data.items():
                status = "✅" if value else "❌"
                st.sidebar.write(f"• {key}: {status}")
        
        # Create DataFrame with one row
        df = pd.DataFrame([combined_data])
        
        log_debug(f"DataFrame created successfully", {"columns": len(df.columns), "rows": len(df)})
        
        return df
    
    except Exception as e:
        log_error("Error creating DataFrame", e)
        return pd.DataFrame()

def create_excel_download(df, output_mode="Create New Excel", existing_excel=None):
    """
    Create Excel file for download with optional appending.
    
    Args:
        df: pandas DataFrame to export
        output_mode: "Create New Excel" or "Append to Existing Excel"
        existing_excel: existing Excel file to append to (if mode is append)
        
    Returns:
        bytes: Excel file content
    """
    try:
        log_debug("Creating Excel file for download")
        
        # Handle existing data if appending
        if output_mode == "Append to Existing Excel" and existing_excel is not None:
            try:
                # Read existing Excel file
                existing_excel.seek(0)  # Reset file pointer
                existing_df = pd.read_excel(existing_excel, sheet_name='KRA_Tax_Data')
                
                # Combine existing and new data
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                log_debug(f"Combining data: {len(existing_df)} existing + {len(df)} new = {len(combined_df)} total")
                
                # Apply deduplication to combined data
                final_df = deduplicate_dataframe(combined_df)
                
                duplicates_removed = len(combined_df) - len(final_df)
                if duplicates_removed > 0:
                    log_debug(f"Deduplication removed {duplicates_removed} duplicate records")
                
            except Exception as e:
                log_error("Error reading existing Excel file, creating new file instead", e)
                final_df = df
        else:
            # Even for new files, check for internal duplicates
            final_df = deduplicate_dataframe(df)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='KRA_Tax_Data')
            
            # Add summary sheet
            summary_data = {
                'Metric': ['Total Records', 'Export Date', 'File Handling Mode'],
                'Value': [len(final_df), pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S'), output_mode]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, index=False, sheet_name='Summary')
        
        excel_data = output.getvalue()
        log_debug(f"Excel file created successfully", {"size_bytes": len(excel_data), "records": len(final_df)})
        
        return excel_data
    
    except Exception as e:
        log_error("Error creating Excel file", e)
        return None

# Main App Interface
def main():
    """Main Streamlit application interface."""
    
    # Beautiful header inspired by KRA website
    st.markdown("""
    <div class="kra-header">
        <h1>🏛️ KRA Tax Notice Processor</h1>
        <p>Advanced AI-powered document processing for tax notices and financial documents</p>
        <p style="font-size: 1rem; margin-top: 1rem;">
            <strong>Kenya Revenue Authority</strong> • Single Document Processing
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Database stats in beautiful cards
    try:
        db_stats = get_database_stats()
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div class="stat-card">
                <h3>{db_stats['total_records']:,}</h3>
                <p>📊 Total Records</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="stat-card">
                <h3>{db_stats['unique_taxpayers']:,}</h3>
                <p>👥 Unique Taxpayers</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="stat-card">
                <h3>{db_stats['unique_officers']:,}</h3>
                <p>👤 Tax Officers</p>
            </div>
            """, unsafe_allow_html=True)
            
    except Exception as e:
        st.info("📊 Database statistics will appear here after first extraction")
    
    # Debug information panel
    if DEBUG_MODE:
        st.markdown("---")
        st.markdown("### 🛠️ System Debug Information")
        
        try:
            # Check Tesseract version
            tesseract_version = pytesseract.get_tesseract_version()
            st.success(f"✅ Tesseract OCR: v{tesseract_version}")
        except Exception as e:
            st.error(f"❌ Tesseract Error: {str(e)}")
        
        try:
            # Check pdf2image
            from pdf2image import convert_from_path
            st.success("✅ pdf2image: Available")
        except Exception as e:
            st.error(f"❌ pdf2image Error: {str(e)}")
        
        try:
            # Check Poppler
            result = subprocess.run(['pdftoppm', '-h'], capture_output=True, text=True)
            if result.returncode == 0:
                st.success("✅ Poppler: Available")
            else:
                st.error("❌ Poppler: Not accessible")
        except Exception as e:
            st.error(f"❌ Poppler Error: {str(e)}")
        
        # Display Python packages info
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"Python: {sys.version.split()[0]}")
            st.info(f"Streamlit: {st.__version__}")
        with col2:
            st.info(f"Pandas: {pd.__version__}")
            st.info(f"PIL: {Image.__version__}")
    
    # Sidebar with instructions
    with st.sidebar:
        st.markdown("""
        <div class="feature-card">
            <h2 style="color: rgb(30,60,114); margin-top: 0;">📋 How to Use</h2>
            <div style="line-height: 1.8;">
                <p><strong>1.</strong> 📁 Upload your file</p>
                <p><strong>2.</strong> ⏳ Wait for OCR processing</p>
                <p><strong>3.</strong> 📊 Review extracted data</p>
                <p><strong>4.</strong> 💾 Download Excel results</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="feature-card">
            <h3 style="color: rgb(30,60,114); margin-top: 0;">🎯 Extracted Fields</h3>
            <div style="font-size: 0.9em; line-height: 1.6;">
                <p><strong>📄 From Document:</strong><br>
                • Date & PIN<br>
                • Taxpayer Name<br>
                • Tax Year<br>
                • Officer Name & Station</p>
                
                <p><strong>⚡ Smart Features:</strong><br>
                • Individual vs Company Names<br>
                • Contact-based Officer Extraction<br>
                • Tax Year Business Logic</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="info-box">
            <p style="margin: 0; font-size: 0.9em; text-align: center;">
                <strong>⚙️ Powered by</strong><br>
                OCR (Tesseract) & AI Pattern Matching
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    # File Upload Section
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div class="upload-section">
            <h3 style="color: rgb(30,60,114); margin-top: 0;">📁 Upload Your Document</h3>
            <p style="color: #6c757d; margin-bottom: 1rem;">
                <strong>Supported formats:</strong><br>
                • PDF: 2-page KRA notices (auto-detects scanned vs text)<br>
                • Images: JPG, PNG, TIFF, BMP (single page processing)
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Output options
        # Database Status
        st.markdown("### 📊 Database Status")
        
        # Get database stats
        db_stats = get_database_stats()
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Records", db_stats['total_records'])
        with col2:
            st.metric("Unique Taxpayers", db_stats['unique_taxpayers'])
        with col3:
            st.metric("Unique Stations", db_stats['unique_stations'])
        with col4:
            if db_stats['total_records'] > 0:
                # Add full database download button
                excel_data = export_database_to_excel()
                if excel_data:
                    st.download_button(
                        label="📥 Download Full Database",
                        data=excel_data,
                        file_name=f"KRA_Complete_Database_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        help="Download complete database with all historical records"
                    )
        
        if db_stats['total_records'] > 0:
            st.info(f"📅 Last updated: {db_stats['last_updated']} | 📊 Date range: {db_stats['date_range']}")
        
        st.info("💾 All extractions are automatically saved to the database with duplicate detection")
        
        # Initialize session state for app.py
        if 'app_processed' not in st.session_state:
            st.session_state.app_processed = False
        if 'app_results' not in st.session_state:
            st.session_state.app_results = None
        
        uploaded_file = st.file_uploader(
            "Choose a PDF or Image file", 
            type=['pdf', 'jpg', 'jpeg', 'png', 'tiff', 'bmp'],
            help="Upload your KRA tax notice (PDF, JPG, PNG, TIFF) to extract structured data",
            label_visibility="collapsed",
            key="file_uploader_app"
        )
    
    if uploaded_file is not None and not st.session_state.app_processed:
        # Display file info
        st.markdown(f"""
        <div class="success-box">
            <h4 style="color: #155724; margin: 0 0 0.5rem 0;">✅ File Successfully Uploaded</h4>
            <p style="margin: 0; color: #155724;">
                <strong>📄 {uploaded_file.name}</strong> ({uploaded_file.size:,} bytes)
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Process PDF button
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            process_button = st.button(
                "🔄 Process Document", 
                type="primary", 
                use_container_width=True,
                help="Start extracting data from your PDF",
                key="process_button_app"
            )
        
        if process_button:
            log_debug("Process button clicked - starting document processing")
            
            # Extract text from uploaded file (PDF or image)
            page1_text, page2_text = process_uploaded_file(uploaded_file)
            
            if page1_text is not None and page2_text is not None:
                # Extract structured data
                with st.spinner("Extracting structured data..."):
                    log_debug("Starting structured data extraction")
                    # Combine both pages for extraction
                    combined_text = page1_text + "\n\n" + page2_text
                    extracted_data = extract_data_from_page2(combined_text)
                    log_debug("Structured data extraction completed")
                
                # Create DataFrame
                df = create_dataframe(extracted_data, {})
                
                # Store results in session state
                st.session_state.app_results = {
                    'df': df,
                    'extracted_data': extracted_data,
                    'page1_text': page1_text,
                    'page2_text': page2_text,
                    'filename': uploaded_file.name
                }
                st.session_state.app_processed = True
                st.rerun()
    
    # Display results if processing is complete
    if st.session_state.app_processed and st.session_state.app_results:
        results_data = st.session_state.app_results
        df = results_data['df']
        extracted_data = results_data['extracted_data']
        page1_text = results_data['page1_text']
        page2_text = results_data['page2_text']
        filename = results_data['filename']
        
        st.success(f"✅ Processed document: {filename}")
        
        # Show extracted text in expandable sections
        st.markdown("### 📄 OCR Results")
        
        col1, col2 = st.columns(2)
        
        with col1:
            with st.expander("📄 Page 1 - OCR Text", expanded=False):
                st.text_area("Page 1 Text", page1_text, height=200, label_visibility="collapsed", key="page1_text_display")
        
        with col2:
            with st.expander("📄 Page 2 - OCR Text", expanded=False):
                st.text_area("Page 2 Text", page2_text, height=200, label_visibility="collapsed", key="page2_text_display")
        
        # Save to database automatically
        if not df.empty:
            st.info("💾 Saving results to database...")
            total_records, new_records, duplicates_removed = save_to_database(df, "app_extractor")
            
            # Display save results
            col1, col2, col3 = st.columns(3)
            with col1:
                st.success(f"✅ {new_records} new record(s) added")
            with col2:
                st.info(f"📊 Total database records: {total_records}")
            with col3:
                if duplicates_removed > 0:
                    st.warning(f"🔍 {duplicates_removed} duplicate(s) found and merged")
                else:
                    st.success("🎉 No duplicates found")
        
        # Display results
        st.markdown("### 📊 Extracted Data")
        st.markdown("""
        <div class="feature-card">
            <p style="margin: 0; color: rgb(30,60,114); font-weight: 500;">
                Below is the structured data extracted from your document:
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.dataframe(
            df, 
            use_container_width=True,
            hide_index=True
        )
        
        # Show extraction summary
        st.markdown("### 📈 Extraction Summary")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                "📄 Total Fields", 
                len([v for v in extracted_data.values() if v]),
                help="Number of fields successfully extracted from the document"
            )
        
        with col2:
            st.metric(
                "📄 Required Fields", 
                6,
                help="Total number of fields the system looks for"
            )
        
        with col3:
            total_fields = len([v for v in extracted_data.values() if v])
            success_rate = (total_fields / 6) * 100 if total_fields > 0 else 0
            st.metric(
                "✅ Success Rate", 
                f"{success_rate:.1f}%",
                help="Percentage of fields successfully extracted"
            )
        
        # Download options
        st.markdown("### 📥 Download Options")
        col1, col2 = st.columns(2)
        
        with col1:
            # Current document download
            if not df.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='KRA_Extraction', index=False)
                
                st.download_button(
                    label="📥 Download Current Document",
                    data=output.getvalue(),
                    file_name=f"KRA_Document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary",
                    use_container_width=True,
                    help="Download results from this document only",
                    key=f"download_current_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                )
        
        with col2:
            # Full database download
            excel_data = export_database_to_excel()
            if excel_data:
                st.download_button(
                    label="📥 Download Full Database",
                    data=excel_data,
                    file_name=f"KRA_Complete_Database_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    help="Download complete database with all historical records",
                    key=f"download_db_app_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                )
            else:
                st.error("Failed to export database")
        
        # Success message
        if not df.empty:
            st.markdown("""
            <div class="success-box">
                <h4 style="color: #155724; margin: 0 0 0.5rem 0;">🎉 Processing Complete!</h4>
                <p style="margin: 0; color: #155724;">
                    Data extraction completed and saved to database successfully!
                </p>
            </div>
            """, unsafe_allow_html=True)
        
        # Add button to process new document
        if st.button("🔄 Process New Document", type="secondary", key="reset_app_button"):
            st.session_state.app_processed = False
            st.session_state.app_results = None
            st.rerun()
    
    else:
        # Welcome message when no file is uploaded or after processing
        st.markdown("""
        <div class="info-box">
            <h4 style="color: #0c5460; margin: 0 0 0.5rem 0;">👋 Welcome!</h4>
            <p style="margin: 0; color: #0c5460; text-align: center;">
                Please upload a PDF file above to get started with data extraction.
            </p>
        </div>
        """, unsafe_allow_html=True)

# Run the app
if __name__ == "__main__":
    main()
                
                st.dataframe(
                    df, 
                    use_container_width=True,
                    hide_index=True
                )
                
                # Show extraction summary
                st.markdown("### 📈 Extraction Summary")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric(
                        "📄 Total Fields", 
                        len([v for v in extracted_data.values() if v]),
                        help="Number of fields successfully extracted from the document"
                    )
                
                with col2:
                    st.metric(
                        "📄 Required Fields", 
                        6,
                        help="Total number of fields the system looks for"
                    )
                
                with col3:
                    total_fields = len([v for v in extracted_data.values() if v])
                    success_rate = (total_fields / 6) * 100 if total_fields > 0 else 0
                    st.metric(
                        "✅ Success Rate", 
                        f"{success_rate:.1f}%",
                        help="Percentage of fields successfully extracted"
                    )
                
                # Save to database automatically
                if not df.empty:
                    st.info("💾 Saving results to database...")
                    total_records, new_records, duplicates_removed = save_to_database(df, "app_extractor")
                    
                    # Display save results
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.success(f"✅ {new_records} new record(s) added")
                    with col2:
                        st.info(f"📊 Total database records: {total_records}")
                    with col3:
                        if duplicates_removed > 0:
                            st.warning(f"🔍 {duplicates_removed} duplicate(s) found and merged")
                        else:
                            st.success("🎉 No duplicates found")
                
                # Display results
                st.markdown("### 📊 Extracted Data")
                st.markdown("""
                <div class="feature-card">
                    <p style="margin: 0; color: rgb(30,60,114); font-weight: 500;">
                        Below is the structured data extracted from your document:
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                st.dataframe(
                    df, 
                    use_container_width=True,
                    hide_index=True
                )
                
                # Show extraction summary
                st.markdown("### � Extraction Summary")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric(
                        "📄 Total Fields", 
                        len([v for v in extracted_data.values() if v]),
                        help="Number of fields successfully extracted from the document"
                    )
                
                with col2:
                    st.metric(
                        "📄 Required Fields", 
                        6,
                        help="Total number of fields the system looks for"
                    )
                
                with col3:
                    total_fields = len([v for v in extracted_data.values() if v])
                    success_rate = (total_fields / 6) * 100 if total_fields > 0 else 0
                    st.metric(
                        "✅ Success Rate", 
                        f"{success_rate:.1f}%",
                        help="Percentage of fields successfully extracted"
                    )
                
                # Download options
                st.markdown("### 📥 Download Options")
                col1, col2 = st.columns(2)
                
                with col1:
                    # Current document download
                    if not df.empty:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, sheet_name='KRA_Extraction', index=False)
                        
                        st.download_button(
                            label="📥 Download Current Document",
                            data=output.getvalue(),
                            file_name=f"KRA_Document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="secondary",
                            use_container_width=True,
                            help="Download results from this document only"
                        )
                
                with col2:
                    # Full database download
                    excel_data = export_database_to_excel()
                    if excel_data:
                        st.download_button(
                            label="📥 Download Full Database",
                            data=excel_data,
                            file_name=f"KRA_Complete_Database_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True,
                            help="Download complete database with all historical records"
                        )
                    else:
                        st.error("Failed to export database")
                
                # Success message
                if not df.empty:
                    st.markdown("""
                    <div class="success-box">
                        <h4 style="color: #155724; margin: 0 0 0.5rem 0;">🎉 Processing Complete!</h4>
                        <p style="margin: 0; color: #155724;">
                            Data extraction completed and saved to database successfully!
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    log_debug("Processing completed successfully")
                else:
                    st.markdown("""
                    <div style="background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 8px; padding: 1rem; margin: 1rem 0;">
                        <h4 style="color: #721c24; margin: 0 0 0.5rem 0;">⚠️ No Data Extracted</h4>
                        <p style="margin: 0; color: #721c24;">
                            Unable to extract data from the document. Please check the PDF quality and try again.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    log_debug("No data was extracted from the document")
            
            else:
                st.markdown("""
                <div style="background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 8px; padding: 1rem; margin: 1rem 0;">
                    <h4 style="color: #721c24; margin: 0 0 0.5rem 0;">❌ Processing Failed</h4>
                    <p style="margin: 0; color: #721c24;">
                        Failed to process the PDF document. Please check the file format and try again.
                    </p>
                </div>
                """, unsafe_allow_html=True)
    
    else:
        # Welcome message when no file is uploaded
        st.markdown("""
        <div class="info-box">
            <h4 style="color: #0c5460; margin: 0 0 0.5rem 0;">👋 Welcome!</h4>
            <p style="margin: 0; color: #0c5460; text-align: center;">
                Please upload a PDF file above to get started with data extraction.
            </p>
        </div>
        """, unsafe_allow_html=True)

# Run the app
if __name__ == "__main__":
    main()