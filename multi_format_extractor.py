"""
KRA Data Extraction System - Enhanced Multi-Format Processor
===========================================================

Processes multiple documents (PDF, Word) from folders and extracts KRA data.
Supports both individual file uploads and batch folder processing.

Author: GitHub Copilot
Date: September 20, 2025
"""

import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes, convert_from_path
from PIL import Image
import re
import io
from pathlib import Path
import tempfile
import os

# Import deduplication utilities
from deduplication_utils import deduplicate_dataframe, compare_extraction_methods
# Import database utilities
from database_utils import save_to_database, get_database_stats, export_database_to_excel
import fitz  # PyMuPDF for efficient PDF handling
import logging
import traceback
from datetime import datetime
import sys
import subprocess
import zipfile
from typing import List, Dict, Any

# Word document processing
try:
    import docx
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import docx2txt
    DOCX2TXT_AVAILABLE = True
except ImportError:
    DOCX2TXT_AVAILABLE = False

# Configure Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Configure page layout
st.set_page_config(
    page_title="KRA Data Extraction System",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS inspired by KRA website
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

# Debug mode toggle
DEBUG_MODE = st.sidebar.toggle("🐛 Debug Mode", value=False, help="Enable detailed debugging information")

def log_debug(message, data=None):
    """Log debug information if debug mode is enabled"""
    if DEBUG_MODE:
        timestamp = datetime.now().strftime("%H:%M:%S")
        st.sidebar.write(f"🐛 **{timestamp}**: {message}")
        if data:
            st.sidebar.json(data)
        logger.info(f"DEBUG: {message}")

def log_error(message, error):
    """Log error information"""
    error_msg = f"ERROR: {message} - {str(error)}"
    st.error(error_msg)
    logger.error(error_msg)
    if DEBUG_MODE:
        st.sidebar.error(f"❌ {message}")
        st.sidebar.code(str(error))

def extract_text_from_word(file_path_or_bytes):
    """
    Extract text from Word document (.docx)
    
    Args:
        file_path_or_bytes: File path string or bytes object
        
    Returns:
        tuple: (extracted_text, extraction_method)
    """
    try:
        log_debug("Starting Word document text extraction")
        
        if isinstance(file_path_or_bytes, str):
            # File path
            if DOCX_AVAILABLE:
                doc = Document(file_path_or_bytes)
                text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                if text.strip():
                    log_debug(f"Word extraction successful via python-docx: {len(text)} characters")
                    return text, "docx_paragraphs"
            
            if DOCX2TXT_AVAILABLE:
                text = docx2txt.process(file_path_or_bytes)
                if text.strip():
                    log_debug(f"Word extraction successful via docx2txt: {len(text)} characters")
                    return text, "docx2txt"
        else:
            # Bytes object (uploaded file)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                tmp_file.write(file_path_or_bytes)
                tmp_path = tmp_file.name
            
            try:
                if DOCX_AVAILABLE:
                    doc = Document(tmp_path)
                    text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                    if text.strip():
                        log_debug(f"Word extraction successful via python-docx: {len(text)} characters")
                        return text, "docx_paragraphs"
                
                if DOCX2TXT_AVAILABLE:
                    text = docx2txt.process(tmp_path)
                    if text.strip():
                        log_debug(f"Word extraction successful via docx2txt: {len(text)} characters")
                        return text, "docx2txt"
            finally:
                os.unlink(tmp_path)
        
        log_debug("Word text extraction failed - no text found")
        return "", "word_extraction_failed"
        
    except Exception as e:
        log_error("Error extracting text from Word document", e)
        return "", f"word_error: {str(e)}"

def extract_text_from_pdf(pdf_file):
    """
    Extract text from PDF - detect if digital or scanned
    
    Args:
        pdf_file: File path, bytes, or Streamlit UploadedFile object
        
    Returns:
        tuple: (extracted_text, extraction_method)
    """
    try:
        log_debug("Starting PDF text extraction")
        
        # Handle different input types
        if hasattr(pdf_file, 'read'):
            # Streamlit UploadedFile
            pdf_bytes = pdf_file.read()
            pdf_file.seek(0)  # Reset for potential reuse
        elif isinstance(pdf_file, bytes):
            pdf_bytes = pdf_file
        else:
            # File path
            with open(pdf_file, 'rb') as f:
                pdf_bytes = f.read()
        
        # Try digital text extraction first
        try:
            pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            text = ""
            for page in pdf_doc:
                text += page.get_text() + "\n"
            pdf_doc.close()
            
            if len(text.strip()) > 100:  # Meaningful text threshold
                log_debug(f"Digital PDF extraction successful: {len(text)} characters")
                return text, "digital_pdf"
        except Exception as e:
            log_debug(f"Digital PDF extraction failed: {e}")
        
        # Fall back to OCR
        try:
            images = convert_from_bytes(pdf_bytes, dpi=300)
            text = ""
            for i, image in enumerate(images):
                page_text = pytesseract.image_to_string(image)
                text += page_text + "\n"
                log_debug(f"OCR page {i+1}: {len(page_text)} characters")
            
            if text.strip():
                log_debug(f"PDF OCR extraction successful: {len(text)} characters")
                return text, "pdf_ocr"
        except Exception as e:
            log_error("PDF OCR extraction failed", e)
        
        return "", "pdf_extraction_failed"
        
    except Exception as e:
        log_error("Error in PDF text extraction", e)
        return "", f"pdf_error: {str(e)}"

def extract_kra_fields(text):
    """
    Final improved KRA field extraction with only the specified 6 fields and comprehensive fixes
    
    Args:
        text: Extracted text from document
        
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
    
    try:
        log_debug("Starting KRA field extraction")
        
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
        
        fields_found = sum(1 for v in data.values() if v)
        log_debug(f"KRA field extraction completed: {fields_found}/6 fields found")
        
        return data
        
    except Exception as e:
        log_error("Error in KRA field extraction", e)
        return data

def process_document(file_path_or_uploaded, file_name):
    """
    Process a single document (PDF or Word) and extract KRA data
    
    Args:
        file_path_or_uploaded: File path or uploaded file object
        file_name: Name of the file for identification
        
    Returns:
        dict: Processing results with extracted data
    """
    # Initialize result with only the 6 core fields
    result = {field: '' for field in ['Date', 'PIN', 'Taxpayer_Name', 'Year', 'Officer_Name', 'Station']}
    
    try:
        log_debug(f"Processing document: {file_name}")
        
        # Determine file type
        file_ext = Path(file_name).suffix.lower()
        
        # Extract text based on file type
        if file_ext == '.pdf':
            text, method = extract_text_from_pdf(file_path_or_uploaded)
        elif file_ext in ['.docx', '.doc']:
            if hasattr(file_path_or_uploaded, 'read'):
                # Uploaded file
                text, method = extract_text_from_word(file_path_or_uploaded.read())
            else:
                # File path
                text, method = extract_text_from_word(file_path_or_uploaded)
        else:
            log_debug(f"Unsupported file type: {file_ext}")
            return result
        
        if not text:
            log_debug("No text extracted from document")
            return result
        
        # Extract KRA fields
        kra_data = extract_kra_fields(text)
        result.update(kra_data)
        
        log_debug(f"Document processed successfully")
        
    except Exception as e:
        log_error(f"Error processing document {file_name}", e)
    
    return result

def process_folder(folder_path):
    """
    Process all supported documents in a folder
    
    Args:
        folder_path: Path to folder containing documents
        
    Returns:
        list: List of processing results
    """
    results = []
    supported_extensions = ['.pdf', '.docx', '.doc']
    
    try:
        folder = Path(folder_path)
        if not folder.exists():
            st.error(f"Folder does not exist: {folder_path}")
            return results
        
        # Find all supported files
        files = []
        for ext in supported_extensions:
            files.extend(folder.glob(f"*{ext}"))
            files.extend(folder.glob(f"*{ext.upper()}"))
        
        if not files:
            st.warning(f"No supported documents found in folder. Looking for: {', '.join(supported_extensions)}")
            return results
        
        st.info(f"Found {len(files)} documents to process")
        
        # Process each file
        progress_bar = st.progress(0)
        status_placeholder = st.empty()
        
        for i, file_path in enumerate(files):
            progress = (i + 1) / len(files)
            progress_bar.progress(progress)
            status_placeholder.info(f"Processing {i+1}/{len(files)}: {file_path.name}")
            
            result = process_document(str(file_path), file_path.name)
            results.append(result)
        
        progress_bar.progress(1.0)
        status_placeholder.success(f"Completed processing {len(files)} documents")
        
    except Exception as e:
        log_error("Error processing folder", e)
    
    return results

def main():
    """Main application function"""
    
    # Beautiful header inspired by KRA website
    st.markdown("""
    <div class="kra-header">
        <h1>🏛️ KRA Data Extraction System</h1>
        <p>Professional data extraction from tax notices and financial documents</p>
        <p style="font-size: 1rem; margin-top: 1rem;">
            <strong>Kenya Revenue Authority</strong> • Advanced Document Processing
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
    
    # Modern sidebar
    st.sidebar.markdown("## 🔧 Processing Options")
    st.sidebar.markdown("---")
    
    processing_mode = st.sidebar.radio(
        "**Select Processing Mode:**",
        ["📄 Individual Files", "📁 Folder Batch Processing"],
        help="Choose between uploading individual files or processing all documents in a folder"
    )
    
    # Dependencies section with modern styling
    st.sidebar.markdown("### 📋 System Status")
    status_color_docx = "🟢" if DOCX_AVAILABLE else "🔴"
    status_color_docx2txt = "🟢" if DOCX2TXT_AVAILABLE else "🔴"
    
    st.sidebar.markdown(f"""
    - {status_color_docx} **Microsoft Word**: {'Ready' if DOCX_AVAILABLE else 'Installing...'}
    - {status_color_docx2txt} **Document Parser**: {'Ready' if DOCX2TXT_AVAILABLE else 'Installing...'}
    - 🟢 **PDF Processing**: Ready
    - 🟢 **OCR Engine**: Ready
    """)
    
    # Install dependencies with better UX
    if not (DOCX_AVAILABLE and DOCX2TXT_AVAILABLE):
        if st.sidebar.button("� Install Word Dependencies", help="Install required packages for Word document processing"):
            with st.spinner("Installing Microsoft Word processing capabilities..."):
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx", "docx2txt"])
                    st.sidebar.success("✅ Dependencies installed! Please restart the app.")
                except Exception as e:
                    st.sidebar.error(f"❌ Installation failed: {e}")
    
    st.sidebar.markdown("---")
    
    if processing_mode == "📄 Individual Files":
        # Individual file processing with beautiful cards
        st.markdown("""
        <div class="kra-card">
            <h3>📄 Individual File Processing</h3>
            <p>Upload PDF or Word documents for immediate data extraction and database storage</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Full database download section
        try:
            db_stats = get_database_stats()
            if db_stats['total_records'] > 0:
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); 
                                padding: 1rem; border-radius: 8px; margin-bottom: 1rem;">
                        <p><strong>📅 Last Updated:</strong> {db_stats['last_updated']}</p>
                        <p><strong>📊 Date Range:</strong> {db_stats['date_range']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    excel_data = export_database_to_excel()
                    if excel_data:
                        st.download_button(
                            label="📥 Download Complete Database",
                            data=excel_data,
                            file_name=f"KRA_Complete_Database_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            help="Download complete database with all historical records"
                        )
        except:
            pass
        
        # Upload section with modern styling
        st.markdown("### � Document Upload")
        st.markdown("""
        <div style="background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); 
                    padding: 1rem; border-radius: 8px; border-left: 4px solid #3b82f6; margin-bottom: 1rem;">
            <p><strong>� Automatic Features:</strong></p>
            <ul style="margin-bottom: 0;">
                <li>✅ Database auto-save with duplicate detection</li>
                <li>🎯 Smart data extraction and validation</li>
                <li>📊 Real-time processing statistics</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        # Initialize session state for file management
        if 'processed_files' not in st.session_state:
            st.session_state.processed_files = False
        if 'processing_results' not in st.session_state:
            st.session_state.processing_results = None
        
        # File uploader with key to enable clearing
        uploaded_files = st.file_uploader(
            "Upload documents (PDF or Word)",
            type=['pdf', 'docx', 'doc'],
            accept_multiple_files=True,
            help="Upload one or more PDF or Word documents containing KRA tax notices",
            key="file_uploader_main"
        )
        
        if uploaded_files and not st.session_state.processed_files:
            st.success(f"📁 {len(uploaded_files)} file(s) uploaded")
            
            if st.button("🚀 Process Uploaded Files", type="primary", key="process_button_main"):
                results = []
                
                progress_bar = st.progress(0)
                status_placeholder = st.empty()
                
                for i, uploaded_file in enumerate(uploaded_files):
                    progress = (i + 1) / len(uploaded_files)
                    progress_bar.progress(progress)
                    status_placeholder.info(f"Processing {i+1}/{len(uploaded_files)}: {uploaded_file.name}")
                    
                    # Reset file pointer
                    uploaded_file.seek(0)
                    result = process_document(uploaded_file, uploaded_file.name)
                    results.append(result)
                
                progress_bar.progress(1.0)
                status_placeholder.success("Processing completed!")
                
                # Store results and mark as processed
                st.session_state.processing_results = results
                st.session_state.processed_files = True
                
                # Clear the file uploader by rerunning
                st.rerun()
        
        # Display results if processing is complete
        if st.session_state.processed_files and st.session_state.processing_results:
            display_results(st.session_state.processing_results)
            
            # Add button to reset for new files
            if st.button("🔄 Process New Files", type="secondary", key="reset_button_main"):
                st.session_state.processed_files = False
                st.session_state.processing_results = None
                st.rerun()
    
    else:  # Folder processing
        st.header("📁 Folder Batch Processing")
        
        # Database Information for folder processing too
        st.subheader("📊 Database Status")
        
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
                        help="Download complete database with all historical records",
                        key="folder_download_db"
                    )
        
        if db_stats['total_records'] > 0:
            st.info(f"📅 Last updated: {db_stats['last_updated']} | 📊 Date range: {db_stats['date_range']}")
        
        st.info("💾 All extractions are automatically saved to the database with duplicate detection")
        
        # Initialize session state for folder processing
        if 'folder_processed' not in st.session_state:
            st.session_state.folder_processed = False
        if 'folder_results' not in st.session_state:
            st.session_state.folder_results = None
        if 'folder_path_processed' not in st.session_state:
            st.session_state.folder_path_processed = ""
        
        folder_path = st.text_input(
            "Enter folder path containing documents:",
            placeholder="C:\\path\\to\\your\\documents",
            help="Enter the full path to the folder containing PDF and Word documents",
            key="folder_path_input"
        )
        
        if folder_path and not st.session_state.folder_processed:
            if st.button("🚀 Process Folder", type="primary", key="process_folder_button"):
                results = process_folder(folder_path)
                if results:
                    st.session_state.folder_results = results
                    st.session_state.folder_processed = True
                    st.session_state.folder_path_processed = folder_path
                    st.rerun()
        
        # Display folder results if processing is complete
        if st.session_state.folder_processed and st.session_state.folder_results:
            st.success(f"✅ Processed folder: {st.session_state.folder_path_processed}")
            display_results(st.session_state.folder_results)
            
            # Add button to reset for new folder
            if st.button("🔄 Process New Folder", type="secondary", key="reset_folder_button"):
                st.session_state.folder_processed = False
                st.session_state.folder_results = None
                st.session_state.folder_path_processed = ""
                st.rerun()

def display_results(results):
    """Display processing results and save to database"""
    if not results:
        st.warning("No results to display")
        return
    
    # Modern results header
    st.markdown("""
    <div class="kra-card">
        <h3>📊 Extraction Results</h3>
        <p>Processing complete with automatic database integration</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Create DataFrame from current results
    current_df = pd.DataFrame(results)
    
    # Apply deduplication to current batch
    deduplicated_current = deduplicate_dataframe(current_df)
    
    if len(deduplicated_current) < len(current_df):
        st.success(f"🔍 Removed {len(current_df) - len(deduplicated_current)} duplicate(s) from current batch")
    
    # Save to database automatically
    with st.spinner("💾 Saving results to database..."):
        total_records, new_records, duplicates_removed = save_to_database(deduplicated_current, "multi_format_extractor")
    
    # Display save results in beautiful cards
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #10b981 0%, #065f46 100%); 
                    color: white; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3 style="margin: 0; color: white;">{new_records}</h3>
            <p style="margin: 0; opacity: 0.9;">✅ New Records Added</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #3b82f6 0%, #1e40af 100%); 
                    color: white; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3 style="margin: 0; color: white;">{total_records:,}</h3>
            <p style="margin: 0; opacity: 0.9;">📊 Total Database Records</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        if duplicates_removed > 0:
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); 
                        color: white; padding: 1rem; border-radius: 8px; text-align: center;">
                <h3 style="margin: 0; color: white;">{duplicates_removed}</h3>
                <p style="margin: 0; opacity: 0.9;">🔍 Duplicates Merged</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #10b981 0%, #065f46 100%); 
                        color: white; padding: 1rem; border-radius: 8px; text-align: center;">
                <h3 style="margin: 0; color: white;">0</h3>
                <p style="margin: 0; opacity: 0.9;">🎉 No Duplicates</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Display current batch results with modern styling
    st.markdown("### 📋 Current Batch Results")
    
    # Enhanced table display
    if not deduplicated_current.empty:
        st.markdown("""
        <div style="background: white; padding: 1rem; border-radius: 8px; 
                    border: 1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);">
        """, unsafe_allow_html=True)
        
        st.dataframe(
            deduplicated_current,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Date": st.column_config.DateColumn("📅 Date"),
                "PIN": st.column_config.TextColumn("🔢 PIN", width="medium"),
                "Taxpayer_Name": st.column_config.TextColumn("👤 Taxpayer", width="large"),
                "Year": st.column_config.NumberColumn("� Year"),
                "Officer_Name": st.column_config.TextColumn("👥 Officer", width="medium"),
                "Station": st.column_config.TextColumn("🏢 Station", width="medium")
            }
        )
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.warning("No data extracted from the uploaded files")
    
    # Enhanced summary statistics
    st.markdown("### 📈 Processing Statistics")
    
    total_files = len(results)
    successful = len([r for r in results if any(r.get(field, '') for field in ['Date', 'PIN', 'Taxpayer_Name', 'Year', 'Officer_Name', 'Station'])])
    success_rate = (successful / total_files * 100) if total_files > 0 else 0
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%); 
                    color: white; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3 style="margin: 0; color: white;">{total_files}</h3>
            <p style="margin: 0; opacity: 0.9;">📄 Files Processed</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #10b981 0%, #065f46 100%); 
                    color: white; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3 style="margin: 0; color: white;">{successful}</h3>
            <p style="margin: 0; opacity: 0.9;">✅ Successful</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        color = "#10b981" if success_rate >= 80 else "#f59e0b" if success_rate >= 60 else "#ef4444"
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, {color} 0%, {color}dd 100%); 
                    color: white; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3 style="margin: 0; color: white;">{success_rate:.1f}%</h3>
            <p style="margin: 0; opacity: 0.9;">🎯 Success Rate</p>
        </div>
    
    with col4:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%); 
                    color: white; padding: 1rem; border-radius: 8px; text-align: center;">
            <h3 style="margin: 0; color: white;">{new_records}</h3>
            <p style="margin: 0; opacity: 0.9;">💾 DB Records Added</p>
        </div>
        """, unsafe_allow_html=True)
    
    # Modern download section
    st.markdown("### 📥 Export Options")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div style="background: #f8fafc; padding: 1rem; border-radius: 8px; 
                    border: 1px solid #e2e8f0; margin-bottom: 1rem;">
            <h4 style="margin-top: 0; color: #1e3a8a;">📄 Current Session</h4>
            <p style="color: #64748b; margin-bottom: 1rem;">Download results from this processing session only</p>
        """, unsafe_allow_html=True)
        
        # Current batch download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            deduplicated_current.to_excel(writer, sheet_name='Current_Batch_Results', index=False)
            
            # Add batch summary
            summary_data = {
                'Metric': [
                    'Files Processed',
                    'Successful Extractions', 
                    'Success Rate (%)',
                    'Processing Date',
                    'Records in Batch'
                ],
                'Value': [
                    total_files,
                    successful,
                    f"{success_rate:.1f}%",
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    len(deduplicated_current)
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Batch_Summary', index=False)
        
        st.download_button(
            label="📥 Download Current Batch",
            data=output.getvalue(),
            file_name=f"KRA_Current_Batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary",
            use_container_width=True,
            help=f"Download results from this processing session ({len(deduplicated_current)} records)",
            key=f"download_batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col2:
        st.markdown("""
        <div style="background: #f8fafc; padding: 1rem; border-radius: 8px; 
                    border: 1px solid #e2e8f0; margin-bottom: 1rem;">
            <h4 style="margin-top: 0; color: #1e3a8a;">🗄️ Complete Database</h4>
            <p style="color: #64748b; margin-bottom: 1rem;">Download all historical records from the database</p>
        """, unsafe_allow_html=True)
        
        # Full database download
        excel_data = export_database_to_excel()
        if excel_data:
            st.download_button(
                label="📥 Download Complete Database",
                data=excel_data,
                file_name=f"KRA_Complete_Database_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                help="Download complete database with all historical records",
                key=f"download_full_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        else:
            st.warning("Database export temporarily unavailable")
            
        st.markdown("</div>", unsafe_allow_html=True)
    
    # Beautiful footer
    st.markdown("---")
    st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); 
                color: white; padding: 2rem; border-radius: 12px; text-align: center; margin-top: 2rem;">
        <h3 style="color: white; margin-bottom: 1rem;">🏛️ Kenya Revenue Authority</h3>
        <p style="opacity: 0.9; margin-bottom: 0;">
            Professional Data Extraction System • Powered by Advanced AI Technology
        </p>
        <p style="opacity: 0.7; margin-top: 0.5rem; font-size: 0.9rem;">
            © 2025 KRA Data Extraction System • All rights reserved
        </p>
    </div>
    """, unsafe_allow_html=True)
                use_container_width=True,
                help=f"Download complete database with all historical records ({total_records} total records)",
                key=f"download_db_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        else:
            st.error("Failed to export database")
    
if __name__ == "__main__":
    main()