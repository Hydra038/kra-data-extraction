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
    page_title="KRA Data Extractor - Multi-Format",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    
    st.title("📊 KRA Data Extractor - Multi-Format Processor")
    st.markdown("Extract structured data from KRA tax notices in PDF and Word documents")
    
    # Sidebar for processing options
    st.sidebar.header("🔧 Processing Options")
    
    processing_mode = st.sidebar.radio(
        "Select Processing Mode:",
        ["📄 Individual Files", "📁 Folder Batch Processing"],
        help="Choose between uploading individual files or processing all documents in a folder"
    )
    
    # Install dependencies button
    if st.sidebar.button("📦 Install Word Processing Dependencies"):
        with st.spinner("Installing python-docx and docx2txt..."):
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx", "docx2txt"])
                st.sidebar.success("Dependencies installed successfully!")
                st.sidebar.info("Please restart the application to use Word processing features.")
            except Exception as e:
                st.sidebar.error(f"Failed to install dependencies: {e}")
    
    # Check dependencies status
    st.sidebar.markdown("### 📋 Dependencies Status")
    st.sidebar.write(f"🔸 **python-docx**: {'✅ Available' if DOCX_AVAILABLE else '❌ Missing'}")
    st.sidebar.write(f"🔸 **docx2txt**: {'✅ Available' if DOCX2TXT_AVAILABLE else '❌ Missing'}")
    
    if processing_mode == "📄 Individual Files":
        st.header("📄 Individual File Processing")
        
        # Database Information
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
                        help="Download complete database with all historical records"
                    )
        
        if db_stats['total_records'] > 0:
            st.info(f"📅 Last updated: {db_stats['last_updated']} | 📊 Date range: {db_stats['date_range']}")
        
        st.subheader("📄 Upload Documents")
        st.info("💾 All extractions are automatically saved to the database with duplicate detection")
        
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
    
    st.header("📊 Extraction Results")
    
    # Create DataFrame from current results
    current_df = pd.DataFrame(results)
    
    # Apply deduplication to current batch
    deduplicated_current = deduplicate_dataframe(current_df)
    
    if len(deduplicated_current) < len(current_df):
        st.info(f"🔍 Removed {len(current_df) - len(deduplicated_current)} duplicate(s) from current batch")
    
    # Save to database automatically
    st.info("💾 Saving results to database...")
    total_records, new_records, duplicates_removed = save_to_database(deduplicated_current, "multi_format_extractor")
    
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
    
    # Display current batch results
    st.subheader("📋 Current Batch Results")
    
    # Show the data in a nice table
    st.dataframe(
        deduplicated_current,
        use_container_width=True,
        hide_index=True
    )
    
    # Summary statistics for current batch
    st.subheader("📈 Current Batch Summary")
    col1, col2, col3, col4 = st.columns(4)
    
    total_files = len(results)
    successful = len([r for r in results if any(r.get(field, '') for field in ['Date', 'PIN', 'Taxpayer_Name', 'Year', 'Officer_Name', 'Station'])])
    success_rate = (successful / total_files * 100) if total_files > 0 else 0
    
    with col1:
        st.metric("Files Processed", total_files)
    with col2:
        st.metric("Successful Extractions", successful)
    with col3:
        st.metric("Success Rate", f"{success_rate:.1f}%")
    with col4:
        st.metric("Records Added to DB", new_records)
    
    # Download options
    st.subheader("📥 Download Options")
    col1, col2 = st.columns(2)
    
    with col1:
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
                help=f"Download complete database with all historical records ({total_records} total records)",
                key=f"download_db_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
        else:
            st.error("Failed to export database")
    
if __name__ == "__main__":
    main()