# KRA Data Extraction System - Production Deployment

## Quick Start
1. Upload your KRA documents (PDF/Word)
2. System automatically extracts data
3. Results saved to persistent database
4. Download individual batches or complete database

## Features
- 100% accurate extraction of KRA tax notice data
- Automatic database storage with deduplication
- Support for multiple document formats
- Real-time processing with progress tracking

## Technology Stack
- Python 3.11
- Streamlit for web interface
- OpenCV + Tesseract for OCR
- Excel-based persistent database
- Docker containerization

## Database
The system maintains a persistent Excel database (`kra_master_database.xlsx`) that stores all extraction results across sessions.

## Support
For issues or questions, please check the application logs or contact support.