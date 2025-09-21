@echo off
echo Starting KRA Multi-Format Data Extractor...
echo =========================================

:: Navigate to the application directory
cd /d "c:\Users\wisem\OneDrive\Desktop\KRA DATA EXTRACTION"

:: Set PATH to include Tesseract and Poppler
set PATH=%PATH%;C:\Program Files\Tesseract-OCR;C:\tools\poppler-24.07.0\Library\bin

:: Check if Tesseract is available
echo Checking Tesseract OCR installation...
tesseract --version >nul 2>&1
if %errorlevel% neq 0 (
    echo WARNING: Tesseract OCR not found in PATH!
    echo Please install Tesseract OCR from: https://github.com/UB-Mannheim/tesseract/wiki
    pause
)

:: Check if Poppler is available
echo Checking Poppler utilities...
pdftoppm -h >nul 2>&1
if %errorlevel% neq 0 (
    echo WARNING: Poppler utilities not found in PATH!
    echo Installing Poppler via Python script...
    python install_poppler.py
)

:: Launch the multi-format extractor
echo.
echo Launching KRA Multi-Format Data Extractor...
echo You can process both PDF and Word documents (.pdf, .docx, .doc)
echo.
streamlit run multi_format_extractor.py --server.port 8501 --server.address localhost

:: Keep window open if there's an error
if %errorlevel% neq 0 (
    echo.
    echo Error occurred. Press any key to exit...
    pause >nul
)