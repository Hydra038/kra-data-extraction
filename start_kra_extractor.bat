@echo off
echo Starting KRA Data Extraction System...
echo.

REM Activate virtual environment
call .venv\Scripts\activate.bat

REM Add Poppler to PATH for current session
set PATH=%PATH%;C:\poppler\poppler-24.02.0\Library\bin

REM Start Streamlit app
streamlit run app.py --server.port 8501

pause