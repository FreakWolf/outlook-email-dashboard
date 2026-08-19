@echo off
title Outlook Email Dashboard
echo ============================================
echo    Outlook Email Dashboard
echo ============================================
echo.

:: Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

:: Install dependencies if needed
echo Checking dependencies...
pip install -r "%~dp0requirements.txt" --quiet

echo.
echo Starting dashboard...
echo Make sure Outlook is open!
echo.
echo Dashboard will open in your browser at http://localhost:8501
echo Press Ctrl+C to stop the server.
echo.

streamlit run "%~dp0dashboard.py" --server.headless true
pause
