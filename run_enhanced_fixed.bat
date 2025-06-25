@echo off
setlocal enabledelayedexpansion
chcp 65001 > nul

echo ========================================
echo    PDF Highlight Extractor - Enhanced
echo ========================================
echo.

REM Check Python installation
python --version > nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed on the system
    echo Please install Python first from: https://python.org
    pause
    exit /b 1
)

REM Check required libraries
echo Checking required libraries...
python -c "import fitz, docx" > nul 2>&1
if errorlevel 1 (
    echo Installing required libraries...
    pip install PyMuPDF python-docx
    if errorlevel 1 (
        echo Error installing libraries
        pause
        exit /b 1
    )
)

echo.
echo Choose execution method:
echo 1. Graphical Interface (Easy to use)
echo 2. Enhanced Version (Advanced extraction)
echo 3. Simple Version (Command line)
echo.
set /p choice="Enter your choice (1-3): "

if "!choice!"=="1" (
    echo Running graphical interface...
    python pdf_highlight_extractor.py
) else if "!choice!"=="2" (
    echo.
    set /p pdf_file="Enter PDF file path (or drag it here): "
    REM Remove quotes from file path
    set pdf_file=!pdf_file:"=!
    if exist "!pdf_file!" (
        echo Running enhanced version with detailed analysis...
        python enhanced_extractor.py "!pdf_file!" extracted_highlights.txt --debug
    ) else (
        echo File not found: !pdf_file!
    )
) else if "!choice!"=="3" (
    echo.
    set /p pdf_file="Enter PDF file path (or drag it here): "
    REM Remove quotes from file path
    set pdf_file=!pdf_file:"=!
    if exist "!pdf_file!" (
        echo Running simple version...
        python simple_extractor.py "!pdf_file!" extracted_highlights.txt
    ) else (
        echo File not found: !pdf_file!
    )
) else (
    echo Invalid choice!
    goto :eof
)

echo.
echo Execution completed
pause
