@echo off
chcp 65001 > nul
echo ========================================
echo      مستخرج النصوص المحددة من PDF
echo ========================================
echo.

REM التحقق من وجود Python
python --version > nul 2>&1
if errorlevel 1 (
    echo خطأ: Python غير مثبت على النظام
    echo يرجى تثبيت Python أولاً من: https://python.org
    pause
    exit /b 1
)

REM التحقق من وجود المكتبات المطلوبة
echo جاري التحقق من المكتبات المطلوبة...
python -c "import fitz, docx" > nul 2>&1
if errorlevel 1 (
    echo تثبيت المكتبات المطلوبة...
    pip install PyMuPDF python-docx
    if errorlevel 1 (
        echo خطأ في تثبيت المكتبات
        pause
        exit /b 1
    )
)

REM تشغيل البرنامج
echo تشغيل مستخرج النصوص المحددة...
echo.
python pdf_highlight_extractor.py

if errorlevel 1 (
    echo.
    echo حدث خطأ أثناء تشغيل البرنامج
    pause
) else (
    echo.
    echo تم إغلاق البرنامج بنجاح
)

pause
