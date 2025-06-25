# PDF Highlight Extractor

A Python program to extract highlighted text (yellow highlights) from PDF files and save them to text or Word files.

## Features

- ✅ Extract yellow highlighted text from PDF files
- ✅ Easy-to-use graphical interface
- ✅ Command line version available
- ✅ Enhanced version for difficult cases
- ✅ Save results to text (.txt) files
- ✅ Save results to Word (.docx) files
- ✅ Display page numbers for each extracted text

## Requirements

- Python 3.7 or newer
- Required libraries (listed in `requirements.txt`):
  - PyMuPDF (fitz)
  - python-docx
  - tkinter (included with Python)

## Installation

1. Make sure Python is installed on your system
2. Install required libraries:

```bash
pip install -r requirements.txt
```

Or:

```bash
pip install PyMuPDF python-docx
```

## Usage

### Method 1: Easy Launcher (Recommended)

Double-click on: `run_enhanced_fixed.bat`

This will:
- Check Python installation automatically
- Install required libraries if needed
- Present you with 3 options:
  1. **Graphical Interface** (Easy to use)
  2. **Enhanced Version** (Advanced extraction + Debug)
  3. **Simple Command Line**

### Method 2: Graphical Interface

Run the program with GUI:

```bash
python pdf_highlight_extractor.py
```

### Method 3: Command Line

Run the simple command line version:

```bash
python simple_extractor.py path/to/your/file.pdf
```

Or with output file:

```bash
python simple_extractor.py path/to/your/file.pdf output.txt
```

### Method 4: Enhanced Version (For Problem Cases)

If the simple methods don't work, use the enhanced version:

```bash
# With detailed analysis
python enhanced_extractor.py your_file.pdf --debug

# Save results to file
python enhanced_extractor.py your_file.pdf results.txt
```

## Files

### Core Python Files:
- `pdf_highlight_extractor.py`: Main program with graphical interface
- `simple_extractor.py`: Simple command line version
- `enhanced_extractor.py`: **Enhanced version for difficult cases**
- `requirements.txt`: List of required libraries

### Launcher:
- `run_enhanced_fixed.bat`: **Easy launcher with automatic setup**

### Documentation:
- `README.md`: This comprehensive usage guide

## How It Works

1. **Select File**: Choose the PDF file you want to extract highlights from
2. **Extract**: Click "Extract Highlighted Text" button
3. **View**: Extracted text will appear in the results area
4. **Save**: Save results as text or Word file

## Technical Features

### Text Extraction
- Recognizes annotations in PDF files
- Extracts text from highlighted areas
- Verifies highlight color to ensure it's yellow

### Yellow Color Detection
- Checks RGB values of highlights
- Considers highlight yellow if Red > 0.7, Green > 0.7, Blue < 0.3
- Enhanced version accepts multiple color variations

### Enhanced Version Features
- 4 different extraction methods
- Support for multiple highlight types (Highlight, Underline, Squiggly, etc.)
- Debug mode for analyzing PDF structure
- Better color detection algorithms

## Troubleshooting

### Problem: "No highlighted text found"

If you see this message, try these solutions:

#### 1. Use Enhanced Version
```bash
python enhanced_extractor.py your_file.pdf --debug
```

#### 2. Check Highlight Type
The enhanced version supports:
- Highlight (traditional highlighting)
- Underline
- Squiggly (wavy underline)
- StrikeOut
- Square
- FreeText

#### 3. Use Debug Mode
```bash
python enhanced_extractor.py your_file.pdf output.txt --debug
```
This will show you:
- Number of annotations on each page
- Types of annotations found
- Colors used
- Detailed file structure

### Installation Issues
```bash
# If there are issues with PyMuPDF
pip install --upgrade PyMuPDF

# If there are issues with python-docx
pip install --upgrade python-docx
```

### Encoding Issues
- Make sure files are saved with UTF-8 encoding
- Use a text editor that supports UTF-8

## Examples

### Example 1: GUI
```bash
python pdf_highlight_extractor.py
```
1. Click "Browse" to select PDF file
2. Click "Extract Highlighted Text"
3. View results and save as needed

### Example 2: Simple Command Line
```bash
# Extract with display only
python simple_extractor.py document.pdf

# Extract with saving
python simple_extractor.py document.pdf extracted_highlights.txt
```

### Example 3: Enhanced Version
```bash
# Debug mode
python enhanced_extractor.py document.pdf --debug

# Save results
python enhanced_extractor.py document.pdf results.txt
```

## Support

If you encounter any issues:

1. Make sure all requirements are installed
2. Check PDF file format
3. Try different PDF files for testing
4. Use the enhanced version with debug mode

## License

This program is free for personal and educational use.

## Notes

- Works best with PDF files containing extractable text
- May not work with scanned PDFs without OCR
- Supports all common fonts and encodings

---

**This program was developed to help students and researchers easily extract important text from PDF files.**
