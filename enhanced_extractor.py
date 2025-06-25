#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Enhanced PDF Highlight Extractor - Debug Version
Advanced version for extracting highlighted text from PDF files
"""

import fitz  # PyMuPDF
import sys
import os
from datetime import datetime


def debug_pdf_structure(pdf_path):
    """Analyze PDF file structure to understand highlight types"""
    
    print("🔍 Analyzing PDF file structure...")
    print("=" * 60)
    
    try:
        doc = fitz.open(pdf_path)
        
        for page_num in range(min(3, len(doc))):  # Analyze first 3 pages only
            page = doc[page_num]
            print(f"\n📄 Page {page_num + 1}:")
            print("-" * 40)
            
            # Analyze annotations
            annotations = page.annots()
            print(f"Number of annotations: {len(annotations)}")
            
            for i, annot in enumerate(annotations):
                print(f"\n  Annotation #{i+1}:")
                print(f"    Type: {annot.type}")
                print(f"    Rectangle: {annot.rect}")
                
                # Try to get color
                try:
                    colors = annot.colors
                    print(f"    Colors: {colors}")
                except:
                    print(f"    Colors: Not available")
                
                # Try to get content
                try:
                    content = annot.content
                    print(f"    Content: {content}")
                except:
                    print(f"    Content: Not available")
                
                # Try to extract text
                try:
                    text = page.get_textbox(annot.rect)
                    print(f"    Extracted text: {text[:100]}...")
                except:
                    print(f"    Text: Extraction failed")
            
            # Analyze drawings
            try:
                drawings = page.get_drawings()
                print(f"\nNumber of drawings: {len(drawings)}")
                
                for i, drawing in enumerate(drawings[:5]):  # First 5 drawings only
                    print(f"  Drawing #{i+1}: {drawing.get('type', 'Unknown')}")
                    if 'fill' in drawing:
                        print(f"    Fill color: {drawing['fill']}")
                    if 'rect' in drawing:
                        print(f"    Rectangle: {drawing['rect']}")
            except Exception as e:
                print(f"Error analyzing drawings: {e}")
            
            # Analyze texts
            try:
                text_dict = page.get_text("dict")
                blocks_with_color = 0
                
                for block in text_dict.get("blocks", []):
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                if span.get("color", 0) != 0:
                                    blocks_with_color += 1
                
                print(f"Colored texts: {blocks_with_color}")
                
            except Exception as e:
                print(f"Error analyzing texts: {e}")
        
        doc.close()
        print("\n" + "=" * 60)
        
    except Exception as e:
        print(f"Error analyzing file: {e}")


def extract_all_highlights(pdf_path, output_path=None):
    """Extract all types of highlights from PDF using multiple methods"""
    
    if not os.path.exists(pdf_path):
        print(f"❌ Error: File not found: {pdf_path}")
        return []
    
    try:
        print(f"📂 Opening file: {os.path.basename(pdf_path)}")
        doc = fitz.open(pdf_path)
        all_extracts = []
        
        print(f"📊 Number of pages: {len(doc)}")
        
        # Method 1: Extract from annotations
        print("\n🎯 Method 1: Searching in annotations...")
        annotations_found = extract_from_annotations(doc, all_extracts)
        
        # Method 2: Extract from colored drawings
        print("\n🎨 Method 2: Searching in colored drawings...")
        drawings_found = extract_from_drawings(doc, all_extracts)
        
        # Method 3: Extract colored texts
        print("\n🌈 Method 3: Searching in colored texts...")
        colored_text_found = extract_colored_texts(doc, all_extracts)
        
        # Method 4: Comprehensive search
        print("\n🔍 Method 4: Comprehensive search...")
        comprehensive_found = extract_comprehensive(doc, all_extracts)
        
        doc.close()
        
        # Remove duplicates
        unique_extracts = remove_duplicates(all_extracts)
        
        print(f"\n📈 Extraction statistics:")
        print(f"  Annotations: {annotations_found}")
        print(f"  Colored drawings: {drawings_found}")
        print(f"  Colored texts: {colored_text_found}")
        print(f"  Comprehensive search: {comprehensive_found}")
        print(f"  Total before removing duplicates: {len(all_extracts)}")
        print(f"  Total after removing duplicates: {len(unique_extracts)}")
        
        # Display results
        display_results(unique_extracts)
        
        # Save results
        if output_path and unique_extracts:
            save_results(unique_extracts, pdf_path, output_path)
        
        return unique_extracts
        
    except Exception as e:
        print(f"❌ Error processing file: {str(e)}")
        return []


def extract_from_annotations(doc, extracts):
    """Extract from annotations"""
    found = 0
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        annotations = page.annots()
        
        for annot in annotations:
            try:
                annot_type = annot.type[1] if len(annot.type) > 1 else annot.type[0]
                
                # Accept all annotation types that might contain highlights
                if annot_type in ['Highlight', 'Squiggly', 'Underline', 'StrikeOut', 
                                'Square', 'FreeText', 'Text', 'Note', 'Polygon']:
                    
                    text = extract_text_from_annotation(page, annot)
                    
                    if text and text.strip():
                        extract_info = {
                            'page': page_num + 1,
                            'text': text.strip(),
                            'method': f'Annotation-{annot_type}',
                            'color': get_annotation_color(annot),
                            'rect': list(annot.rect)
                        }
                        extracts.append(extract_info)
                        found += 1
                        print(f"    ✓ Page {page_num + 1}: {text[:50]}...")
                        
            except Exception as e:
                print(f"    ✗ Error in annotation: {e}")
    
    return found


def extract_from_drawings(doc, extracts):
    """Extract from colored drawings"""
    found = 0
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        try:
            drawings = page.get_drawings()
            
            for drawing in drawings:
                if 'fill' in drawing and drawing['fill']:
                    fill_color = drawing['fill']
                    
                    # التحقق من أن اللون فاتح (قد يكون تحديد)
                    if is_light_color(fill_color):
                        rect = drawing.get('rect')
                        if rect:
                            # توسيع المنطقة قليلاً
                            expanded_rect = fitz.Rect(
                                rect[0] - 2, rect[1] - 2,
                                rect[2] + 2, rect[3] + 2
                            )
                            
                            text = page.get_textbox(expanded_rect)
                            
                            if text and text.strip():
                                extract_info = {
                                    'page': page_num + 1,
                                    'text': text.strip(),
                                    'method': 'Drawing',
                                    'color': fill_color,
                                    'rect': rect
                                }
                                extracts.append(extract_info)
                                found += 1
                                print(f"    ✓ Page {page_num + 1}: {text[:50]}...")
                                
        except Exception as e:
            print(f"    ✗ Error in drawings page {page_num + 1}: {e}")
    
    return found


def extract_colored_texts(doc, extracts):
    """Extract colored texts"""
    found = 0
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        try:
            text_dict = page.get_text("dict")
            
            for block in text_dict.get("blocks", []):
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            # التحقق من وجود لون مختلف
                            if span.get("color", 0) != 0 or span.get("flags", 0) != 0:
                                text = span.get("text", "").strip()
                                
                                if text and len(text) > 3:  # نص ذو معنى
                                    extract_info = {
                                        'page': page_num + 1,
                                        'text': text,
                                        'method': 'ColoredText',
                                        'color': span.get("color", 0),
                                        'flags': span.get("flags", 0),
                                        'rect': span.get("bbox", [])
                                    }
                                    extracts.append(extract_info)
                                    found += 1
                                    print(f"    ✓ Page {page_num + 1}: {text[:50]}...")
                                    
        except Exception as e:
            print(f"    ✗ Error in colored texts page {page_num + 1}: {e}")
    
    return found


def extract_comprehensive(doc, extracts):
    """Comprehensive search for any distinctive content"""
    found = 0
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        try:
            # Search for text in specific areas (might be highlighted)
            words = page.get_text("words")
            
            # Group adjacent words
            lines = []
            current_line = []
            last_y = None
            
            for word in words:
                x0, y0, x1, y1, text, block_no, line_no, word_no = word
                
                if last_y is None or abs(y0 - last_y) < 5:  # Same line approximately
                    current_line.append(text)
                else:
                    if current_line:
                        lines.append(" ".join(current_line))
                    current_line = [text]
                
                last_y = y0
            
            if current_line:
                lines.append(" ".join(current_line))
            
            # Search for important lines (might be highlighted)
            for line_text in lines:
                if (len(line_text) > 10 and 
                    (any(keyword in line_text.lower() for keyword in 
                     ['important', 'note', 'key', 'main', 'primary', 'essential']) or
                     line_text.isupper() or  # Text in uppercase
                     line_text.count('*') > 0 or  # Contains asterisks
                     line_text.count('-') > 2)):  # Contains many dashes
                    
                    extract_info = {
                        'page': page_num + 1,
                        'text': line_text.strip(),
                        'method': 'Comprehensive',
                        'reason': 'Pattern-based detection'
                    }
                    extracts.append(extract_info)
                    found += 1
                    print(f"    ✓ Page {page_num + 1}: {line_text[:50]}...")
                    
        except Exception as e:
            print(f"    ✗ Error in comprehensive search page {page_num + 1}: {e}")
    
    return found


def extract_text_from_annotation(page, annot):
    """Extract text from annotation using multiple methods"""
    
    rect = annot.rect
    
    # Method 1: Direct text
    text = page.get_textbox(rect)
    if text and text.strip():
        return text
    
    # Method 2: Expand area
    expanded_rect = fitz.Rect(rect.x0 - 3, rect.y0 - 3, rect.x1 + 3, rect.y1 + 3)
    text = page.get_textbox(expanded_rect)
    if text and text.strip():
        return text
    
    # Method 3: Extract overlapping words
    words = page.get_text("words")
    overlapping_text = ""
    
    for word in words:
        word_rect = fitz.Rect(word[:4])
        if word_rect.intersects(rect):
            overlapping_text += word[4] + " "
    
    if overlapping_text.strip():
        return overlapping_text
    
    # Method 4: Annotation content itself
    try:
        content = annot.content
        if content:
            return content
    except:
        pass
    
    return ""


def get_annotation_color(annot):
    """Get annotation color"""
    try:
        colors = annot.colors
        if colors:
            return colors.get("stroke", colors.get("fill", None))
    except:
        pass
    return None


def is_light_color(color):
    """Check if color is light"""
    if not color:
        return False
    
    if isinstance(color, (list, tuple)) and len(color) >= 3:
        # حساب السطوع
        brightness = sum(color[:3]) / len(color[:3])
        return brightness > 0.5
    
    return False


def remove_duplicates(extracts):
    """Remove duplicate entries"""
    unique = []
    seen_texts = set()
    
    for extract in extracts:
        text_clean = extract['text'].strip().lower()
        
        # Ignore very short or duplicate texts
        if len(text_clean) > 5 and text_clean not in seen_texts:
            seen_texts.add(text_clean)
            unique.append(extract)
    
    return unique


def display_results(extracts):
    """Display results"""
    if not extracts:
        print("\n❌ No highlighted text found in this file.")
        return
    
    print(f"\n✅ Found {len(extracts)} highlighted text(s)!")
    print("=" * 60)
    
    for i, extract in enumerate(extracts, 1):
        print(f"\n[{i}] Page {extract['page']} - Method: {extract['method']}")
        print("-" * 50)
        print(extract['text'])
        if 'color' in extract and extract['color']:
            print(f"Color: {extract['color']}")


def save_results(extracts, pdf_path, output_path):
    """Save results to file"""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("Highlighted Text from PDF - Enhanced Version\n")
            f.write("=" * 60 + "\n")
            f.write(f"Source file: {os.path.basename(pdf_path)}\n")
            f.write(f"Extraction date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Number of extracted texts: {len(extracts)}\n")
            f.write("=" * 60 + "\n\n")
            
            for i, extract in enumerate(extracts, 1):
                f.write(f"[{i}] Page {extract['page']} - Method: {extract['method']}\n")
                f.write("-" * 50 + "\n")
                f.write(f"{extract['text']}\n")
                if 'color' in extract and extract['color']:
                    f.write(f"Color: {extract['color']}\n")
                f.write("-" * 50 + "\n\n")
        
        print(f"\n💾 Results saved to: {output_path}")
        
    except Exception as e:
        print(f"❌ Error saving file: {str(e)}")


def main():
    """Main function"""
    print("🚀 PDF Highlight Extractor - Enhanced Version")
    print("=" * 60)
    
    if len(sys.argv) < 2:
        print("Usage:")
        print(f"python {sys.argv[0]} <PDF_file_path> [output_file_path] [--debug]")
        print("\nOptions:")
        print("  --debug    Display detailed analysis of file structure")
        print("\nExamples:")
        print(f"python {sys.argv[0]} document.pdf")
        print(f"python {sys.argv[0]} document.pdf output.txt")
        print(f"python {sys.argv[0]} document.pdf output.txt --debug")
        return
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith('--') else None
    debug_mode = '--debug' in sys.argv
    
    # Run detailed analysis if requested
    if debug_mode:
        debug_pdf_structure(pdf_path)
        print("\n" + "="*60 + "\n")
    
    # Run enhanced extraction
    extracts = extract_all_highlights(pdf_path, output_path)
    
    if extracts:
        print(f"\n🎉 Completed successfully! Extracted {len(extracts)} text(s).")
    else:
        print("\n😔 No highlighted text found.")
        print("💡 Try using the --debug option to analyze file structure")


if __name__ == "__main__":
    main()
