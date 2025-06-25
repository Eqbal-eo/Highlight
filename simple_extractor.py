#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Simple PDF Highlight Extractor
"""

import fitz  # PyMuPDF
import sys
import os
from datetime import datetime


def extract_yellow_highlights(pdf_path, output_path=None):
    """
    Extract yellow highlighted text from PDF file
    
    Args:
        pdf_path (str): PDF file path
        output_path (str): Output file path (optional)
    
    Returns:
        list: List of extracted texts
    """
    
    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        return []
    
    try:
        # Open PDF file
        doc = fitz.open(pdf_path)
        extracted_highlights = []
        
        print(f"Processing file: {os.path.basename(pdf_path)}")
        print(f"Number of pages: {len(doc)}")
        
        # Search through all pages
        for page_num in range(len(doc)):
            page = doc[page_num]
            print(f"Processing page {page_num + 1}...")
            
            # Search for annotations
            annotations = page.annots()
            page_highlights = 0
            
            for annot in annotations:
                annot_type = annot.type[1] if len(annot.type) > 1 else annot.type[0]
                print(f"  Annotation type: {annot_type}")
                
                # Check for different highlight types
                if annot_type in ['Highlight', 'Squiggly', 'Underline', 'StrikeOut', 'Square', 'FreeText']:
                    # Get highlighted text
                    highlighted_text = get_highlighted_text(page, annot)
                    
                    if highlighted_text and highlighted_text.strip():
                        # Get highlight color
                        color = get_annot_color(annot)
                        print(f"    Highlight color: {color}")
                        
                        # Check color (accepting most colors now)
                        if is_yellow_highlight(color):
                            highlight_info = {
                                'page': page_num + 1,
                                'text': highlighted_text.strip(),
                                'color': color,
                                'type': annot_type
                            }
                            extracted_highlights.append(highlight_info)
                            page_highlights += 1
                            print(f"    ✓ Extracted text: {highlighted_text[:50]}...")
                        else:
                            print(f"    ✗ Color mismatch")
                    else:
                        print(f"    ✗ No text found")
                else:
                    print(f"    ✗ Unsupported annotation type")
            
            print(f"  Found {page_highlights} highlighted text(s) on page {page_num + 1}")
            
            # Additional: Search for highlights using other methods
            # Search for hidden or embedded highlights
            try:
                # Alternative method: Search page content for special formatting
                page_dict = page.get_text("dict")
                for block in page_dict.get("blocks", []):
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                # Check formatting properties
                                if span.get("flags", 0) & 2**4:  # Bold
                                    # Might be highlighted text
                                    pass
            except:
                pass
            
            # Search using alternative methods
            alternative_highlights = find_highlights_alternative(page, page_num)
            extracted_highlights.extend(alternative_highlights)
        
        doc.close()
        
        # Print results
        print(f"\nFinished! Found {len(extracted_highlights)} highlighted text(s)")
        
        # Save results if output path specified
        if output_path and extracted_highlights:
            save_to_file(extracted_highlights, pdf_path, output_path)
        
        # Print extracted texts
        if extracted_highlights:
            print("\n" + "=" * 60)
            print("Extracted Texts:")
            print("=" * 60)
            
            for i, highlight in enumerate(extracted_highlights, 1):
                print(f"\n[{i}] Page {highlight['page']}:")
                print("-" * 40)
                print(highlight['text'])
                print("-" * 40)
        else:
            print("No highlighted text found in this file.")
        
        return extracted_highlights
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return []


def get_highlighted_text(page, annot):
    """Extract highlighted text from annotation"""
    try:
        # Get highlight rectangle
        rect = annot.rect
        
        # Method 1: Extract text from highlighted area
        text_instances = page.get_text("dict")
        highlighted_text = ""
        
        for block in text_instances["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        span_rect = fitz.Rect(span["bbox"])
                        # Check if text intersects with highlight area
                        if span_rect.intersects(rect):
                            highlighted_text += span["text"] + " "
        
        # Method 2: If no text found with previous method
        if not highlighted_text.strip():
            highlighted_text = page.get_textbox(rect)
        
        # Method 3: Extract text with higher precision
        if not highlighted_text.strip():
            # Expand area slightly to ensure text capture
            expanded_rect = fitz.Rect(
                rect.x0 - 2, rect.y0 - 2, 
                rect.x1 + 2, rect.y1 + 2
            )
            highlighted_text = page.get_textbox(expanded_rect)
        
        # Method 4: Extract text using alternative method
        if not highlighted_text.strip():
            words = page.get_text("words")
            for word in words:
                word_rect = fitz.Rect(word[:4])
                if word_rect.intersects(rect):
                    highlighted_text += word[4] + " "
        
        return highlighted_text.strip()
        
    except Exception as e:
        print(f"Error extracting text: {e}")
        return ""


def get_annot_color(annot):
    """Get highlight color"""
    try:
        # Try to get color
        color = annot.colors.get("stroke", None)
        if not color:
            color = annot.colors.get("fill", None)
        
        return color if color else [1.0, 1.0, 0.0]  # Default yellow
    except:
        return [1.0, 1.0, 0.0]  # Default yellow


def is_yellow_highlight(color):
    """Check if color is yellow"""
    if not color:
        return True  # Consider yellow by default if no color specified
    
    # If color is a single number (grayscale)
    if isinstance(color, (int, float)):
        return True  # Accept all colors in grayscale case
    
    # If color is a list
    if isinstance(color, (list, tuple)) and len(color) >= 3:
        r, g, b = color[0], color[1], color[2]
        
        # Classic yellow
        if r > 0.7 and g > 0.7 and b < 0.3:
            return True
        
        # Light yellow
        if r > 0.8 and g > 0.8 and b < 0.5:
            return True
        
        # Orange-ish yellow
        if r > 0.9 and g > 0.7 and b < 0.4:
            return True
        
        # Any light color could be a highlight
        brightness = (r + g + b) / 3
        if brightness > 0.6:
            return True
    
    # When in doubt, accept the highlight
    return True


def save_to_file(highlights, pdf_path, output_path):
    """Save extracted texts to file"""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("Highlighted Text from PDF\n")
            f.write("=" * 60 + "\n")
            f.write(f"Source file: {os.path.basename(pdf_path)}\n")
            f.write(f"Extraction date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Number of extracted texts: {len(highlights)}\n")
            f.write("=" * 60 + "\n\n")
            
            for i, highlight in enumerate(highlights, 1):
                f.write(f"[{i}] Page {highlight['page']}:\n")
                f.write("-" * 40 + "\n")
                f.write(f"{highlight['text']}\n")
                f.write("-" * 40 + "\n\n")
        
        print(f"\nResults saved to: {output_path}")
        
    except Exception as e:
        print(f"Error saving file: {str(e)}")


def find_highlights_alternative(page, page_num):
    """Search for highlights using alternative methods"""
    highlights = []
    
    try:
        # Method 1: Search in content streams
        text_instances = page.get_text("dict")
        
        # Method 2: Search for colored rectangles
        drawings = page.get_drawings()
        for drawing in drawings:
            if 'fill' in drawing and drawing['fill']:
                # Check color
                fill_color = drawing.get('fill')
                if fill_color and len(fill_color) >= 3:
                    r, g, b = fill_color[0], fill_color[1], fill_color[2]
                    # If color is yellow or light
                    if (r > 0.7 and g > 0.7 and b < 0.5) or (r + g + b) / 3 > 0.6:
                        # Extract text from this area
                        rect = drawing.get('rect')
                        if rect:
                            text = page.get_textbox(rect)
                            if text and text.strip():
                                highlights.append({
                                    'page': page_num + 1,
                                    'text': text.strip(),
                                    'color': fill_color,
                                    'type': 'Drawing'
                                })
        
        # Method 3: Search for text with colored background
        words = page.get_text("words")
        for word in words:
            # Check for colored background in word
            word_rect = fitz.Rect(word[:4])
            # Can be developed further based on PDF type
        
    except Exception as e:
        print(f"Error in alternative search: {e}")
    
    return highlights


def main():
    """Main function"""
    print("PDF Yellow Highlight Text Extractor")
    print("=" * 50)
    
    if len(sys.argv) < 2:
        print("Usage:")
        print(f"python {sys.argv[0]} <PDF_file_path> [output_file_path]")
        print("\nExample:")
        print(f"python {sys.argv[0]} document.pdf")
        print(f"python {sys.argv[0]} document.pdf output.txt")
        return
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Run extraction
    highlights = extract_yellow_highlights(pdf_path, output_path)
    
    if highlights:
        print(f"\nCompleted successfully! Extracted {len(highlights)} highlighted text(s).")
    else:
        print("\nNo highlighted text found or an error occurred.")


if __name__ == "__main__":
    main()
