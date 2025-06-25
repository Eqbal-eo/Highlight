#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Highlight Text Extractor

This program extracts highlighted text (in yellow) from PDF files
and saves them to a text or Word file
"""

import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime
from docx import Document
from docx.shared import Inches


class PDFHighlightExtractor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Highlight Text Extractor")
        self.root.geometry("600x500")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.pdf_file = None
        self.extracted_highlights = []
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup user interface"""
        # Main title
        title_label = tk.Label(
            self.root, 
            text="PDF Highlight Text Extractor",
            font=("Arial", 16, "bold"),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title_label.pack(pady=20)
        
        # File selection frame
        file_frame = tk.Frame(self.root, bg='#f0f0f0')
        file_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(
            file_frame, 
            text="Select PDF file:",
            font=("Arial", 12),
            bg='#f0f0f0'
        ).pack(anchor='w')
        
        self.file_path_var = tk.StringVar()
        self.file_entry = tk.Entry(
            file_frame, 
            textvariable=self.file_path_var,
            font=("Arial", 10),
            width=50,
            state='readonly'
        )
        self.file_entry.pack(side='left', padx=(0, 10), fill='x', expand=True)
        
        browse_btn = tk.Button(
            file_frame,
            text="Browse",
            command=self.browse_file,
            bg='#3498db',
            fg='white',
            font=("Arial", 10),
            width=10
        )
        browse_btn.pack(side='right')
        
        # Extract button
        extract_btn = tk.Button(
            self.root,
            text="Extract Highlighted Text",
            command=self.extract_highlights,
            bg='#e74c3c',
            fg='white',
            font=("Arial", 12, "bold"),
            pady=10
        )
        extract_btn.pack(pady=20)
        
        # Results display area
        results_frame = tk.Frame(self.root, bg='#f0f0f0')
        results_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        tk.Label(
            results_frame,
            text="Extracted Text:",
            font=("Arial", 12, "bold"),
            bg='#f0f0f0'
        ).pack(anchor='w')
        
        # Text area with scrollbar
        text_frame = tk.Frame(results_frame)
        text_frame.pack(fill='both', expand=True)
        
        self.results_text = tk.Text(
            text_frame,
            font=("Arial", 10),
            wrap='word',
            bg='white',
            relief='sunken',
            borderwidth=1
        )
        
        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        self.results_text.pack(side='left', fill='both', expand=True)
        
        self.results_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.results_text.yview)
        
        # Save buttons frame
        save_frame = tk.Frame(self.root, bg='#f0f0f0')
        save_frame.pack(pady=10)
        
        save_txt_btn = tk.Button(
            save_frame,
            text="Save as Text File",
            command=self.save_as_txt,
            bg='#27ae60',
            fg='white',
            font=("Arial", 10),
            width=15
        )
        save_txt_btn.pack(side='left', padx=5)
        
        save_docx_btn = tk.Button(
            save_frame,
            text="Save as Word File",
            command=self.save_as_docx,
            bg='#8e44ad',
            fg='white',
            font=("Arial", 10),
            width=15
        )
        save_docx_btn.pack(side='left', padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready...")
        status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            relief='sunken',
            anchor='w',
            bg='#ecf0f1',
            font=("Arial", 9)
        )
        status_label.pack(side='bottom', fill='x')
    
    def browse_file(self):
        """Browse and select PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            self.pdf_file = file_path
            self.file_path_var.set(file_path)
            self.status_var.set(f"File selected: {os.path.basename(file_path)}")
    
    def extract_highlights(self):
        """Extract highlighted text from PDF"""
        if not self.pdf_file:
            messagebox.showerror("Error", "Please select a PDF file first!")
            return
        
        try:
            self.status_var.set("Extracting highlighted text...")
            self.root.update()
            
            # Open PDF file
            doc = fitz.open(self.pdf_file)
            self.extracted_highlights = []
            
            # Search through all pages
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Search for annotations
                annotations = page.annots()
                
                for annot in annotations:
                    annot_type = annot.type[1] if len(annot.type) > 1 else annot.type[0]
                    
                    # Check for different highlight types
                    if annot_type in ['Highlight', 'Squiggly', 'Underline', 'StrikeOut', 'Square', 'FreeText']:
                        # Get highlighted text
                        highlighted_text = self.get_highlighted_text(page, annot)
                        
                        if highlighted_text:
                            highlight_info = {
                                'page': page_num + 1,
                                'text': highlighted_text.strip(),
                                'color': self.get_annot_color(annot),
                                'type': annot_type
                            }
                            
                            # Check for yellow color (accepting most colors now)
                            if self.is_yellow_highlight(highlight_info['color']):
                                self.extracted_highlights.append(highlight_info)
                
                # Additional: Search for highlights using alternative methods
                try:
                    # Search in drawings
                    drawings = page.get_drawings()
                    for drawing in drawings:
                        if 'fill' in drawing and drawing['fill']:
                            fill_color = drawing.get('fill')
                            if fill_color and len(fill_color) >= 3:
                                r, g, b = fill_color[0], fill_color[1], fill_color[2]
                                # If color is light (might be a highlight)
                                if (r + g + b) / 3 > 0.6:
                                    rect = drawing.get('rect')
                                    if rect:
                                        text = page.get_textbox(rect)
                                        if text and text.strip():
                                            highlight_info = {
                                                'page': page_num + 1,
                                                'text': text.strip(),
                                                'color': fill_color,
                                                'type': 'Drawing'
                                            }
                                            self.extracted_highlights.append(highlight_info)
                except:
                    pass
            
            doc.close()
            
            # Display results
            self.display_results()
            
            self.status_var.set(f"Extracted {len(self.extracted_highlights)} highlighted text(s)")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while extracting text: {str(e)}")
            self.status_var.set("Error occurred during extraction")
    
    def get_highlighted_text(self, page, annot):
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
    
    def get_annot_color(self, annot):
        """Get highlight color"""
        try:
            # Try to get color
            color = annot.colors.get("stroke", None)
            if not color:
                color = annot.colors.get("fill", None)
            
            return color if color else [1.0, 1.0, 0.0]  # Default yellow
        except:
            return [1.0, 1.0, 0.0]  # Default yellow
    
    def is_yellow_highlight(self, color):
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
    
    def display_results(self):
        """Display results in text area"""
        self.results_text.delete(1.0, tk.END)
        
        if not self.extracted_highlights:
            self.results_text.insert(tk.END, "No highlighted text found in this file.")
            return
        
        for i, highlight in enumerate(self.extracted_highlights, 1):
            self.results_text.insert(tk.END, f"[{i}] Page {highlight['page']}:\n")
            self.results_text.insert(tk.END, f"{highlight['text']}\n")
            self.results_text.insert(tk.END, "-" * 50 + "\n\n")
    
    def save_as_txt(self):
        """Save results as text file"""
        if not self.extracted_highlights:
            messagebox.showwarning("Warning", "No text to save!")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Save as Text File"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("Highlighted Text from PDF\n")
                    f.write("=" * 50 + "\n")
                    f.write(f"Source file: {os.path.basename(self.pdf_file)}\n")
                    f.write(f"Extraction date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Number of extracted texts: {len(self.extracted_highlights)}\n")
                    f.write("=" * 50 + "\n\n")
                    
                    for i, highlight in enumerate(self.extracted_highlights, 1):
                        f.write(f"[{i}] Page {highlight['page']}:\n")
                        f.write(f"{highlight['text']}\n")
                        f.write("-" * 50 + "\n\n")
                
                messagebox.showinfo("Success", f"File saved successfully at:\n{file_path}")
                self.status_var.set("Text file saved successfully")
                
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while saving file: {str(e)}")
    
    def save_as_docx(self):
        """Save results as Word file"""
        if not self.extracted_highlights:
            messagebox.showwarning("Warning", "No text to save!")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
            title="Save as Word File"
        )
        
        if file_path:
            try:
                doc = Document()
                
                # Add title
                heading = doc.add_heading('Highlighted Text from PDF', 0)
                heading.alignment = 1  # Center alignment
                
                # Add file information
                info_para = doc.add_paragraph()
                info_para.add_run(f"Source file: {os.path.basename(self.pdf_file)}\n").bold = True
                info_para.add_run(f"Extraction date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                info_para.add_run(f"Number of extracted texts: {len(self.extracted_highlights)}")
                
                doc.add_paragraph()  # Empty line
                
                # Add extracted texts
                for i, highlight in enumerate(self.extracted_highlights, 1):
                    # Number and page of text
                    header_para = doc.add_paragraph()
                    header_para.add_run(f"[{i}] Page {highlight['page']}:").bold = True
                    
                    # Highlighted text
                    text_para = doc.add_paragraph(highlight['text'])
                    text_para.style = 'Quote'
                    
                    # Separator line
                    doc.add_paragraph('_' * 50)
                
                doc.save(file_path)
                messagebox.showinfo("Success", f"File saved successfully at:\n{file_path}")
                self.status_var.set("Word file saved successfully")
                
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while saving file: {str(e)}")
    
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    """Main function"""
    app = PDFHighlightExtractor()
    app.run()


if __name__ == "__main__":
    main()
