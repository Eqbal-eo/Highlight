#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
نسخة مبسطة لاستخراج النصوص المحددة باللون الأصفر من PDF
Simple PDF Highlight Extractor
"""

import fitz  # PyMuPDF
import sys
import os
from datetime import datetime


def extract_yellow_highlights(pdf_path, output_path=None):
    """
    استخراج النصوص المحددة باللون الأصفر من ملف PDF
    
    Args:
        pdf_path (str): مسار ملف PDF
        output_path (str): مسار ملف الإخراج (اختياري)
    
    Returns:
        list: قائمة بالنصوص المستخرجة
    """
    
    if not os.path.exists(pdf_path):
        print(f"خطأ: الملف غير موجود: {pdf_path}")
        return []
    
    try:
        # فتح ملف PDF
        doc = fitz.open(pdf_path)
        extracted_highlights = []
        
        print(f"جاري معالجة ملف: {os.path.basename(pdf_path)}")
        print(f"عدد الصفحات: {len(doc)}")
        
        # البحث في كل صفحة
        for page_num in range(len(doc)):
            page = doc[page_num]
            print(f"معالجة الصفحة {page_num + 1}...")
            
            # البحث عن التحديدات (annotations)
            annotations = page.annots()
            
            for annot in annotations:
                # التحقق من نوع التحديد
                if annot.type[1] == 'Highlight':
                    # الحصول على النص المحدد
                    highlighted_text = get_highlighted_text(page, annot)
                    
                    if highlighted_text and highlighted_text.strip():
                        # الحصول على لون التحديد
                        color = get_annot_color(annot)
                        
                        # التحقق من اللون الأصفر
                        if is_yellow_highlight(color):
                            highlight_info = {
                                'page': page_num + 1,
                                'text': highlighted_text.strip(),
                                'color': color
                            }
                            extracted_highlights.append(highlight_info)
                            print(f"  تم العثور على نص محدد في الصفحة {page_num + 1}")
        
        doc.close()
        
        # طباعة النتائج
        print(f"\nتم الانتهاء! تم العثور على {len(extracted_highlights)} نص محدد باللون الأصفر")
        
        # حفظ النتائج إذا تم تحديد مسار الإخراج
        if output_path and extracted_highlights:
            save_to_file(extracted_highlights, pdf_path, output_path)
        
        # طباعة النصوص المستخرجة
        if extracted_highlights:
            print("\n" + "=" * 60)
            print("النصوص المستخرجة:")
            print("=" * 60)
            
            for i, highlight in enumerate(extracted_highlights, 1):
                print(f"\n[{i}] صفحة {highlight['page']}:")
                print("-" * 40)
                print(highlight['text'])
                print("-" * 40)
        else:
            print("لم يتم العثور على نصوص محددة باللون الأصفر في هذا الملف.")
        
        return extracted_highlights
        
    except Exception as e:
        print(f"خطأ أثناء معالجة الملف: {str(e)}")
        return []


def get_highlighted_text(page, annot):
    """استخراج النص المحدد من التعليق التوضيحي"""
    try:
        # الحصول على مستطيل التحديد
        rect = annot.rect
        
        # استخراج النص من المنطقة المحددة
        text_instances = page.get_text("dict")
        highlighted_text = ""
        
        for block in text_instances["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        span_rect = fitz.Rect(span["bbox"])
                        # التحقق من تداخل النص مع منطقة التحديد
                        if span_rect.intersects(rect):
                            highlighted_text += span["text"]
        
        # إذا لم نجد نص بالطريقة السابقة، نجرب طريقة أخرى
        if not highlighted_text.strip():
            highlighted_text = page.get_textbox(rect)
        
        return highlighted_text
        
    except Exception as e:
        print(f"خطأ في استخراج النص: {e}")
        return ""


def get_annot_color(annot):
    """الحصول على لون التحديد"""
    try:
        # محاولة الحصول على اللون
        color = annot.colors.get("stroke", None)
        if not color:
            color = annot.colors.get("fill", None)
        
        return color if color else [1.0, 1.0, 0.0]  # أصفر افتراضي
    except:
        return [1.0, 1.0, 0.0]  # أصفر افتراضي


def is_yellow_highlight(color):
    """التحقق من كون اللون أصفر"""
    if not color or len(color) < 3:
        return True  # افتراضياً نعتبره أصفر
    
    # التحقق من القيم RGB للون الأصفر
    # اللون الأصفر عادة ما يكون (1.0, 1.0, 0.0) أو قريب منه
    r, g, b = color[0], color[1], color[2]
    
    # نتحقق من أن الأحمر والأخضر مرتفعان والأزرق منخفض
    return r > 0.7 and g > 0.7 and b < 0.3


def save_to_file(highlights, pdf_path, output_path):
    """حفظ النصوص المستخرجة في ملف"""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("النصوص المحددة باللون الأصفر من PDF\n")
            f.write("=" * 60 + "\n")
            f.write(f"الملف المصدر: {os.path.basename(pdf_path)}\n")
            f.write(f"تاريخ الاستخراج: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"عدد النصوص المستخرجة: {len(highlights)}\n")
            f.write("=" * 60 + "\n\n")
            
            for i, highlight in enumerate(highlights, 1):
                f.write(f"[{i}] صفحة {highlight['page']}:\n")
                f.write("-" * 40 + "\n")
                f.write(f"{highlight['text']}\n")
                f.write("-" * 40 + "\n\n")
        
        print(f"\nتم حفظ النتائج في: {output_path}")
        
    except Exception as e:
        print(f"خطأ في حفظ الملف: {str(e)}")


def main():
    """الدالة الرئيسية"""
    print("مستخرج النصوص المحددة باللون الأصفر من PDF")
    print("=" * 50)
    
    if len(sys.argv) < 2:
        print("الاستخدام:")
        print(f"python {sys.argv[0]} <مسار_ملف_PDF> [مسار_ملف_الإخراج]")
        print("\nمثال:")
        print(f"python {sys.argv[0]} document.pdf")
        print(f"python {sys.argv[0]} document.pdf output.txt")
        return
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    # تشغيل الاستخراج
    highlights = extract_yellow_highlights(pdf_path, output_path)
    
    if highlights:
        print(f"\nتم الانتهاء بنجاح! تم استخراج {len(highlights)} نص محدد.")
    else:
        print("\nلم يتم العثور على نصوص محددة أو حدث خطأ.")


if __name__ == "__main__":
    main()
