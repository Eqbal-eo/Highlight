#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
مثال لاختبار مستخرج النصوص المحددة
Test Example for PDF Highlight Extractor
"""

from simple_extractor import extract_yellow_highlights
import os


def test_extractor():
    """اختبار البرنامج مع ملف تجريبي"""
    
    print("مثال لاختبار مستخرج النصوص المحددة")
    print("=" * 50)
    
    # مسار الملف التجريبي
    test_pdf = "test_document.pdf"
    
    if not os.path.exists(test_pdf):
        print(f"تحذير: الملف التجريبي غير موجود: {test_pdf}")
        print("يرجى وضع ملف PDF للاختبار في نفس المجلد")
        return False
    
    print(f"اختبار الملف: {test_pdf}")
    
    # تشغيل الاستخراج
    highlights = extract_yellow_highlights(test_pdf, "test_output.txt")
    
    if highlights:
        print(f"\nنجح الاختبار! تم استخراج {len(highlights)} نص محدد")
        
        # عرض أول 3 نصوص
        print("\nأول 3 نصوص مستخرجة:")
        for i, highlight in enumerate(highlights[:3], 1):
            print(f"{i}. صفحة {highlight['page']}: {highlight['text'][:100]}...")
        
        return True
    else:
        print("فشل الاختبار: لم يتم استخراج أي نصوص")
        return False


def create_sample_instructions():
    """إنشاء تعليمات لإنشاء ملف PDF تجريبي"""
    
    instructions = """
تعليمات إنشاء ملف PDF تجريبي للاختبار:

1. افتح برنامج معالج النصوص (مثل Microsoft Word)

2. اكتب بعض النصوص، مثل:
   - هذا نص عادي
   - هذا نص مهم يجب تحديده
   - معلومة مفيدة أخرى
   - نص آخر للاختبار

3. حدد بعض النصوص واستخدم أداة التحديد (Highlight) باللون الأصفر

4. احفظ الملف بصيغة PDF باسم "test_document.pdf"

5. ضع الملف في نفس مجلد البرنامج

6. شغل هذا الاختبار مرة أخرى
"""
    
    with open("create_test_pdf_instructions.txt", "w", encoding="utf-8") as f:
        f.write(instructions)
    
    print("تم إنشاء ملف التعليمات: create_test_pdf_instructions.txt")


def main():
    """الدالة الرئيسية"""
    if not test_extractor():
        print("\n" + "=" * 50)
        create_sample_instructions()


if __name__ == "__main__":
    main()
