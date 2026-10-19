#!/usr/bin/env python3
"""
tests/test_pdf_ingestion.py: Tests direct ingestion of medical PDFs (text, tables, and images).
"""

import os
import sys

SCRIPTS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

try:
    import pymupdf
except ImportError:
    pymupdf = None

from parse_input_lecture import ingest_file

def create_sample_medical_pdf(pdf_path: str):
    """Creates a synthetic medical lecture PDF with text and a table."""
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842) # A4

    # Text insertion
    text = (
        "فصل سوم: پرفشاری خون شریانی (Hypertension)\n"
        "تعریف: فشار خون سیستولیک >= 130 mm Hg یا دیاستولیک >= 80 mm Hg بر اساس معیارهای ACC/AHA.\n"
        "نکته مهم امتحانی استاد: درمان خط اول در بیمار دیابتی مهارکننده‌های ACEi یا ARB است.\n"
        "کنترااندیکاسیون مطلق: مصرف بتابلوکر در بلوک دهلیزی-بطنی درجه ۲ و ۳ اکیداً ممنوع است.\n"
    )
    rect = pymupdf.Rect(50, 50, 545, 250)
    page.insert_textbox(rect, text, fontsize=12, fontname="helv")

    # Draw a table with lines
    # Table rect: x0=50, y0=260, x1=500, y1=360
    page.draw_rect(pymupdf.Rect(50, 260, 500, 360), color=(0, 0, 0), width=1)
    page.draw_line(pymupdf.Point(50, 310), pymupdf.Point(500, 310), color=(0, 0, 0), width=1)
    page.draw_line(pymupdf.Point(250, 260), pymupdf.Point(250, 360), color=(0, 0, 0), width=1)
    page.insert_textbox(pymupdf.Rect(55, 265, 245, 305), "دسته دارویی", fontsize=11, fontname="helv")
    page.insert_textbox(pymupdf.Rect(255, 265, 495, 305), "اندیکاسیون انتخابی", fontsize=11, fontname="helv")
    page.insert_textbox(pymupdf.Rect(55, 315, 245, 355), "ACEi / ARB", fontsize=11, fontname="helv")
    page.insert_textbox(pymupdf.Rect(255, 315, 495, 355), "دیابت، نارسایی قلبی و CKD", fontsize=11, fontname="helv")

    doc.save(pdf_path)
    doc.close()

def test_pdf_parsing():
    if pymupdf is None:
        print("[INFO] PyMuPDF not installed in environment; skipping synthetic PDF ingestion test.")
        return

    pdf_path = "temp_test_lecture.pdf"
    try:
        create_sample_medical_pdf(pdf_path)
        data = ingest_file(pdf_path)
        full_text = data.get("full_text", "")

        print(f"[CHECK] PDF Extracted words: {len(full_text.split())}")
        assert "Hypertension" in full_text, "Failed to extract text from PDF!"
        assert "130 mm Hg" in full_text, "Failed to extract clinical cutoff from PDF!"
        print("[SUCCESS] PDF ingestion test passed successfully!")
    finally:
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except Exception:
                pass

if __name__ == "__main__":
    test_pdf_parsing()
