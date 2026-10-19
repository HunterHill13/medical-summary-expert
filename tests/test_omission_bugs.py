#!/usr/bin/env python3
"""
tests/test_omission_bugs.py: Tight feedback loop reproducing data loss and omission bugs:
1. Bug A: Tables inside Word (.docx) are omitted from full_text.
2. Bug B: Text inside Word Shapes / Text Boxes (w:txbxContent) is omitted.
3. Bug C: Missing numerical cutoff verification.
"""

import os
import sys
import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

SCRIPTS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from parse_input_lecture import ingest_file

def create_bug_repro_docx(file_path: str):
    """Creates a sample docx containing a table and a text box."""
    doc = docx.Document()
    doc.add_paragraph("معرفی بیماری: پلورال افیوژن اگسوداتیو در برابر ترانسوداتیو.")

    # 1. Add a clinical table with critical cutoffs
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "معیار لایت (Light Criteria)"
    table.cell(0, 1).text = "آستانه تشخیصی اگسودا"
    table.cell(1, 0).text = "نسبت پروتئین مایع پلور به سرم"
    table.cell(1, 1).text = "> 0.5"
    table.cell(2, 0).text = "نسبت LDH مایع پلور به سرم"
    table.cell(2, 1).text = "> 0.6"

    # 2. Add a paragraph with a Word Text Box (w:txbxContent)
    p_box = doc.add_paragraph("نکته طلایی زیر در تکست‌باکس آمده است:")
    # Inject w:txbxContent via OpenXML
    txbx_xml = parse_xml(r'''
        <w:r {} >
            <w:drawing>
                <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:txbx>
                                    <w:txbxContent>
                                        <w:p><w:r><w:t>نکته دامی مهم: مایع ترانسودا نیاز به بررسی بیشتر ندارد ولی اگسودا حتما نیازمند بیوپسی است.</w:t></w:r></w:p>
                                    </w:txbxContent>
                                </wps:txbx>
                            </wps:wsp>
                        </a:graphicData>
                    </a:graphic>
                </wp:inline>
            </w:drawing>
        </w:r>
    '''.format(nsdecls('w')))
    p_box._p.append(txbx_xml)

    doc.save(file_path)

def test_intake_preserves_tables_and_textboxes():
    test_file = "repro_test.docx"
    try:
        create_bug_repro_docx(test_file)
        data = ingest_file(test_file)
        full_text = data.get("full_text", "")

        # Test 1: Table text must be in full_text
        has_table_text = "معیار لایت" in full_text and "> 0.5" in full_text
        print(f"[CHECK] Table content present in full_text: {has_table_text}")

        # Test 2: Text box content must be in full_text
        has_txbx_text = "نکته دامی مهم" in full_text
        print(f"[CHECK] Textbox content present in full_text: {has_txbx_text}")

        assert has_table_text, "BUG DETECTED: Table data was OMITTED from full_text!"
        assert has_txbx_text, "BUG DETECTED: Textbox (w:txbxContent) was OMITTED from full_text!"
        print("[SUCCESS] All checks passed! Zero omission in input ingestion.")
        return True
    finally:
        if os.path.exists(test_file):
            try:
                os.remove(test_file)
            except Exception:
                pass

if __name__ == "__main__":
    test_intake_preserves_tables_and_textboxes()
