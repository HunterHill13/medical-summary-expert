#!/usr/bin/env python3
"""
build_summary_pack.py: Unified Publication-Grade Medical Summary & Revision Pack Builder.
Specifically optimized for Medical Students in Physiopathology (فیزیوپاتولوژی) & Clinical Sciences.

Produces the Triple-Pack:
1. Microsoft Word Study Document (.docx) with Persian RTL typography, Dubai font, custom callout boxes & styled tables.
2. Formatted Markdown Document (.md) with Mermaid flowcharts & structured sections.
3. Anki Spaced-Repetition Deck (.tsv & optional .apkg) with atomic cloze cards following Wozniak's 20 Rules.
"""

import os
import sys
import json
import re
import html
import argparse
import hashlib
import io
from typing import List, Dict, Any, Optional, Union

# Force UTF-8 encoding in stdout/stderr
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Teal Medical Palette Specifications (v3.9.0 Table-First Hybrid)
COLOR_PRIMARY_TEAL      = (0x00, 0x7A, 0x87)  # #007A87 - Main Headings & Table Headers
COLOR_TEAL_DARK         = (0x0B, 0x72, 0x85)  # #0B7285 - Deep Teal Accent
COLOR_SECONDARY_CYAN    = (0x08, 0x91, 0xB2)  # #0891B2 - Subheadings & Section Borders
COLOR_TEXT_CHARCOAL     = (0x1F, 0x29, 0x37)  # #1F2937 - Main Body Text (Dark Slate)
COLOR_MUTED_GRAY        = (0x64, 0x74, 0x8B)  # #64748B - Metadata & Subtitles
COLOR_CARD_BORDER       = "CBD5E1"            # Table Border Crisp Light Gray
COLOR_CARD_BG           = "F8FAFC"            # Soft Slate BG
COLOR_AMBER_BORDER      = "D97706"            # Amber/Orange Border for High-Yield Pearls
COLOR_AMBER_BG          = "FFFBEB"            # Soft Yellow/Cream Fill for Pearls
COLOR_TABLE_HEADER_BG   = "007A87"            # Teal Medical Header Background (#007A87)
COLOR_TABLE_HEADER_TXT  = "FFFFFF"            # White Table Header Text
COLOR_TABLE_ALT_ROW     = "F0FDFA"            # Subtle Mint/Teal Zebra Row Fill (#F0FDFA)
COLOR_RUBY_RED          = (0xB9, 0x1C, 0x1C)  # #B91C1C - Red Trap Badges (🔴 تله امتحانی)
COLOR_EMERALD_GREEN     = (0x04, 0x78, 0x57)  # #047857 - Gold Standard Badges (💎 استاندارد طلایی)
COLOR_AMBER_ORANGE      = (0xC2, 0x41, 0x0C)  # #C2410C - Common Error Badges (⚠️ خطای رایج)
COLOR_BLUE_BORDER       = "38BDF8"            # Sky Blue Border
COLOR_BLUE_BG           = "F0F9FF"            # Sky Blue Fill
COLOR_SEMIOLOGY_TEAL    = "0D9488"            # Teal for Clinical Semiology
COLOR_SEMIOLOGY_BG      = "F0FDFA"            # Light Teal Fill
COLOR_CHECKLIST_NAVY    = "0369A1"            # Navy for Morning-of-Exam Capsule
COLOR_CHECKLIST_BG      = "FEFDE8"            # Warm Amber Tint for Morning Checklist

# Legacy compatibility aliases
COLOR_PRIMARY_BURGUNDY  = COLOR_PRIMARY_TEAL
COLOR_SECONDARY_TEAL    = COLOR_SECONDARY_CYAN

FONT_DUBAI = "Dubai"

# ---------------------------------------------------------------------------
# 6 Context-Aware Table Archetypes (v3.9.0 Table-First Hybrid)
# ---------------------------------------------------------------------------
TABLE_ARCHETYPES = {
    "pathophysiology": {
        "name": "اصول پاتوفیزیولوژی و اتیولوژی",
        "headers": ["مکانیسم پاتوفیزیولوژیک", "علل و اتیولوژی‌های شایع", "نمونه‌های بالینی و نکات امتحانی"]
    },
    "anatomy": {
        "name": "آناتومی بالینی و مجاورات توده‌ای",
        "headers": ["عنصر آناتومیک", "موقعیت و ساختار / اثر مکانیکی", "پیامد بالینی، علائم تیپیک و نکات جراحی"]
    },
    "pharmacology": {
        "name": "کلاس‌های بیوشیمیایی، هورمون‌ها و داروشناسی",
        "headers": ["کلاس بیوشیمیایی / دارویی", "نمونه‌های شاخص و لیگاندها", "ویژگی‌های اختصاصی، فارماکوکینتیک و نکات امتحانی"]
    },
    "diagnostic": {
        "name": "معیارهای تشخیصی، آزمایشگاهی و کات‌آف‌ها",
        "headers": ["تست تشخیصی / روش سنجش", "طبیعی (Normal)", "پیش‌دیابت / بینابینی", "حالت بیماری / آستانه تشخیصی"]
    },
    "differential": {
        "name": "ماتریس جامع تشخیص افتراقی (Illness Script)",
        "headers": ["ویژگی افتراقی", "بیماری الف", "بیماری ب", "بیماری ج"]
    },
    "multiorgan": {
        "name": "پاتوژنز چندارگانی و تست‌های دینامیک",
        "headers": ["ارگان درگیر / گام تست", "نقص پاتوفیزیولوژیک / اندیکاسیون", "پیامد بالینی / پروتکل محرک-مهار", "درمان دارویی / تفسیر تشخیصی"]
    }
}

# ---------------------------------------------------------------------------
# BiDi & OpenXML Typography Helpers for Word (.docx) & TeX Sanitization
# ---------------------------------------------------------------------------

def sanitize_medical_text_symbols(text: str) -> str:
    """
    Sanitizes TeX math delimiters, arrows, comparison operators, and common chemical
    formulas into clean Unicode representation, eliminating raw LaTeX token leaks and
    BiDi character collision in Persian RTL text.
    """
    if not text:
        return ""
    s = str(text)

    # 1. Chemical formulas & Hydroxyapatite isolation
    s = re.sub(
        r'(\\text\{Ca\}\_\{10\}\(\\text\{PO\}\_4\)\_6\(\\text\{OH\}\)\_2|\$?\\text\{Ca\}_?\{?10\}?\(?.*?\)?\$?)',
        'Ca₁₀(PO₄)₆(OH)₂',
        s
    )
    s = s.replace(r'\text{Ca}_{10}(\text{PO}_4)_6(\text{OH})_2', 'Ca₁₀(PO₄)₆(OH)₂')
    s = s.replace(r'Ca_{10}(PO_4)_6(OH)_2', 'Ca₁₀(PO₄)₆(OH)₂')

    # Common chemical ions and hormone notations
    s = s.replace(r'$E_1$', 'E₁').replace(r'$E_2$', 'E₂').replace(r'$E_3$', 'E₃')
    s = s.replace(r'E_1', 'E₁').replace(r'E_2', 'E₂').replace(r'E_3', 'E₃')
    s = s.replace(r'T_3', 'T₃').replace(r'T_4', 'T₄').replace(r'FT_4', 'FT₄').replace(r'FT_3', 'FT₃')
    s = s.replace(r'$T_3$', 'T₃').replace(r'$T_4$', 'T₄').replace(r'$FT_4$', 'FT₄').replace(r'$FT_3$', 'FT₃')
    s = s.replace(r'Ca^{2+}', 'Ca²⁺').replace(r'Ca^2+', 'Ca²⁺').replace(r'$Ca^{2+}$', 'Ca²⁺')
    s = s.replace(r'Na^+', 'Na⁺').replace(r'K^+', 'K⁺').replace(r'H^+', 'H⁺').replace(r'Mg^{2+}', 'Mg²⁺')
    s = s.replace(r'HCO_3^-', 'HCO₃⁻').replace(r'PO_4^{3-}', 'PO₄³⁻')

    # 2. Arrows & Relational Operators
    s = s.replace(r'$\rightarrow$', '←').replace(r'\rightarrow', '←')
    s = s.replace(r'$\leftarrow$', '→').replace(r'\leftarrow', '→')
    s = s.replace(r'$\leftrightarrow$', '↔').replace(r'\leftrightarrow', '↔')
    s = s.replace(r'$\ge$', '≥').replace(r'\ge', '≥')
    s = s.replace(r'$\le$', '≤').replace(r'\le', '≤')
    s = s.replace(r'$\pm$', '±').replace(r'\pm', '±')
    s = s.replace(r'$\times$', '×').replace(r'\times', '×')
    s = s.replace(r'$\approx$', '≈').replace(r'\approx', '≈')
    s = s.replace(r'$\neq$', '≠').replace(r'\neq', '≠')

    # 3. TeX wrapper cleanups
    s = re.sub(r'\\text\{([^}]+)\}', r'\1', s)
    s = re.sub(r'\\mathrm\{([^}]+)\}', r'\1', s)
    s = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', s)

    # 4. Remove unescaped dollar signs wrapping plain text
    s = re.sub(r'\$([^\$]+)\$', r'\1', s)

    return s


def format_inline_markdown(text: str) -> str:
    """
    Sanitizes inline Markdown and TeX symbols into standard HTML tags and Unicode.
    Converts **bold** to <strong>, `code` to <code class="med-code">, and TeX math to Unicode.
    """
    if not text:
        return ""
    s = sanitize_medical_text_symbols(str(text))

    # 1. Bold: **text** -> <strong>text</strong>
    s = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', s)

    # 2. Inline code: `text` -> <code class="med-code">\1</code>
    s = re.sub(r'`(.*?)`', r'<code class="med-code">\1</code>', s)

    # 3. Wrap chemical formulas in <span dir="ltr">
    s = re.sub(
        r'(Ca₁₀\(PO₄\)₆\(OH\)₂)',
        r'<span dir="ltr" class="chem-formula">\1</span>',
        s
    )

    return s


def format_cell_html(text: str) -> str:
    """
    Formats a table cell's multi-line text with bullets, bolding, and colored badge spans
    matching the v3.9.0 Table-First Hybrid specification.
    """
    if not text:
        return ""
    lines = str(text).splitlines()
    formatted_lines = []
    for line in lines:
        l = line.strip()
        if not l:
            continue
        # In-cell badge replacements with high-contrast colored spans
        l = re.sub(r'(🔴\s*تله\s*(?:بزرگ\s*)?امتحانی(?:\s*:\s*)?)', r'<span class="badge-trap">\1</span>', l)
        l = re.sub(r'(⚠️\s*خطای\s*(?:فاحش\s*|رایج\s*)(?:پزشکی\s*)?(?:\s*:\s*)?)', r'<span class="badge-error">\1</span>', l)
        l = re.sub(r'(💎\s*استاندارد\s*طلایی(?:\s*:\s*)?)', r'<span class="badge-gold">\1</span>', l)
        l = re.sub(r'(📌\s*نکته(?:\s*دقیق\s*پاتولوژی|\s*مهم|\s*امتحانی)?(?:\s*:\s*)?)', r'<span class="badge-note">\1</span>', l)
        l = re.sub(r'(🏆\s*شایع‌ترین(?:\s*فرم|\s*علت|\s*نمونه)?(?:\s*:\s*)?)', r'<span class="badge-freq">\1</span>', l)
        l = re.sub(r'(•\s*)مهم\s+(?!<)', r'\1<span class="badge-trap">مهم: </span>', l)
        l = format_inline_markdown(l)
        if l.startswith("• "):
            formatted_lines.append(f'<div class="cell-bullet">{l}</div>')
        else:
            formatted_lines.append(f'<div class="cell-line">{l}</div>')
    return "".join(formatted_lines)


def render_formula_html(raw_f_code: str, title: str = "") -> str:
    """
    Renders high-contrast, bilingual diagnostic formula cards with fractions,
    clean LTR alignment, and explanatory variable legends.
    """
    clean = str(raw_f_code).strip()
    clean_sanitized = sanitize_medical_text_symbols(clean)

    if "CCCR" in clean or "Clearance" in clean:
        badge = title or "🧮 فرمول تشخیصی نسبت کلیرانس کلسیم به کلیرانس کراتینین (CCCR)"
        return f"""
        <div class="formula-card">
            <div class="formula-badge">{badge}</div>
            <div class="formula-container" dir="ltr">
                <span class="formula-lhs">CCCR</span>
                <span class="formula-eq-sign">=</span>
                <div class="formula-fraction">
                    <div class="formula-num">U<sub>Ca</sub> × S<sub>Cr</sub></div>
                    <div class="formula-denom">S<sub>Ca</sub> × U<sub>Cr</sub></div>
                </div>
            </div>
            <div class="formula-caption">
                <span><strong>U<sub>Ca</sub>:</strong> کلسیم ادرار ۲۴ ساعته</span> | 
                <span><strong>S<sub>Cr</sub>:</strong> کراتینین سرم</span> | 
                <span><strong>S<sub>Ca</sub>:</strong> کلسیم تام سرم</span> | 
                <span><strong>U<sub>Cr</sub>:</strong> کراتینین ادرار ۲۴ ساعته</span>
            </div>
        </div>
        """
    elif "Corrected" in clean or "آلبومین" in clean or "Albumin" in clean:
        badge = title or "🧮 فرمول تصحیح کلسیم سرم بر اساس سطح آلبومین (Albumin Correction)"
        return f"""
        <div class="formula-card">
            <div class="formula-badge">{badge}</div>
            <div class="formula-container" dir="ltr">
                <span class="formula-lhs">Corrected Ca (mg/dL)</span>
                <span class="formula-eq-sign">=</span>
                <span class="formula-rhs">Measured Ca + 0.8 × (4.0 − Serum Albumin)</span>
            </div>
            <div class="formula-caption">
                به ازای هر ۱ g/dL افت آلبومین سرم کمتر از ۴.۰ g/dL، مقدار ۰.۸ mg/dL به کلسیم اندازه‌گیری‌شده اضافه می‌شود.
            </div>
        </div>
        """
    elif "Anion Gap" in clean or "آنیون گپ" in clean:
        badge = title or "🧮 فرمول محاسبه آنیون گپ سرم (Serum Anion Gap)"
        return f"""
        <div class="formula-card">
            <div class="formula-badge">{badge}</div>
            <div class="formula-container" dir="ltr">
                <span class="formula-lhs">Anion Gap (mEq/L)</span>
                <span class="formula-eq-sign">=</span>
                <span class="formula-rhs">Na<sup>+</sup> − (Cl<sup>−</sup> + HCO<sub>3</sub><sup>−</sup>)</span>
            </div>
            <div class="formula-caption">
                مقدار نرمال: ۱۰ تا ۱۲ mEq/L (آستانه غیرطبیعی: بالای ۱۲).
            </div>
        </div>
        """
    else:
        badge = title or "🧮 فرمول محاسباتی بالینی"
        return f"""
        <div class="formula-card">
            <div class="formula-badge">{badge}</div>
            <div class="formula-container" dir="ltr">
                <span class="formula-rhs">{clean_sanitized}</span>
            </div>
        </div>
        """


def sanitize_anki_html(text: str) -> str:
    """
    Escapes standalone comparison operators (<, >) into HTML entities (&lt;, &gt;)
    to prevent GenAnki UserWarnings and broken card rendering,
    while strictly preserving valid HTML tags (<br>, <b>, <strong>, <span>, etc.)
    and Cloze deletions ({{c1::...}}).
    """
    if not text:
        return ""
    valid_tag_pattern = re.compile(r'(</?(?:b|strong|i|em|u|br|span|div|hr|sub|sup)(?:\s+[^<>]*)?>)', re.IGNORECASE)
    parts = valid_tag_pattern.split(str(text))
    out = []
    for part in parts:
        if valid_tag_pattern.match(part):
            out.append(part)
        else:
            p_esc = part.replace('&lt;', '<').replace('&gt;', '>')
            p_esc = p_esc.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            p_esc = re.sub(r'&amp;(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);', r'&\1;', p_esc)
            out.append(p_esc)
    return "".join(out)


def apply_table_xml_safety(table, is_multi_row_tabular: bool = True):
    """
    Applies publication-grade XML properties to Word tables:
    1. Sets w:bidiVisual on tblPr so table columns read right-to-left.
    2. Sets w:cantSplit on EVERY row to prevent awkward page boundary splits.
    3. If is_multi_row_tabular is True, sets w:tblHeader on row 0 so the table header
       repeats cleanly across page breaks on multi-page tables.
    """
    tblPr = table._tbl.tblPr
    if not tblPr.xpath('./w:bidiVisual'):
        tblPr.append(parse_xml(r'<w:bidiVisual ' + nsdecls('w') + ' w:val="1"/>'))

    for idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        if not trPr.xpath('./w:cantSplit'):
            trPr.append(parse_xml(r'<w:cantSplit ' + nsdecls('w') + '/>'))
        if is_multi_row_tabular and idx == 0 and len(table.rows) > 1:
            if not trPr.xpath('./w:tblHeader'):
                trPr.append(parse_xml(r'<w:tblHeader ' + nsdecls('w') + '/>'))


def safe_write_pdf(output_path: str, write_fn_or_bytes) -> str:
    """
    Safely writes PDF output with Windows Exclusive Lock fallback handling.
    If output_path is locked by Acrobat Reader or Edge browser, writes to an _Updated.pdf
    fallback path instead of crashing with PermissionError [Errno 13].
    """
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    target_path = output_path
    for attempt in range(10):
        try:
            if callable(write_fn_or_bytes):
                with open(target_path, "wb") as f_out:
                    write_fn_or_bytes(f_out)
            else:
                with open(target_path, "wb") as f_out:
                    f_out.write(write_fn_or_bytes)
            if attempt > 0:
                print(f"[WARN] Target PDF was locked. Saved successfully to fallback: {target_path}")
            return target_path
        except PermissionError:
            base, ext = os.path.splitext(output_path)
            target_path = f"{base}_Updated_{attempt+1}{ext}" if attempt > 0 else f"{base}_Updated{ext}"
            print(f"[WARN] Target PDF '{output_path}' is locked by another application. Trying fallback: {target_path}")
    return output_path


def add_formula_card(doc, raw_f_code: str, title: str = ""):
    """Renders a styled formula card in Word (.docx) with LTR typography and variable explanation."""
    if not raw_f_code:
        return
    badge = title or "🧮 فرمول تشخیصی و محاسبات بالینی"
    table = doc.add_table(rows=2, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    apply_table_xml_safety(table, is_multi_row_tabular=False)

    # Header
    h_cell = table.cell(0, 0)
    set_cell_margins(h_cell, top=120, bottom=120, left=160, right=160)
    set_cell_shading(h_cell, "EBF5FB")
    h_cell._tc.get_or_add_tcPr().append(parse_xml(
        r'<w:tcBorders {}><w:top w:val="single" w:sz="10" w:color="2980B9"/><w:left w:val="single" w:sz="10" w:color="2980B9"/><w:right w:val="single" w:sz="10" w:color="2980B9"/><w:bottom w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
    ))
    p_h = h_cell.paragraphs[0]
    set_p_rtl(p_h, space_before_pt=0, space_after_pt=0)
    add_r(p_h, badge, font_name=FONT_DUBAI, size_pt=10.5, bold=True,
          color_rgb=(0x1A, 0x52, 0x76), is_rtl=True)

    # Body
    b_cell = table.cell(1, 0)
    set_cell_margins(b_cell, top=140, bottom=140, left=160, right=160)
    set_cell_shading(b_cell, "FBFCFD")
    b_cell._tc.get_or_add_tcPr().append(parse_xml(
        r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="single" w:sz="10" w:color="2980B9"/><w:right w:val="single" w:sz="10" w:color="2980B9"/><w:bottom w:val="single" w:sz="10" w:color="2980B9"/></w:tcBorders>'.format(nsdecls('w'))
    ))
    p_f = b_cell.paragraphs[0]
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f.paragraph_format.space_before = Pt(4)
    p_f.paragraph_format.space_after = Pt(4)

    clean_formula = sanitize_medical_text_symbols(raw_f_code)
    _add_single_r(p_f, clean_formula, font_name="Cambria Math", size_pt=11.5, bold=True,
                  color_rgb=(0x1B, 0x4F, 0x72), is_rtl=False)

    p_after = doc.add_paragraph()
    set_p_rtl(p_after, space_before_pt=0, space_after_pt=6)


def _add_single_r(p, text: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
                  bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL,
                  is_rtl: bool = True):
    """Low-level run builder with explicit Complex Script settings for Word."""
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(r'<w:rFonts {} w:ascii="{}" w:hAnsi="{}" w:cs="{}"/>'.format(
            nsdecls('w'), font_name, font_name, font_name))
        rPr.append(rFonts)
    else:
        rFonts.set(qn('w:cs'), font_name)
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)

    rtl_elem = rPr.find(qn('w:rtl'))
    if is_rtl:
        if rtl_elem is None:
            rPr.append(parse_xml(r'<w:rtl {} w:val="1"/>'.format(nsdecls('w'))))
        else:
            rtl_elem.set(qn('w:val'), '1')
    else:
        if rtl_elem is None:
            rPr.append(parse_xml(r'<w:rtl {} w:val="0"/>'.format(nsdecls('w'))))
        else:
            rtl_elem.set(qn('w:val'), '0')

    # Explicit Complex Script sizing for Word
    sz_half_pts = str(int(round(size_pt * 2)))
    szCs_elem = rPr.find(qn('w:szCs'))
    if szCs_elem is None:
        rPr.append(parse_xml(r'<w:szCs {} w:val="{}"/>'.format(nsdecls('w'), sz_half_pts)))
    else:
        szCs_elem.set(qn('w:val'), sz_half_pts)

    if bold:
        bCs_elem = rPr.find(qn('w:bCs'))
        if bCs_elem is None:
            rPr.append(parse_xml(r'<w:bCs {}/>'.format(nsdecls('w'))))
    if italic:
        iCs_elem = rPr.find(qn('w:iCs'))
        if iCs_elem is None:
            rPr.append(parse_xml(r'<w:iCs {}/>'.format(nsdecls('w'))))

    return run


def add_r(p, text: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
          bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL,
          is_rtl: bool = True):
    """Adds a run or parses inline markdown **bold** and sanitizes TeX symbols into Word runs."""
    clean_text = sanitize_medical_text_symbols(str(text))
    if "**" in clean_text:
        parts = re.split(r'(\*\*.*?\*\*)', clean_text)
        last_run = None
        for part in parts:
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) >= 4:
                last_run = _add_single_r(p, part[2:-2], font_name=font_name, size_pt=size_pt,
                                         bold=True, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl)
            else:
                last_run = _add_single_r(p, part, font_name=font_name, size_pt=size_pt,
                                         bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl)
        return last_run
    return _add_single_r(p, clean_text, font_name=font_name, size_pt=size_pt,
                         bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl)


def add_bidi_text(p, text: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
                  bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL):
    """
    Renders mixed Persian/English scientific text with strict BiDi isolation.
    Preserves arrows (→, ⇌), chemical symbols (Ca²⁺, HCO₃⁻), inequalities (≥, ≤, ±)
    and Latin abbreviations without reversing English parentheses.
    Sanitizes raw LaTeX/TeX tokens to Unicode, isolates chemical formulas,
    and converts Markdown **bold** segments into true bold runs (zero raw asterisk leaks).
    """
    if not text:
        return

    text = sanitize_medical_text_symbols(text)

    # Handle explicit line breaks
    lines = text.split('\n')
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break()
        # Parse inline markdown **bold**
        parts = re.split(r'(\*\*.*?\*\*)', line)
        for part in parts:
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) >= 4:
                inner_text = part[2:-2]
                _render_bidi_line(p, inner_text, font_name=font_name, size_pt=size_pt,
                                  bold=True, italic=italic, color_rgb=color_rgb)
            else:
                _render_bidi_line(p, part, font_name=font_name, size_pt=size_pt,
                                  bold=bold, italic=italic, color_rgb=color_rgb)


def _render_bidi_line(p, text_line: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
                      bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL):
    if not text_line:
        return
    if not any('\u0600' <= c <= '\u06FF' for c in text_line):
        add_r(p, text_line, font_name=font_name, size_pt=size_pt, bold=bold, italic=italic,
              color_rgb=color_rgb, is_rtl=False)
        return

    sci_chars = r'A-Za-z0-9_\-\./%α-ωΑ-Ω\+\^±≥≤≈≠×°µ·²³⁺⁻⁰¹⁴⁵⁶⁷⁸⁹⁼⁽⁾ⁿ₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎'
    pattern = (
        rf'(\d+(?:[,\.]\d+)?\([A-Za-z0-9_\-\s,\./%α-ωΑ-Ω\+\^±≥≤≈≠×°µ·²³⁺⁻⁰¹⁴⁵⁶⁷⁸⁹⁼⁽⁾ⁿ₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎]+\)[{sci_chars}]*'
        rf'|\([{sci_chars}\s,]+\)'
        rf'|\[[{sci_chars}\s,]+\]'
        rf'|\([^\)]*[\d\u06F0-\u06F9]+\s*[\+\-×÷=/\^±≥≤≈≠]\s*[^\)]*\)'
        rf'|\[[^\]]*[\d\u06F0-\u06F9]+\s*[\+\-×÷=/\^±≥≤≈≠]\s*[^\]]*\]'
        rf'|(?:<=|>=|≥|≤|≈|≠|[<>=])\s*\d+(?:\.\d+)?'
        rf'|[{sci_chars}]+(?:\s*(?:→|->|⇌|<->|⇄|<=|>=|≥|≤|≈|≠|<|>|=)\s*[{sci_chars}]+)+'
        rf'|[{sci_chars}]{{2,}}'
        rf'|[→⇌⇄≥≤≈≠×°µ·])'
    )
    tokens = re.split(pattern, text_line)
    for tok in tokens:
        if not tok:
            continue
        has_persian = any('\u0600' <= c <= '\u06FF' for c in tok)
        is_math_expr = bool(
            re.search(r'[\d\u06F0-\u06F9]\s*[\+\-×÷=/\^±≥≤≈≠]\s*', tok)
            and (tok.strip().startswith('(') or tok.strip().startswith('['))
        )
        if (not has_persian and (re.search(r'[A-Za-z0-9→⇌⇄<>=≥≤≈≠×°µ·]', tok) or tok.startswith('('))) or is_math_expr:
            add_r(p, " " + tok.strip() + " ", font_name=font_name, size_pt=size_pt,
                  bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=False)
        else:
            add_r(p, tok, font_name=font_name, size_pt=size_pt,
                  bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=True)


def set_p_rtl(p, space_before_pt: float = 0, space_after_pt: float = 3.5, line_spacing: float = 1.15):
    """Enforces RTL paragraph direction, alignment, and Persian typesetting spacing."""
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(space_before_pt)
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.line_spacing = line_spacing
    pPr = p._p.get_or_add_pPr()
    bidi = pPr.find(qn('w:bidi'))
    if bidi is None:
        pPr.append(parse_xml(r'<w:bidi {} w:val="1"/>'.format(nsdecls('w'))))
    else:
        bidi.set(qn('w:val'), '1')


def set_cell_margins(cell, top: int = 120, bottom: int = 120, left: int = 160, right: int = 160):
    """Sets internal padding for table cells (in twips, 20 twips = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(r'''
        <w:tcMar {} >
            <w:top w:w="{}" w:type="dxa"/>
            <w:bottom w:w="{}" w:type="dxa"/>
            <w:left w:w="{}" w:type="dxa"/>
            <w:right w:w="{}" w:type="dxa"/>
        </w:tcMar>
    '''.format(nsdecls('w'), top, bottom, left, right))
    tcPr.append(tcMar)


def set_cell_shading(cell, fill_hex: str):
    """Applies background shading to a cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), fill_hex))
    tcPr.append(shd)


def set_cell_borders(cell, top: str = None, bottom: str = None, left: str = None, right: str = None,
                     color: str = "CCCCCC", sz: str = "4"):
    """Configures specific borders on a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders_xml = [f'<w:tcBorders {nsdecls("w")}>']
    for side, active in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if active:
            borders_xml.append(f'<{active} w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>')
        else:
            borders_xml.append(f'<{side} w:val="none"/>')
    borders_xml.append('</w:tcBorders>')
    tcPr.append(parse_xml("".join(borders_xml)))


# ---------------------------------------------------------------------------
# High-Yield Callout Boxes & Tables
# ---------------------------------------------------------------------------

def add_callout_box(doc, title: str, content_items: List[str],
                    border_color: str, header_bg: str, body_bg: str, icon: str = "💡"):
    """Renders a 2-row styled callout box in Word with header bar and pure white/tinted body."""
    if not content_items:
        return
    valid_items = [str(x).strip() for x in content_items if x and str(x).strip()]
    if not valid_items:
        return
    table = doc.add_table(rows=2, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    apply_table_xml_safety(table, is_multi_row_tabular=False)

    # Header Row
    header_cell = table.cell(0, 0)
    set_cell_margins(header_cell, top=140, bottom=140, left=180, right=180)
    set_cell_shading(header_cell, header_bg)
    header_cell._tc.get_or_add_tcPr().append(parse_xml(
        r'<w:tcBorders {}><w:top w:val="single" w:sz="12" w:color="{}"/><w:left w:val="single" w:sz="12" w:color="{}"/><w:right w:val="single" w:sz="12" w:color="{}"/><w:bottom w:val="none"/></w:tcBorders>'.format(
            nsdecls('w'), border_color, border_color, border_color)
    ))
    p_head = header_cell.paragraphs[0]
    set_p_rtl(p_head, space_before_pt=0, space_after_pt=0)
    add_r(p_head, f"{icon} {title}", font_name=FONT_DUBAI, size_pt=11.0, bold=True,
          color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)

    # Body Row
    body_cell = table.cell(1, 0)
    set_cell_margins(body_cell, top=160, bottom=160, left=180, right=180)
    set_cell_shading(body_cell, body_bg)
    body_cell._tc.get_or_add_tcPr().append(parse_xml(
        r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="single" w:sz="12" w:color="{}"/><w:right w:val="single" w:sz="12" w:color="{}"/><w:bottom w:val="single" w:sz="12" w:color="{}"/></w:tcBorders>'.format(
            nsdecls('w'), border_color, border_color, border_color)
    ))

    p_first = body_cell.paragraphs[0]
    for idx, item in enumerate(valid_items):
        p = p_first if idx == 0 else body_cell.add_paragraph()
        set_p_rtl(p, space_before_pt=2, space_after_pt=3.5)
        # Bullet indicator
        add_r(p, "• ", font_name=FONT_DUBAI, size_pt=10.5, bold=True,
              color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
        add_bidi_text(p, item, font_name=FONT_DUBAI, size_pt=10.5, bold=False,
                      color_rgb=COLOR_TEXT_CHARCOAL)

    # Space after table
    p_after = doc.add_paragraph()
    set_p_rtl(p_after, space_before_pt=0, space_after_pt=6)


def add_thematic_table(doc, headers: List[str], rows: List[List[str]], title: Optional[str] = None):
    """
    Renders an exam-grade Teal Medical table matching the v3.9.0 Table-First Hybrid specification:
    - #007A87 header with bold white text
    - Repeating header on page break (w:tblHeader)
    - Prevention of mid-row page splits (w:cantSplit)
    - #F0FDFA zebra striping on alternating rows
    - In-cell badge recognition and colored typography
    """
    if not headers or not rows:
        return
    valid_rows = [r for r in rows if r and any(str(c).strip() for c in r)]
    if not valid_rows:
        return

    if title:
        p_title = doc.add_paragraph()
        set_p_rtl(p_title, space_before_pt=10, space_after_pt=4)
        add_r(p_title, f"■ {title}", font_name=FONT_DUBAI, size_pt=12.0, bold=True,
              color_rgb=COLOR_PRIMARY_TEAL, is_rtl=True)

    table = doc.add_table(rows=len(valid_rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    apply_table_xml_safety(table, is_multi_row_tabular=True)

    # Header Row
    trPr0 = table.rows[0]._tr.get_or_add_trPr()
    trPr0.append(parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w'))))
    trPr0.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))

    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        set_cell_margins(cell, top=130, bottom=130, left=130, right=130)
        set_cell_shading(cell, COLOR_TABLE_HEADER_BG)  # #007A87
        cell._tc.get_or_add_tcPr().append(parse_xml(
            r'<w:tcBorders {}><w:top w:val="single" w:sz="6" w:color="007A87"/><w:bottom w:val="single" w:sz="12" w:color="007A87"/><w:left w:val="single" w:sz="4" w:color="005C66"/><w:right w:val="single" w:sz="4" w:color="005C66"/></w:tcBorders>'.format(nsdecls('w'))
        ))
        p = cell.paragraphs[0]
        set_p_rtl(p, space_before_pt=0, space_after_pt=0)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        add_r(p, str(h_text), font_name=FONT_DUBAI, size_pt=9.5, bold=True,
              color_rgb=(0xFF, 0xFF, 0xFF), is_rtl=True)

    # Data Rows
    for row_idx, r_data in enumerate(valid_rows):
        trPr = table.rows[row_idx + 1]._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))

        bg_fill = COLOR_TABLE_ALT_ROW if row_idx % 2 == 1 else "FFFFFF"
        safe_row = r_data[:len(headers)] if isinstance(r_data, (list, tuple)) else [str(r_data)]
        for col_idx, cell_value in enumerate(safe_row):
            cell = table.cell(row_idx + 1, col_idx)
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            if bg_fill != "FFFFFF":
                set_cell_shading(cell, bg_fill)
            cell._tc.get_or_add_tcPr().append(parse_xml(
                r'<w:tcBorders {}><w:top w:val="none"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="E2E8F0"/><w:right w:val="single" w:sz="4" w:color="E2E8F0"/></w:tcBorders>'.format(nsdecls('w'))
            ))

            lines = str(cell_value).splitlines()
            p_first = cell.paragraphs[0]
            for l_idx, line in enumerate(lines):
                p = p_first if l_idx == 0 else cell.add_paragraph()
                set_p_rtl(p, space_before_pt=1, space_after_pt=2)
                l = line.strip()
                if not l:
                    continue
                # In-cell badge recognition
                badge_color = None
                if any(k in l for k in ["تله امتحانی", "🔴", "تله بزرگ"]):
                    badge_color = COLOR_RUBY_RED
                elif any(k in l for k in ["خطای رایج", "خطای فاحش", "⚠️"]):
                    badge_color = COLOR_AMBER_ORANGE
                elif any(k in l for k in ["استاندارد طلایی", "💎"]):
                    badge_color = COLOR_EMERALD_GREEN
                elif any(k in l for k in ["نکته", "مهم", "📌"]):
                    badge_color = COLOR_PRIMARY_TEAL

                is_anchor = (col_idx == 0)
                if badge_color:
                    add_bidi_text(p, l, font_name=FONT_DUBAI, size_pt=9.0,
                                  bold=True, color_rgb=badge_color)
                else:
                    add_bidi_text(p, l, font_name=FONT_DUBAI, size_pt=9.0,
                                  bold=is_anchor, color_rgb=COLOR_PRIMARY_TEAL if is_anchor else COLOR_TEXT_CHARCOAL)

    # Space after table
    p_after = doc.add_paragraph()
    set_p_rtl(p_after, space_before_pt=0, space_after_pt=6)


def add_comparison_table(doc, title: str, headers: List[str], rows: List[List[str]]):
    """Renders a reference-grade Persian comparison table with styled header and alternating rows."""
    if not headers or not rows:
        return
    valid_rows = [r for r in rows if r and any(str(c).strip() for c in r)]
    if not valid_rows:
        return

    p_title = doc.add_paragraph()
    set_p_rtl(p_title, space_before_pt=8, space_after_pt=4)
    add_r(p_title, f"📊 {title}", font_name=FONT_DUBAI, size_pt=11.5, bold=True,
          color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

    table = doc.add_table(rows=len(valid_rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    apply_table_xml_safety(table, is_multi_row_tabular=True)

    # Header Row
    for col_idx, h_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        set_cell_margins(cell, top=140, bottom=140, left=140, right=140)
        set_cell_shading(cell, COLOR_TABLE_HEADER_BG)
        cell._tc.get_or_add_tcPr().append(parse_xml(
            r'<w:tcBorders {}><w:top w:val="single" w:sz="6" w:color="78281F"/><w:bottom w:val="single" w:sz="12" w:color="78281F"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
        ))
        p = cell.paragraphs[0]
        set_p_rtl(p, space_before_pt=0, space_after_pt=0)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        add_r(p, h_text, font_name=FONT_DUBAI, size_pt=10.0, bold=True,
              color_rgb=(0xFF, 0xFF, 0xFF), is_rtl=True)

    # Data Rows
    for row_idx, r_data in enumerate(valid_rows):
        bg_fill = COLOR_TABLE_ALT_ROW if row_idx % 2 == 1 else "FFFFFF"
        # Slice r_data to match headers length to prevent IndexError on column overflow
        safe_row = r_data[:len(headers)] if isinstance(r_data, (list, tuple)) else [str(r_data)]
        for col_idx, cell_value in enumerate(safe_row):
            cell = table.cell(row_idx + 1, col_idx)
            set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
            if bg_fill != "FFFFFF":
                set_cell_shading(cell, bg_fill)
            cell._tc.get_or_add_tcPr().append(parse_xml(
                r'<w:tcBorders {}><w:top w:val="none"/><w:bottom w:val="single" w:sz="4" w:color="E0E0E0"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
            ))
            p = cell.paragraphs[0]
            set_p_rtl(p, space_before_pt=1, space_after_pt=2)
            # Make the first column slightly bolder as the condition/drug anchor
            is_anchor = (col_idx == 0)
            add_bidi_text(p, str(cell_value), font_name=FONT_DUBAI, size_pt=9.5,
                          bold=is_anchor, color_rgb=COLOR_PRIMARY_BURGUNDY if is_anchor else COLOR_TEXT_CHARCOAL)

    # Space after table
    p_after = doc.add_paragraph()
    set_p_rtl(p_after, space_before_pt=0, space_after_pt=8)


def add_figure_block(doc, image_path: str, caption: str):
    """Embeds an image into Word with center alignment, maximum width constraints, and caption."""
    if not image_path or not os.path.exists(image_path):
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(2)
    run_img = p_img.add_run()
    try:
        run_img.add_picture(image_path, width=Inches(5.0))
    except Exception as e:
        print(f"[WARN] Could not embed image {image_path}: {e}")
        return

    if caption:
        p_cap = doc.add_paragraph()
        set_p_rtl(p_cap, space_before_pt=2, space_after_pt=6)
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_r(p_cap, f"🔍 تصویر آموزشی: {caption}", font_name=FONT_DUBAI, size_pt=9.5, italic=True,
              color_rgb=COLOR_MUTED_GRAY, is_rtl=True)

def add_diagram_picture_block(doc, image_path: str, title: str = "دیاگرام الگوریتم بالینی و مسیر پاتوفیزیولوژی"):
    """Embeds a rendered High-DPI Mermaid diagram into Word with center alignment and clear caption."""
    if not image_path or not os.path.exists(image_path):
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(3)
    run_img = p_img.add_run()
    try:
        run_img.add_picture(image_path, width=Inches(5.8))
    except Exception as e:
        print(f"[WARN] Could not embed diagram picture {image_path}: {e}")
        return

    if title:
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(8)
        add_r(p_cap, f"🧬 {title}", font_name=FONT_DUBAI, size_pt=9.5, bold=True,
              color_rgb=(0x29, 0x80, 0xB9), is_rtl=True)

def add_visual_atlas_section(doc, figures: List[Dict[str, Any]], title: str = "📸 اطلس تصویری و تحلیل اسلایدهای بالینی تدریس استاد"):
    """
    Renders the dedicated Multimodal Visual Atlas in Word (.docx) with High-DPI pictures,
    centered alignment, and structured diagnostic captioning (landmarks, findings, exam pearls).
    """
    if not figures or not isinstance(figures, list):
        return
    valid_figs = []
    for f in figures:
        if isinstance(f, dict):
            f_p = f.get("image_path") or f.get("path") or f.get("src")
            if f_p and os.path.exists(f_p):
                valid_figs.append(f)
    if not valid_figs:
        return

    p_hdr = doc.add_paragraph()
    set_p_rtl(p_hdr, space_before_pt=16, space_after_pt=6)
    add_r(p_hdr, title, font_name=FONT_DUBAI, size_pt=13.0, bold=True,
          color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)

    for fig in valid_figs:
        f_p = fig.get("image_path") or fig.get("path") or fig.get("src")
        f_title = fig.get("title", "تصویر اسلاید کلاسی")
        f_caption = fig.get("caption") or fig.get("clinical_findings") or ""
        f_pearl = fig.get("exam_pearl") or ""

        # Centered Image
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(2)
        run_img = p_img.add_run()
        try:
            run_img.add_picture(f_p, width=Inches(4.8))
        except Exception as e:
            print(f"[WARN] Could not embed atlas image {f_p}: {e}")
            continue

        # Structured Caption
        p_cap = doc.add_paragraph()
        set_p_rtl(p_cap, space_before_pt=2, space_after_pt=8)
        p_cap.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        add_r(p_cap, f"📷 {f_title}\n", font_name=FONT_DUBAI, size_pt=10.0, bold=True,
              color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
        if f_caption:
            add_r(p_cap, f"• تحلیل یافته‌های بالینی و رادیولوژیک: {f_caption}", font_name=FONT_DUBAI, size_pt=9.5,
                  bold=False, color_rgb=COLOR_TEXT_CHARCOAL, is_rtl=True)
        if f_pearl:
            p_cap.add_run("\n")
            add_r(p_cap, f"⭐ نکته کلیدی استاد / تله امتحانی: {f_pearl}", font_name=FONT_DUBAI, size_pt=9.5,
                  bold=True, color_rgb=(0xD3, 0x54, 0x00), is_rtl=True)

def clean_mermaid_code(code: str) -> str:
    """
    Sanitizes Mermaid code, enforces vertical stacking (flowchart TD),
    fixes invalid syntax, and cleans TeX symbols inside node labels.
    """
    if not code or not str(code).strip():
        return "flowchart TD\n    A[ملاحظه بالینی]"

    raw = str(code).strip()
    lines = raw.splitlines()
    first_line = lines[0].strip() if lines else ""

    if first_line.startswith("graph LR") or first_line.startswith("flowchart LR"):
        lines[0] = "flowchart TD"
    elif not (first_line.startswith("graph ") or first_line.startswith("flowchart ")):
        lines.insert(0, "flowchart TD")

    sanitized_lines = []
    for l in lines:
        sl = l
        # Sanitize TeX symbols inside labels
        sl = sl.replace(r'$\rightarrow$', '→').replace(r'\rightarrow', '→')
        sl = sl.replace(r'$\leftarrow$', '←').replace(r'\leftarrow', '←')
        sl = sl.replace(r'$\ge$', '≥').replace(r'\ge', '≥')
        sl = sl.replace(r'$\le$', '≤').replace(r'\le', '≤')
        sl = sl.replace(r'$\pm$', '±').replace(r'\pm', '±')
        sl = sl.replace(r'$\times$', '×').replace(r'\times', '×')
        sl = re.sub(r'\\text\{([^}]+)\}', r'\1', sl)
        sanitized_lines.append(sl)

    return "\n".join(sanitized_lines)


def render_mermaid_diagram_to_png(
    mermaid_code: str,
    output_png_path: str,
    title: str = "دیاگرام الگوریتم بالینی",
    min_width_px: int = 760,
    output_svg_path: Optional[str] = None
) -> bool:
    """
    Renders a Mermaid flowchart/diagram to a High-DPI transparent PNG image
    using Playwright with a specialized clinical theme, Rule 10 typography standard (17px),
    and Golden Aspect Ratio constraints.
    """
    if not mermaid_code or not str(mermaid_code).strip():
        return False

    if os.path.exists(output_png_path) and os.path.getsize(output_png_path) > 1000:
        return True

    os.makedirs(os.path.dirname(os.path.abspath(output_png_path)), exist_ok=True)
    clean_code = clean_mermaid_code(mermaid_code)

    edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    exec_path = edge_exe if os.path.exists(edge_exe) else None

    html_content = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="utf-8">
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
    <script>
        mermaid.initialize({{
            startOnLoad: true,
            theme: 'base',
            themeVariables: {{
                fontFamily: 'Dubai, Vazirmatn, Segoe UI, Tahoma, sans-serif',
                primaryColor: '#EBF5FB',
                primaryBorderColor: '#2980B9',
                primaryTextColor: '#1A5276',
                lineColor: '#2980B9',
                secondaryColor: '#FEF9E7',
                tertiaryColor: '#FDEDEC',
                edgeLabelBackground: '#ffffff',
                fontSize: '17px'
            }}
        }});
    </script>
    <style>
        body {{
            background: #ffffff;
            margin: 0;
            padding: 16px;
            font-family: 'Dubai', 'Vazirmatn', 'Segoe UI', Tahoma, sans-serif;
            direction: rtl;
        }}
        .card {{
            border: 1.5px solid #2980B9;
            border-radius: 10px;
            padding: 16px 20px;
            background: #FAFCFD;
            box-shadow: 0 4px 12px rgba(41, 128, 185, 0.08);
            width: 760px;
            max-width: 760px;
            box-sizing: border-box;
            margin: 0 auto;
            text-align: center;
        }}
        .card-header {{
            font-size: 16px;
            font-weight: bold;
            color: #1A5276;
            margin-bottom: 14px;
            border-bottom: 1.5px solid #AED6F1;
            padding-bottom: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }}
        .mermaid {{
            display: flex;
            justify-content: center;
        }}
        .mermaid svg {{
            max-width: 100% !important;
            height: auto !important;
        }}
    </style>
</head>
<body dir="rtl">
    <div class="card">
        <div class="card-header">🧬 {title}</div>
        <div class="mermaid">
{clean_code}
        </div>
    </div>
</body>
</html>"""

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path=exec_path, headless=True)
            page = browser.new_page(device_scale_factor=2)
            page.set_content(html_content)
            page.wait_for_selector(".mermaid svg", timeout=12000)
            card_el = page.locator(".card")
            card_box = card_el.bounding_box()
            if card_box and card_box["height"] > 0:
                aspect_ratio = card_box["width"] / card_box["height"]
                if aspect_ratio > 1.8:
                    print(f"[WARN] [FLOWCHART_LEGIBILITY_GATE]: Diagram '{title}' aspect ratio {aspect_ratio:.2f} > 1.8 exceeds Golden Aspect Ratio (0.7:1 - 1.6:1). Vertical flowchart TD stacking recommended.")
            card_el.screenshot(path=output_png_path)
            if output_svg_path:
                svg_content = page.locator(".mermaid").inner_html()
                os.makedirs(os.path.dirname(os.path.abspath(output_svg_path)), exist_ok=True)
                with open(output_svg_path, "w", encoding="utf-8") as sf:
                    sf.write(f'<div class="diagram-card"><div class="diagram-header">🧬 {title}</div>{svg_content}</div>')
            browser.close()
        return True
    except Exception as ex:
        print(f"[WARN] Mermaid diagram render fallback: {ex}")
        return False

def add_clinical_vignettes(doc, vignettes: List[Dict[str, Any]]):
    """Renders active recall case scenarios with choices, correct answers and pathophysiological rationale."""
    if not vignettes:
        return
    for idx, v in enumerate(vignettes, 1):
        table = doc.add_table(rows=2, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        apply_table_xml_safety(table, is_multi_row_tabular=False)

        # Header
        h_cell = table.cell(0, 0)
        set_cell_margins(h_cell, top=140, bottom=140, left=180, right=180)
        set_cell_shading(h_cell, "F5EEF8")  # Soft Lavender
        h_cell._tc.get_or_add_tcPr().append(parse_xml(
            r'<w:tcBorders {}><w:top w:val="single" w:sz="12" w:color="6C3483"/><w:left w:val="single" w:sz="12" w:color="6C3483"/><w:right w:val="single" w:sz="12" w:color="6C3483"/><w:bottom w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
        ))
        p_head = h_cell.paragraphs[0]
        set_p_rtl(p_head, space_before_pt=0, space_after_pt=0)
        title_txt = v.get("scenario_title", f"سناریوی بالینی شماره {idx}")
        add_r(p_head, f"📝 خودآزمایی فعال: {title_txt}",
              font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=(0x6C, 0x34, 0x83), is_rtl=True)

        # Body
        b_cell = table.cell(1, 0)
        set_cell_margins(b_cell, top=160, bottom=160, left=180, right=180)
        set_cell_shading(b_cell, "FFFFFF")
        b_cell._tc.get_or_add_tcPr().append(parse_xml(
            r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="single" w:sz="12" w:color="6C3483"/><w:right w:val="single" w:sz="12" w:color="6C3483"/><w:bottom w:val="single" w:sz="12" w:color="6C3483"/></w:tcBorders>'.format(nsdecls('w'))
        ))
        # Scenario text
        p_scen = b_cell.paragraphs[0]
        set_p_rtl(p_scen, space_before_pt=2, space_after_pt=4)
        add_bidi_text(p_scen, v.get("scenario", ""), font_name=FONT_DUBAI, size_pt=10.5, bold=False)

        # Options
        for opt in v.get("options", []):
            p_opt = b_cell.add_paragraph()
            set_p_rtl(p_opt, space_before_pt=1, space_after_pt=2)
            add_bidi_text(p_opt, opt, font_name=FONT_DUBAI, size_pt=10.0, color_rgb=(0x26, 0x26, 0x26))

        # Answer & Explanation
        p_ans = b_cell.add_paragraph()
        set_p_rtl(p_ans, space_before_pt=4, space_after_pt=2)
        add_r(p_ans, f"✓ {v.get('correct_answer', '')}", font_name=FONT_DUBAI, size_pt=10.5, bold=True,
              color_rgb=(0x0E, 0x62, 0x51), is_rtl=True)

        if "rationale" in v:
            p_rat = b_cell.add_paragraph()
            set_p_rtl(p_rat, space_before_pt=1, space_after_pt=2)
            add_bidi_text(p_rat, f"💡 تحلیل فیزیوپاتولوژی: {v['rationale']}", font_name=FONT_DUBAI, size_pt=10.0,
                          italic=True, color_rgb=(0x5D, 0x40, 0x37))

        p_after = doc.add_paragraph()
        set_p_rtl(p_after, space_before_pt=0, space_after_pt=6)


def add_semiology_box(doc, semiology_items: List[Dict[str, Any]]):
    """Renders Track 4: Clinical Semiology & Physical Examination maneuvers in Word."""
    if not semiology_items:
        return
    formatted_items = []
    for s in semiology_items:
        if not isinstance(s, dict):
            continue
        man = s.get("maneuver_or_sign", "")
        tech = s.get("technique", "")
        sig = s.get("clinical_significance", "")
        pearl = s.get("exam_pearl", "")
        line = f"🩺 {man}"
        if tech:
            line += f" — روش انجام: {tech}"
        if sig:
            line += f" ← اهمیت بالینی: {sig}"
        if pearl:
            line += f" (نکته امتحانی: {pearl})"
        formatted_items.append(line)
    if formatted_items:
        add_callout_box(
            doc,
            title="🩺 سمیولوژی و مانورهای بالینی (Clinical Semiology & Physical Exam)",
            content_items=formatted_items,
            border_color=COLOR_SEMIOLOGY_TEAL,
            header_bg=COLOR_SEMIOLOGY_BG,
            body_bg="FFFFFF",
            icon="🩺"
        )


def add_morning_exam_checklist(doc, checklist_items: List[str]):
    """Renders the 30-Point Morning-of-Exam Capsule with check-boxes in Word."""
    if not checklist_items:
        return
    valid_items = [str(x).strip() for x in checklist_items if x and str(x).strip()]
    if not valid_items:
        return
    p_head = doc.add_paragraph()
    set_p_rtl(p_head, space_before_pt=18, space_after_pt=6)
    add_r(p_head, "📋 چک‌لیست طلایی ۳۰ نکته‌ای صبح آزمون (Last-Minute High-Yield Capsule)",
          font_name=FONT_DUBAI, size_pt=14.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)

    table = doc.add_table(rows=len(valid_items), cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    apply_table_xml_safety(table, is_multi_row_tabular=False)

    for idx, item in enumerate(valid_items):
        cell = table.cell(idx, 0)
        set_cell_margins(cell, top=100, bottom=100, left=160, right=160)
        fill_bg = COLOR_CHECKLIST_BG if idx % 2 == 0 else "FFFFFF"
        set_cell_shading(cell, fill_bg)
        set_cell_borders(cell, bottom="single", color="E0E0E0", sz="4")
        p = cell.paragraphs[0]
        set_p_rtl(p, space_before_pt=1, space_after_pt=2)
        add_r(p, "☐ ", font_name="Arial", size_pt=12.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=False)
        add_r(p, f"نکته {idx+1}: ", font_name=FONT_DUBAI, size_pt=10.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
        add_bidi_text(p, item, font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_CHARCOAL)

    p_after = doc.add_paragraph()
    set_p_rtl(p_after, space_before_pt=0, space_after_pt=8)


def add_comparative_matrices(doc, matrices: List[Dict[str, Any]]):
    """Renders specialized comparative matrices (e.g. Thyroiditis, Amenorrhea, Emergency Resuscitation) in Word."""
    if not matrices or not isinstance(matrices, list):
        return
    for mat in matrices:
        if not isinstance(mat, dict):
            continue
        title = mat.get("title", "ماتریس مقایسه‌ای تخصصی")
        headers = mat.get("headers", [])
        rows = mat.get("rows", [])
        if headers and rows:
            add_comparison_table(doc, title=title, headers=headers, rows=rows)

def format_emergency_items(emerg_items: Any) -> List[str]:
    """Polymorphically normalizes emergency management items (dict, list, string) into concise strings."""
    if not emerg_items:
        return []
    if isinstance(emerg_items, dict):
        cond = emerg_items.get("condition") or emerg_items.get("indication")
        act = emerg_items.get("action") or emerg_items.get("intervention") or emerg_items.get("management")
        tgt = emerg_items.get("target") or emerg_items.get("goal")
        lines = []
        if cond and act:
            lines.append(f"شرایط بالینی: {cond} -> اقدام فوری: {act}" + (f" (هدف: {tgt})" if tgt else ""))
        elif act:
            lines.append(f"اقدام فوری: {act}" + (f" (هدف: {tgt})" if tgt else ""))
        elif cond:
            lines.append(f"شرایط بالینی: {cond}")
        for k, v in emerg_items.items():
            if k not in ("condition", "indication", "action", "intervention", "management", "target", "goal") and v:
                lines.append(f"{k}: {v}")
        return lines if lines else [str(emerg_items)]
    elif isinstance(emerg_items, list):
        res = []
        for item in emerg_items:
            if isinstance(item, dict):
                res.extend(format_emergency_items(item))
            elif item and str(item).strip():
                res.append(str(item).strip())
        return res
    elif isinstance(emerg_items, str):
        return [emerg_items.strip()] if emerg_items.strip() else []
    return [str(emerg_items)]

# ---------------------------------------------------------------------------
# Core Summary Document Compilation Engine
# ---------------------------------------------------------------------------

def ensure_master_matrices(data: Dict[str, Any]):
    """
    Ensures both master_pharmacology_matrix and last_minute_crash_sheet
    are synthesized if missing, ensuring 100% content parity across Word and Markdown.
    """
    modules = data.get("modules") or []

    # 1. Master Pharmacology Matrix
    if not data.get("master_pharmacology_matrix"):
        all_pharm_rows = []
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            for drug in (mod.get("pharmacology_and_management") or []):
                if not isinstance(drug, dict):
                    continue
                d_class = drug.get("drug_class", "-") or "-"
                d_mech = drug.get("mechanism_of_action") or drug.get("mechanism") or "-"
                d_ind = drug.get("indications", "-") or "-"
                d_contra = drug.get("contraindications", "-") or "-"
                d_pearl = drug.get("exam_pearl") or drug.get("pearl") or "-"
                all_pharm_rows.append([m_name, d_class, d_mech, d_ind, d_contra, d_pearl])
        if len(all_pharm_rows) >= 2:
            data["master_pharmacology_matrix"] = {
                "title": "ماتریس جامع مقایسه داروشناسی بالینی (Master Pharmacology Matrix)",
                "headers": ["مبحث / بیماری", "دسته دارویی و نمونه‌ها", "مکانیسم اثر سلولی", "اندیکاسیون دقیق", "کنترااندیکاسیون / عوارض مرگبار", "نکته کلیدی استاد / دام آزمون"],
                "rows": all_pharm_rows
            }

    # 2. Last-Minute Crash Sheet
    if not data.get("last_minute_crash_sheet") and len(modules) >= 1:
        crash_rows = []
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = str(mod.get("name") or "")
            mod_type = mod.get("type", "illness_script")
            if mod_type == "clinical_approach" or mod.get("step_by_step_algorithm"):
                m_etio = "رویکرد تشخیصی / تریاژ"
                alg_list = mod.get("step_by_step_algorithm") or ["-"]
                m_mech = str(alg_list[0] if alg_list else "-")[:60]
                m_sign = str(mod.get("initial_presentation") or "-")[:50]
                pv_list = mod.get("pivot_points") or mod.get("high_yield_decision_points") or ["-"]
                m_dx = str(pv_list[0] if pv_list else "-")[:50]
                m_rx = str(mod.get("definitive_action") or "-")[:50]
            elif mod_type == "foundational_concept" or mod.get("feedback_cascade"):
                m_etio = "هموستاز فیزیولوژیک"
                f_list = mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback") or ["-"]
                m_mech = str(f_list[0] if f_list else "-")[:60]
                roles = mod.get("physiological_role") or ["-"]
                m_sign = str(roles[0] if isinstance(roles, list) else roles)[:50]
                tests = mod.get("baseline_tests") or ["-"]
                m_dx = str(tests[0] if isinstance(tests, list) else tests)[:50]
                rels = mod.get("clinical_relevance") or ["-"]
                m_rx = str(rels[0] if isinstance(rels, list) else rels)[:50]
            else:
                etio_items = mod.get("definition_and_etiology") or ["-"]
                m_etio = " ".join(str(x) for x in etio_items if x)[:60] or "-"
                mech_items = mod.get("pathophysiology_cascade") or ["-"]
                m_mech = str(mech_items[0] if mech_items else "-")[:60]
                pres_items = mod.get("clinical_presentation") or []
                m_sign = str(pres_items[0].get("sign", "-") if pres_items and isinstance(pres_items[0], dict) else "-")[:50]
                diag_dict = mod.get("diagnostics") or {}
                m_dx = str(diag_dict.get("best_initial_test") or "-")[:50]
                rx_items = mod.get("pharmacology_and_management") or []
                if rx_items and isinstance(rx_items[0], dict):
                    m_rx = str(rx_items[0].get("drug_class", "-"))[:50]
                elif mod.get("surgical_and_interventional"):
                    surg = mod.get("surgical_and_interventional")
                    m_rx = str(surg[0] if isinstance(surg, list) else surg)[:50]
                elif mod.get("lifestyle_and_conservative"):
                    life = mod.get("lifestyle_and_conservative")
                    m_rx = str(life[0] if isinstance(life, list) else life)[:50]
                else:
                    m_rx = "-"
            crash_rows.append([m_name, m_etio, m_mech, m_sign, m_dx, m_rx])
        if crash_rows:
            data["last_minute_crash_sheet"] = {
                "headers": ["بیماری / مبحث", "اتیولوژی اصلی", "مکانیسم پاتوفیزیولوژی کلیدی", "علامت شاخص / تریاد", "اقدام تشخیصی نخستین", "درمان انتخابی"],
                "rows": crash_rows
            }

    # 3. Morning-of-Exam 30-Point High-Yield Capsule (Auto-synthesized if missing)
    if not data.get("morning_exam_checklist") and len(modules) >= 1:
        bullets = []
        # Priority 0: Thematic Sections & In-cell Badges (v3.9.0 Table-First Hybrid)
        for sec in (data.get("sections") or data.get("thematic_sections") or []):
            if not isinstance(sec, dict):
                continue
            sec_t = sec.get("title", "")
            for row in (sec.get("rows") or []):
                if not isinstance(row, (list, tuple)):
                    continue
                for cell in row:
                    for line in str(cell).splitlines():
                        l = line.strip().lstrip("• ")
                        if any(b in l for b in ["تله امتحانی", "خطای رایج", "استاندارد طلایی", "مهم"]):
                            bullets.append(f"{l} ({sec_t})")

        # Priority 0: Immediate Emergency Stabilization
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            emerg = format_emergency_items(mod.get("emergency_management") or mod.get("immediate_stabilization"))
            for em in emerg:
                if em and str(em).strip():
                    bullets.append(f"احیای فوری و اورژانس ({m_name}): {str(em).strip()}")
        # Priority 1: Instructor Exam Alerts
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            for a in (mod.get("instructor_exam_alerts") or []):
                if a and str(a).strip():
                    bullets.append(f"تأکید مستقیم استاد ({m_name}): {str(a).strip()}")
        # Priority 2: Absolute Contraindications & Lethal Withholds
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            for c in (mod.get("contraindications_and_warnings") or []):
                if c and str(c).strip():
                    bullets.append(f"ممنوعیت مرگبار ({m_name}): {str(c).strip()}")
        # Priority 3: Diagnostic Gold Standards & Staging Systems
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            diag = mod.get("diagnostics") or {}
            gs = diag.get("gold_standard") if isinstance(diag, dict) else None
            if gs and str(gs).strip() and str(gs).lower() not in {"null", "none", "ذکر نشده"}:
                bullets.append(f"استاندارد طلایی تشخیص {m_name}: {str(gs).strip()}")
            crit = mod.get("staging_and_classification") or mod.get("diagnostic_criteria") or mod.get("scoring_criteria")
            if isinstance(crit, list):
                for cr in crit:
                    if isinstance(cr, str) and cr.strip():
                        bullets.append(f"معیار بالینی ({m_name}): {cr.strip()}")
            elif isinstance(crit, dict) and crit.get("title"):
                bullets.append(f"سیستم طبقه‌بندی ({m_name}): {crit.get('title')}")
        # Priority 4: High-Yield Pearls & Clinical Pivot Points
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            for p in (mod.get("high_yield_pearls") or []):
                if p and str(p).strip():
                    bullets.append(f"نکته های‌یلد ({m_name}): {str(p).strip()}")
            for pv in (mod.get("pivot_points") or mod.get("high_yield_decision_points") or []):
                if pv and str(pv).strip():
                    bullets.append(f"نقطه انشعاب بالینی ({m_name}): {str(pv).strip()}")
            for cr in (mod.get("clinical_relevance") or []):
                if cr and str(cr).strip():
                    bullets.append(f"اهمیت بالینی پایه ({m_name}): {str(cr).strip()}")
        # Priority 5: Surgical Indications & Severe Complications
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            surg = mod.get("surgical_and_interventional") or mod.get("surgical_and_interventional_indications") or []
            if isinstance(surg, str):
                surg = [surg]
            for s in surg:
                if s and str(s).strip():
                    bullets.append(f"اندیکاسیون جراحی ({m_name}): {str(s).strip()}")
            comps = mod.get("complications_and_prognosis") or []
            if isinstance(comps, str):
                comps = [comps]
            for comp in comps:
                if comp and str(comp).strip():
                    bullets.append(f"عارضه وخیم ({m_name}): {str(comp).strip()}")
        if bullets:
            data["morning_exam_checklist"] = bullets[:30]


def build_docx_summary(data: Dict[str, Any], output_path: str):
    """
    Builds the complete publication-grade Word study guide document (.docx).
    Enforces the 6-part standardized Physiopathology model + Crash Review Sheet.
    """
    ensure_master_matrices(data)
    doc = docx.Document()

    # Page Margins (1 inch all around)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Document Header / Banner
    p_meta = doc.add_paragraph()
    set_p_rtl(p_meta, space_before_pt=0, space_after_pt=4)
    add_r(p_meta, "🎓 جزوه خلاصه مرور فیزیوپاتولوژی | Medical Study & High-Yield Revision Guide",
          font_name=FONT_DUBAI, size_pt=9.5, bold=False, color_rgb=COLOR_MUTED_GRAY, is_rtl=True)

    # Main Document Title
    topic_title = data.get("topic_title", "خلاصه مبحث فیزیوپاتولوژی")
    english_title = data.get("topic_title_en", "")
    p_title = doc.add_paragraph()
    set_p_rtl(p_title, space_before_pt=4, space_after_pt=6)
    add_r(p_title, f"🩺 {topic_title}", font_name=FONT_DUBAI, size_pt=16.0, bold=True,
          color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
    if english_title:
        add_r(p_title, f"  ({english_title})", font_name=FONT_DUBAI, size_pt=13.0, bold=True,
              color_rgb=COLOR_SECONDARY_TEAL, is_rtl=False)

    # System & Module Tag
    organ_system = data.get("organ_system", "داخلی / فیزیوپاتولوژی")
    p_tag = doc.add_paragraph()
    set_p_rtl(p_tag, space_before_pt=0, space_after_pt=12)
    add_r(p_tag, f"📌 کورس: {organ_system}  |  🎯 سطح: فیزیوپاتولوژی بالینی  |  ⚡ نوع سند: مرور های‌یلد و شب امتحان",
          font_name=FONT_DUBAI, size_pt=10.0, bold=False, color_rgb=COLOR_MUTED_GRAY, is_rtl=True)

    # -------------------------------------------------------------------
    # Thematic Table Sections (v3.9.0 Table-First Hybrid Architecture)
    # -------------------------------------------------------------------
    sections = data.get("sections") or data.get("thematic_sections") or []
    if sections:
        for sec_idx, sec in enumerate(sections, 1):
            if not isinstance(sec, dict):
                continue
            sec_title = sec.get("title") or sec.get("section_title") or f"بخش {sec_idx}: مبحث بالینی"
            sec_headers = sec.get("headers") or []
            sec_rows = sec.get("rows") or []

            p_sec = doc.add_paragraph()
            set_p_rtl(p_sec, space_before_pt=14, space_after_pt=4)
            add_r(p_sec, f"■ {sec_title}", font_name=FONT_DUBAI, size_pt=13.0, bold=True,
                  color_rgb=COLOR_PRIMARY_TEAL, is_rtl=True)

            if sec_headers and sec_rows:
                add_thematic_table(doc, headers=sec_headers, rows=sec_rows)

            sec_flowchart = sec.get("flowchart") or sec.get("mermaid_code")
            if sec_flowchart:
                m_hash = hashlib.md5(sec_flowchart.strip().encode("utf-8")).hexdigest()[:10]
                candidate_png = os.path.join("dist", "diagrams", f"diag_sec_{m_hash}.png")
                if render_mermaid_diagram_to_png(sec_flowchart, candidate_png, title=f"الگوریتم {sec_title}"):
                    add_diagram_picture_block(doc, candidate_png, title=f"الگوریتم بالینی {sec_title}")

            sec_callouts = sec.get("callouts") or []
            if sec_callouts:
                callout_list = sec_callouts if isinstance(sec_callouts, list) else [sec_callouts]
                add_callout_box(
                    doc,
                    title="🚨 هشدار اورژانسی و خطای بالینی",
                    content_items=callout_list,
                    border_color="C0392B",
                    header_bg="FDEDEC",
                    body_bg="FFFFFF",
                    icon="🚨"
                )

    # Iterate over modules/diseases in the data
    modules = data.get("modules", [])
    for mod_idx, mod in enumerate(modules, 1):
        mod_name = mod.get("name", f"مبحث شماره {mod_idx}")
        mod_name_en = mod.get("name_en", "")
        mod_type = mod.get("type") or mod.get("module_type") or "illness_script"
        is_foundation = (
            mod_type == "foundational_concept"
            or bool(mod.get("feedback_cascade"))
            or bool(mod.get("regulatory_mechanisms_and_feedback"))
            or bool(mod.get("physiological_role"))
            or bool(mod.get("anatomical_physiological_basis"))
        )
        is_approach = (
            mod_type == "clinical_approach"
            or bool(mod.get("step_by_step_algorithm"))
        )

        # Heading 1 (Module Title)
        p_h1 = doc.add_paragraph()
        set_p_rtl(p_h1, space_before_pt=14, space_after_pt=6)
        type_badge = ""
        if is_foundation:
            type_badge = " [مفاهیم و فیزیولوژی پایه]"
        elif is_approach:
            type_badge = " [رویکرد و الگوریتم تشخیصی]"
        add_r(p_h1, f"■ {mod_name}{type_badge}", font_name=FONT_DUBAI, size_pt=13.5, bold=True,
              color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
        if mod_name_en:
            add_r(p_h1, f"  ({mod_name_en})", font_name=FONT_DUBAI, size_pt=12.0, bold=True,
                  color_rgb=COLOR_SECONDARY_TEAL, is_rtl=False)

        # Module-level Thematic Tables (v3.9.0 Table-First Hybrid)
        if mod.get("tables") and isinstance(mod["tables"], list):
            for t_item in mod["tables"]:
                if isinstance(t_item, dict):
                    add_thematic_table(doc, headers=t_item.get("headers", []), rows=t_item.get("rows", []), title=t_item.get("title"))
        elif mod.get("thematic_table") and isinstance(mod["thematic_table"], dict):
            t_item = mod["thematic_table"]
            add_thematic_table(doc, headers=t_item.get("headers", []), rows=t_item.get("rows", []), title=t_item.get("title"))

        # -------------------------------------------------------------------
        # Archetype 1: Foundational Physiology & Regulatory Loops
        # -------------------------------------------------------------------
        if is_foundation:
            roles = mod.get("physiological_role") or mod.get("anatomical_physiological_basis")
            if roles:
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۱. نقش و تنظیم فیزیولوژیک (Physiological Role & Homeostasis)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                roles_list = roles if isinstance(roles, list) else [roles]
                for r in roles_list:
                    p_r = doc.add_paragraph()
                    set_p_rtl(p_r, space_before_pt=1, space_after_pt=3)
                    add_r(p_r, "• ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                    add_bidi_text(p_r, str(r), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

            f_casc = mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback")
            if f_casc and isinstance(f_casc, list):
                add_callout_box(
                    doc,
                    title="۲. مسیر فیدبک هورمونی و سیگنالینگ تنظیمی (Regulatory & Feedback Loops)",
                    content_items=f_casc,
                    border_color=COLOR_BLUE_BORDER,
                    header_bg=COLOR_BLUE_BG,
                    body_bg="FFFFFF",
                    icon="🔄"
                )

            if mod.get("baseline_tests") and isinstance(mod["baseline_tests"], list):
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۳. تست‌های ارزیابی عملکرد پایه و مقادیر مرجع (Baseline Functional Tests)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                for t in mod["baseline_tests"]:
                    p_t = doc.add_paragraph()
                    set_p_rtl(p_t, space_before_pt=1, space_after_pt=3)
                    add_r(p_t, "• ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                    add_bidi_text(p_t, str(t), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

            if mod.get("clinical_relevance") and isinstance(mod["clinical_relevance"], list):
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۴. اهمیت بالینی در پاتولوژی و تله‌های امتحانی (Clinical Relevance & Exam Anchors)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                for cr in mod["clinical_relevance"]:
                    p_cr = doc.add_paragraph()
                    set_p_rtl(p_cr, space_before_pt=1, space_after_pt=3)
                    add_r(p_cr, "⚡ ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                    add_bidi_text(p_cr, str(cr), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

        # -------------------------------------------------------------------
        # Archetype 2: Clinical Diagnostic & Decision Approach (Algorithms)
        # -------------------------------------------------------------------
        elif is_approach:
            inits = mod.get("initial_presentation") or mod.get("chief_complaint_or_scenario")
            if inits:
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۱. تظاهر بالینی و شکایت اولیه (Initial Clinical Presentation)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                inits_list = inits if isinstance(inits, list) else [inits]
                for init in inits_list:
                    p_init = doc.add_paragraph()
                    set_p_rtl(p_init, space_before_pt=1, space_after_pt=3)
                    add_r(p_init, "🔹 ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                    add_bidi_text(p_init, str(init), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

            steps = mod.get("step_by_step_algorithm")
            if steps and isinstance(steps, list):
                step_items = []
                for s in steps:
                    if isinstance(s, dict):
                        st = s.get("step", "")
                        ac = s.get("action", "")
                        step_items.append(f"{st}: {ac}" if st and ac else (st or ac))
                    else:
                        step_items.append(str(s))
                add_callout_box(
                    doc,
                    title="۲. الگوریتم گام‌به‌گام تصمیم‌گیری بالینی (Step-by-Step Triage Algorithm)",
                    content_items=step_items,
                    border_color=COLOR_AMBER_BORDER,
                    header_bg=COLOR_AMBER_BG,
                    body_bg="FFFFFF",
                    icon="🧭"
                )

            pivots = mod.get("pivot_points") or mod.get("high_yield_decision_points")
            if pivots and isinstance(pivots, list):
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۳. نقاط انشعاب تشخیصی و تصمیم‌گیری (Decision Bifurcations / Pivot Criteria)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                for pv in pivots:
                    p_pv = doc.add_paragraph()
                    set_p_rtl(p_pv, space_before_pt=1, space_after_pt=3)
                    add_r(p_pv, "⚖️ ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                    add_bidi_text(p_pv, str(pv), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

            if mod.get("definitive_action"):
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۴. اقدام قطعی بعدی (Definitive Next Step / Confirmatory Action)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                actions = mod["definitive_action"] if isinstance(mod["definitive_action"], list) else [mod["definitive_action"]]
                for act in actions:
                    p_act = doc.add_paragraph()
                    set_p_rtl(p_act, space_before_pt=1, space_after_pt=3)
                    add_r(p_act, "⭐ ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                    add_bidi_text(p_act, str(act), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

        # -------------------------------------------------------------------
        # Archetype 3: Disease Illness Script (Pathologies & Syndromes)
        # -------------------------------------------------------------------
        else:
            # 1. Definition, Etiology & Risk Factors
            risks = mod.get("risk_factors_and_epidemiology") or mod.get("risk_factors_and_predisposing")
            defs = mod.get("definition_and_etiology")
            if defs or risks:
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۱. تعریف، اتیولوژی و عوامل خطر (Definition, Etiology & Risk Factors)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

                for b in (defs or []):
                    p_b = doc.add_paragraph()
                    set_p_rtl(p_b, space_before_pt=1, space_after_pt=3)
                    add_r(p_b, "• ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                    add_bidi_text(p_b, str(b), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

                risk_items = risks if isinstance(risks, list) else ([risks] if risks else [])
                if isinstance(risk_items, str):
                    risk_items = [risk_items]
                for r in risk_items:
                    p_r = doc.add_paragraph()
                    set_p_rtl(p_r, space_before_pt=1, space_after_pt=3)
                    add_r(p_r, "👥 ریسک‌فاکتور / اپیدمیولوژی: ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_MUTED_GRAY, is_rtl=True)
                    add_bidi_text(p_r, str(r), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

            # 2. Pathophysiology Cascade (in Callout Box)
            if mod.get("pathophysiology_cascade") and isinstance(mod["pathophysiology_cascade"], list):
                cascade_steps = mod["pathophysiology_cascade"]
                add_callout_box(
                    doc,
                    title="۲. زنجیره و فلوچارت مکانیسم پاتوفیزیولوژی (Pathophysiological Cascade)",
                    content_items=cascade_steps,
                    border_color=COLOR_BLUE_BORDER,
                    header_bg=COLOR_BLUE_BG,
                    body_bg="FFFFFF",
                    icon="🔄"
                )

            # 3. Clinical Presentation (Symptoms & Signs) with Mechanistic Links
            if mod.get("clinical_presentation") and isinstance(mod["clinical_presentation"], list):
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۳. تظاهرات بالینی و توجیه مکانیسمی (Clinical Presentation & Signs)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

                for item in mod["clinical_presentation"]:
                    if not isinstance(item, dict):
                        continue
                    p_item = doc.add_paragraph()
                    set_p_rtl(p_item, space_before_pt=1, space_after_pt=3)
                    sign = item.get("sign", "")
                    mech = item.get("mechanism", "")
                    add_r(p_item, f"• {sign}: ", font_name=FONT_DUBAI, size_pt=10.5, bold=True,
                          color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                    add_bidi_text(p_item, f"علت فیزیوپاتولوژیک ← {mech}", font_name=FONT_DUBAI,
                                  size_pt=10.5, bold=False, color_rgb=COLOR_TEXT_CHARCOAL)

            # Track 4: Clinical Semiology & Physical Examination (سیر سمیولوژی و مانورهای بالینی)
            if mod.get("clinical_semiology") and isinstance(mod["clinical_semiology"], list):
                add_semiology_box(doc, mod["clinical_semiology"])

            # 4. Paraclinical & Diagnostic Findings
            if mod.get("diagnostics") and isinstance(mod["diagnostics"], dict):
                diag = mod["diagnostics"]
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۴. بررسی‌های پاراکلینیک و تشخیصی (Diagnostic Workup)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

                best_init = diag.get("best_initial_test")
                if best_init and str(best_init).strip() and str(best_init).lower() not in {"null", "none", "ذکر نشده"}:
                    p_init = doc.add_paragraph()
                    set_p_rtl(p_init, space_before_pt=1, space_after_pt=2)
                    add_r(p_init, "🔹 اقدام تشخیصی نخستین (Best Next Step / Initial): ",
                          font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                    add_bidi_text(p_init, str(best_init), font_name=FONT_DUBAI, size_pt=10.5)

                gold_std = diag.get("gold_standard")
                if gold_std and str(gold_std).strip() and str(gold_std).lower() not in {"null", "none", "ذکر نشده"}:
                    p_gold = doc.add_paragraph()
                    set_p_rtl(p_gold, space_before_pt=1, space_after_pt=2)
                    add_r(p_gold, "⭐ استاندارد طلایی و قطعی (Gold Standard): ",
                          font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                    add_bidi_text(p_gold, str(gold_std), font_name=FONT_DUBAI, size_pt=10.5)

                if diag.get("key_labs") and isinstance(diag["key_labs"], list):
                    for lab in diag["key_labs"]:
                        if lab and str(lab).strip():
                            p_lab = doc.add_paragraph()
                            set_p_rtl(p_lab, space_before_pt=1, space_after_pt=2)
                            add_r(p_lab, "• ", font_name=FONT_DUBAI, size_pt=10.5, bold=True,
                                  color_rgb=COLOR_MUTED_GRAY, is_rtl=True)
                            add_bidi_text(p_lab, str(lab), font_name=FONT_DUBAI, size_pt=10.5)

            # 5. Comprehensive Management: Medical, Surgical & Non-Pharmacological
            has_pharma = bool(mod.get("pharmacology_and_management"))
            raw_surg = mod.get("surgical_and_interventional") or mod.get("surgical_and_interventional_indications")
            has_surg = bool(raw_surg)
            has_life = bool(mod.get("lifestyle_and_conservative"))

            if has_pharma or has_surg or has_life:
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۵. درمان و مدیریت جامع بالینی (Medical, Surgical & Conservative Rx)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

                if has_pharma and isinstance(mod["pharmacology_and_management"], list):
                    for rx in mod["pharmacology_and_management"]:
                        if not isinstance(rx, dict):
                            continue
                        p_rx = doc.add_paragraph()
                        set_p_rtl(p_rx, space_before_pt=1, space_after_pt=3)
                        drug_class = rx.get("drug_class", "")
                        mechanism = rx.get("mechanism_of_action") or rx.get("mechanism") or ""
                        pearl = rx.get("pearl") or rx.get("exam_pearl") or ""
                        add_r(p_rx, f"💊 داروی انتخابی ({drug_class}): ", font_name=FONT_DUBAI, size_pt=10.5,
                              bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                        text_desc = f"مکانیسم: {mechanism}"
                        if pearl:
                            text_desc += f" | نکته: {pearl}"
                        add_bidi_text(p_rx, text_desc, font_name=FONT_DUBAI, size_pt=10.5)

                if has_surg:
                    surg_items = raw_surg if isinstance(raw_surg, list) else [raw_surg]
                    for s in surg_items:
                        p_s = doc.add_paragraph()
                        set_p_rtl(p_s, space_before_pt=1, space_after_pt=3)
                        add_r(p_s, "🔪 اندیکاسیون جراحی / اقدام مداخله‌ای: ", font_name=FONT_DUBAI, size_pt=10.5,
                              bold=True, color_rgb=(0x92, 0x2B, 0x21), is_rtl=True)
                        add_bidi_text(p_s, str(s), font_name=FONT_DUBAI, size_pt=10.5)

                if has_life:
                    life_items = mod["lifestyle_and_conservative"] if isinstance(mod["lifestyle_and_conservative"], list) else [mod["lifestyle_and_conservative"]]
                    for l in life_items:
                        p_l = doc.add_paragraph()
                        set_p_rtl(p_l, space_before_pt=1, space_after_pt=3)
                        add_r(p_l, "🌱 درمان غیردارویی و تغییر سبک زندگی: ", font_name=FONT_DUBAI, size_pt=10.5,
                              bold=True, color_rgb=(0x11, 0x78, 0x64), is_rtl=True)
                        add_bidi_text(p_l, str(l), font_name=FONT_DUBAI, size_pt=10.5)

            # 6. Complications & Prognosis
            if mod.get("complications_and_prognosis"):
                p_sub = doc.add_paragraph()
                set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
                add_r(p_sub, "۶. عوارض وخیم و عوامل پیش‌آگهی (Complications & Prognosis)",
                      font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                comps = mod["complications_and_prognosis"] if isinstance(mod["complications_and_prognosis"], list) else [mod["complications_and_prognosis"]]
                for comp in comps:
                    p_comp = doc.add_paragraph()
                    set_p_rtl(p_comp, space_before_pt=1, space_after_pt=3)
                    add_r(p_comp, "🚨 ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                    add_bidi_text(p_comp, str(comp), font_name=FONT_DUBAI, size_pt=10.5, color_rgb=COLOR_TEXT_CHARCOAL)

            # 7. Differential Diagnosis Table
            if mod.get("differential_table") and isinstance(mod["differential_table"], dict):
                diff = mod["differential_table"]
                headers = diff.get("headers", ["بیماری / سندروم", "علائم مشترک", "نقطه تمایز کلیدی (Pivot Feature)", "تست تشخیصی قطعی"])
                rows = diff.get("rows", [])
                add_comparison_table(doc, title=f"تشخیص افتراقی‌های مهم {mod_name}", headers=headers, rows=rows)

        # High-Yield Pearls & Traps Box
        if mod.get("high_yield_pearls") and isinstance(mod["high_yield_pearls"], list):
            add_callout_box(
                doc,
                title="⚡ نکات دامی آزمونی و مرواریدهای کلیدی (High-Yield Pearls & Pitfalls)",
                content_items=mod["high_yield_pearls"],
                border_color=COLOR_AMBER_BORDER,
                header_bg=COLOR_AMBER_BG,
                body_bg="FFFFFF",
                icon="⚠️"
            )

        # Instructor Exam Alerts (⭐ تأکیدات قطعی استاد)
        if mod.get("instructor_exam_alerts") and isinstance(mod["instructor_exam_alerts"], list):
            add_callout_box(
                doc,
                title="⭐ اولویت اول مرور: تأکیدات مستقیم و سوالات قطعی استاد در کلاس",
                content_items=mod["instructor_exam_alerts"],
                border_color="D4AC0D",
                header_bg="FEFDE8",
                body_bg="FFFFFF",
                icon="⭐"
            )

        # Emergency Management & Immediate Stabilization (🚨 اقدامات اورژانس و احیا)
        emerg_items = mod.get("emergency_management") or mod.get("immediate_stabilization")
        if emerg_items:
            emerg_list = format_emergency_items(emerg_items)
            add_callout_box(
                doc,
                title="🚨 اقدامات اورژانس و احیای فوری بیمار (Emergency Resuscitation & Immediate Stabilization)",
                content_items=emerg_list,
                border_color="C0392B",
                header_bg="FDEDEC",
                body_bg="FFFFFF",
                icon="🚨"
            )

        # Contraindications & Lethal Withholds (⛔ موارد مصرف ممنوع)
        if mod.get("contraindications_and_warnings") and isinstance(mod["contraindications_and_warnings"], list):
            add_callout_box(
                doc,
                title="⛔ موارد مصرف ممنوع و خطاهای مرگبار (Absolute Contraindications & Lethal Withholds)",
                content_items=mod["contraindications_and_warnings"],
                border_color="C0392B",
                header_bg="FDEDEC",
                body_bg="FFFFFF",
                icon="⛔"
            )

        # Special Populations & Pregnancy (🤰 ملاحظات گروه‌های خاص، بارداری و سالمندان)
        special_items = mod.get("special_populations_and_pregnancy") or mod.get("special_populations") or mod.get("pregnancy_and_pediatrics")
        if special_items:
            spec_list = special_items if isinstance(special_items, list) else [special_items]
            add_callout_box(
                doc,
                title="🤰 ملاحظات گروه‌های خاص، بارداری و سالمندان (Special Populations & Pregnancy)",
                content_items=spec_list,
                border_color="8E44AD",
                header_bg="F4ECF7",
                body_bg="FFFFFF",
                icon="🤰"
            )

        # Clinical Scoring Criteria, Staging & Diagnostic Systems (📋 معیارها و سیستم‌های مرحله‌بندی)
        criteria_obj = mod.get("staging_and_classification") or mod.get("scoring_criteria") or mod.get("diagnostic_criteria")
        if criteria_obj:
            if isinstance(criteria_obj, dict):
                sc_title = criteria_obj.get("title", f"سیستم نمره‌دهی و طبقه‌بندی بالینی {mod_name}")
                sc_headers = criteria_obj.get("headers", ["معیار بالینی", "آستانه / امتیاز", "تفسیر بالینی"])
                sc_rows = criteria_obj.get("rows", [])
                add_comparison_table(doc, title=sc_title, headers=sc_headers, rows=sc_rows)
            elif isinstance(criteria_obj, list) and criteria_obj:
                if isinstance(criteria_obj[0], dict):
                    sc_headers = list(criteria_obj[0].keys())
                    sc_rows = [[str(item.get(k, "-")) for k in sc_headers] for item in criteria_obj]
                    add_comparison_table(doc, title=f"جدول مرحله‌بندی و معیارهای بالینی {mod_name}", headers=sc_headers, rows=sc_rows)
                else:
                    add_callout_box(
                        doc,
                        title=f"📋 سیستم طبقه‌بندی و معیارهای بالینی {mod_name} (Staging & Diagnostic Criteria)",
                        content_items=[str(x) for x in criteria_obj if x],
                        border_color=COLOR_TEAL_BORDER,
                        header_bg=COLOR_TEAL_BG,
                        body_bg="FFFFFF",
                        icon="📋"
                    )

        # Diagnostic Formulas & Calculations in Word (🧮 فرمول‌های تشخیصی)
        f_list = mod.get("formulas") or mod.get("diagnostic_formulas")
        if f_list and isinstance(f_list, list):
            for f_item in f_list:
                if isinstance(f_item, dict):
                    add_formula_card(doc, f_item.get("formula", ""), f_item.get("title", ""))
                elif isinstance(f_item, str):
                    add_formula_card(doc, f_item, f"فرمول تشخیصی {mod_name}")

        # Diagram Rendering (Playwright PNG embed - Dual-Target High-DPI Engine)
        diag_png = mod.get("diagram_image_path") or mod.get("diagram_png")
        m_code = mod.get("mermaid_code") or mod.get("flowchart_mermaid") or mod.get("pathophysiology_cascade_mermaid") or mod.get("algorithm_mermaid")
        d_title = mod.get("diagram_title") or f"دیاگرام الگوریتم بالینی و مسیر پاتوفیزیولوژی {mod_name}"

        if not diag_png and m_code:
            m_hash = hashlib.md5(m_code.strip().encode("utf-8")).hexdigest()[:10]
            candidate_png = os.path.join("dist", "diagrams", f"diag_{m_hash}.png")
            if render_mermaid_diagram_to_png(m_code, candidate_png, title=d_title):
                diag_png = candidate_png

        if diag_png and os.path.exists(diag_png):
            add_diagram_picture_block(doc, diag_png, title=d_title)

        # Embedded Figures & Module-level Visual Atlas (اطلس تصویری و تحلیل اسلایدهای بالینی)
        atlas_figs = mod.get("visual_atlas") or mod.get("clinical_figures") or []
        if atlas_figs:
            add_visual_atlas_section(doc, atlas_figs)
        elif mod.get("figures"):
            for fig in (mod.get("figures") or []):
                if isinstance(fig, dict):
                    add_figure_block(doc, fig.get("path", ""), fig.get("caption", ""))

        # Clinical Vignettes (Active Recall MCQs)
        if mod.get("clinical_vignettes") and isinstance(mod["clinical_vignettes"], list):
            add_clinical_vignettes(doc, mod["clinical_vignettes"])

    # -----------------------------------------------------------------------
    # Track 4: Clinical Semiology & Physical Examination (سمیولوژی و معاینات بالینی)
    # -----------------------------------------------------------------------
    if data.get("clinical_semiology") and isinstance(data["clinical_semiology"], list):
        add_semiology_box(doc, data["clinical_semiology"])

    # -----------------------------------------------------------------------
    # Master Pharmacology Matrix (ماتریس جامع داروشناسی بالینی)
    # -----------------------------------------------------------------------
    if data.get("master_pharmacology_matrix") and isinstance(data["master_pharmacology_matrix"], dict):
        matrix = data["master_pharmacology_matrix"]
        p_matrix_head = doc.add_paragraph()
        set_p_rtl(p_matrix_head, space_before_pt=18, space_after_pt=6)
        add_r(p_matrix_head, f"💊 {matrix.get('title', 'ماتریس جامع مقایسه داروشناسی بالینی')}",
              font_name=FONT_DUBAI, size_pt=14.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)

        headers = matrix.get("headers", ["مبحث / بیماری", "دسته دارویی", "مکانیسم اثر", "اندیکاسیون", "کنترااندیکاسیون", "نکته آزمونی"])
        rows = matrix.get("rows", [])
        add_comparison_table(doc, title=matrix.get("title", "جدول مقایسه داروها"), headers=headers, rows=rows)

    # -----------------------------------------------------------------------
    # Dual-Speed Review: Last-Minute Crash Sheet (جدول کپسولی شب امتحان)
    # -----------------------------------------------------------------------
    if data.get("last_minute_crash_sheet") and isinstance(data["last_minute_crash_sheet"], dict):
        crash = data["last_minute_crash_sheet"]
        p_crash_head = doc.add_paragraph()
        set_p_rtl(p_crash_head, space_before_pt=18, space_after_pt=6)
        add_r(p_crash_head, "🚀 کپسول طلایی شب امتحان (Last-Minute High-Yield Cheat Sheet)",
              font_name=FONT_DUBAI, size_pt=14.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)

        p_desc = doc.add_paragraph()
        set_p_rtl(p_desc, space_before_pt=0, space_after_pt=6)
        add_r(p_desc, "مرور سریع ۲ صفحه‌ای مخصوص صبح آزمون؛ چکیده مهم‌ترین پاتوفیزیولوژی‌ها، تریادها و درمان‌های انتخابی:",
              font_name=FONT_DUBAI, size_pt=10.0, bold=False, color_rgb=COLOR_MUTED_GRAY, is_rtl=True)

        headers = crash.get("headers", ["بیماری", "اتیولوژی اصلی", "مکانیسم پاتوفیزیولوژی کلیدی", "علامت شاخص / تریاد", "اقدام تشخیصی نخستین", "درمان انتخابی"])
        rows = crash.get("rows", [])
        add_comparison_table(doc, title="جدول مقایسه‌ای جامع مباحث برای مرور سریع", headers=headers, rows=rows)

    # -----------------------------------------------------------------------
    # Comparative Matrices (ماتریس‌های تطبیقی جامع تخصصی)
    # -----------------------------------------------------------------------
    if data.get("comparative_matrices") and isinstance(data["comparative_matrices"], list):
        add_comparative_matrices(doc, data["comparative_matrices"])

    # -----------------------------------------------------------------------
    # 30-Point Morning-of-Exam Capsule (چک‌لیست طلایی ۳۰ نکته‌ای صبح آزمون)
    # -----------------------------------------------------------------------
    if data.get("morning_exam_checklist") and isinstance(data["morning_exam_checklist"], list):
        add_morning_exam_checklist(doc, data["morning_exam_checklist"])

    # -----------------------------------------------------------------------
    # Multimodal Visual Atlas (اطلس تصویری و تحلیل اسلایدهای بالینی تدریس استاد)
    # -----------------------------------------------------------------------
    top_atlas = data.get("visual_atlas") or data.get("clinical_figures")
    if top_atlas and isinstance(top_atlas, list):
        add_visual_atlas_section(doc, top_atlas)

    # Save Word Document
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    print(f"[OK] Word Study Guide generated successfully: {output_path}")


# ---------------------------------------------------------------------------
# Markdown Generation Engine
# ---------------------------------------------------------------------------

def format_md_table_row(row) -> str:
    """Safely escapes pipe characters and collapses newlines in Markdown table cells."""
    clean_cells = [
        " ".join(str(c).replace("\r", " ").replace("\n", " ").split()).replace("|", "\\|")
        for c in row
    ]
    return "| " + " | ".join(clean_cells) + " |"

def add_markdown_alert(lines: List[str], alert_type: str, title: str, items: List[Any], icon: str = ""):
    """
    Appends a GitHub/Obsidian alert block with strict line-by-line blockquote continuity (> prefix).
    Sanitizes raw TeX symbols and prevents unrendered LaTeX or split quotes.
    """
    if not items:
        return
    clean_title = sanitize_medical_text_symbols(str(title).strip())
    lines.append(f"> [!{alert_type}] **{clean_title}**")
    for item in items:
        clean_item = sanitize_medical_text_symbols(str(item).strip())
        sub_lines = clean_item.splitlines()
        for s_idx, s_line in enumerate(sub_lines):
            s_clean = s_line.strip()
            if not s_clean:
                continue
            if s_idx == 0:
                bullet = f"{icon} " if icon else ""
                lines.append(f"> - {bullet}{s_clean}")
            else:
                lines.append(f">   {s_clean}")
    lines.append("")


def build_markdown_summary(data: Dict[str, Any], output_path: str):
    """Generates a clean, GitHub/Obsidian-compatible Markdown study guide."""
    ensure_master_matrices(data)
    topic_title = data.get("topic_title", "خلاصه مبحث فیزیوپاتولوژی")
    topic_title_en = data.get("topic_title_en", "")
    organ_system = data.get("organ_system", "فیزیوپاتولوژی")

    lines = []
    lines.append(f"# 🩺 {topic_title} ({topic_title_en})")
    lines.append(f"> **کورس آموزشی:** {organ_system} | **مقطع:** فیزیوپاتولوژی بالینی | **نوع سند:** راهنمای مرور های‌یلد\n")

    # Thematic Table Sections (v3.9.0 Table-First Hybrid Architecture)
    sections = data.get("sections") or data.get("thematic_sections") or []
    if sections:
        for sec_idx, sec in enumerate(sections, 1):
            if not isinstance(sec, dict):
                continue
            sec_title = sec.get("title") or sec.get("section_title") or f"بخش {sec_idx}: مبحث بالینی"
            sec_headers = sec.get("headers") or []
            sec_rows = sec.get("rows") or []

            lines.append(f"## ■ {sec_title}\n")
            if sec_headers and sec_rows:
                lines.append(format_md_table_row(sec_headers))
                lines.append("| " + " | ".join([":---"] * len(sec_headers)) + " |")
                for r in sec_rows:
                    safe_r = r[:len(sec_headers)] if isinstance(r, (list, tuple)) else [str(r)]
                    clean_r = ["<br>".join(str(c).splitlines()) for c in safe_r]
                    lines.append(format_md_table_row(clean_r))
                lines.append("")

            sec_flowchart = sec.get("flowchart") or sec.get("mermaid_code")
            if sec_flowchart:
                lines.append("```mermaid")
                lines.append(clean_mermaid_code(sec_flowchart.strip()))
                lines.append("```\n")

            sec_callouts = sec.get("callouts") or []
            if sec_callouts:
                c_list = sec_callouts if isinstance(sec_callouts, list) else [sec_callouts]
                add_markdown_alert(lines, "WARNING", "هشدار اورژانسی و بالینی", c_list, icon="🚨")

    for mod in (data.get("modules") or []):
        if not isinstance(mod, dict):
            continue
        m_name = mod.get("name", "")
        m_name_en = mod.get("name_en", "")
        mod_type = mod.get("type") or mod.get("module_type") or "illness_script"
        is_foundation = (
            mod_type == "foundational_concept"
            or bool(mod.get("feedback_cascade"))
            or bool(mod.get("regulatory_mechanisms_and_feedback"))
            or bool(mod.get("physiological_role"))
            or bool(mod.get("anatomical_physiological_basis"))
        )
        is_approach = (
            mod_type == "clinical_approach"
            or bool(mod.get("step_by_step_algorithm"))
        )

        type_badge = ""
        if is_foundation:
            type_badge = " [مفاهیم و فیزیولوژی پایه]"
        elif is_approach:
            type_badge = " [رویکرد و الگوریتم تشخیصی]"

        lines.append(f"## ■ {m_name}{type_badge} ({m_name_en})\n")

        # Archetype 1: Foundational Physiology & Regulatory Loops
        if is_foundation:
            roles = mod.get("physiological_role") or mod.get("anatomical_physiological_basis")
            if roles:
                lines.append("### ۱. نقش و تنظیم فیزیولوژیک (Physiological Role & Homeostasis)")
                roles_list = roles if isinstance(roles, list) else [roles]
                for r in roles_list:
                    lines.append(f"- {r}")
                lines.append("")

            f_casc = mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback")
            if f_casc and isinstance(f_casc, list):
                lines.append("### ۲. مسیر فیدبک هورمونی و سیگنالینگ تنظیمی (Regulatory & Feedback Loops)")
                f_cascade = [
                    str(c).replace('"', "'").replace('[', '(').replace(']', ')').replace("\n", " ").strip()
                    for c in f_casc if c and str(c).strip()
                ]
                if len(f_cascade) > 1:
                    lines.append("```mermaid")
                    lines.append("graph TD")
                    for i in range(len(f_cascade) - 1):
                        lines.append(f'    fnode{i}["{f_cascade[i]}"] --> fnode{i+1}["{f_cascade[i+1]}"]')
                    lines.append("```")
                for item in f_cascade:
                    lines.append(f"- **مسیر تنظیمی:** {item}")
                lines.append("")

            if mod.get("baseline_tests") and isinstance(mod["baseline_tests"], list):
                lines.append("### ۳. تست‌های ارزیابی عملکرد پایه و مقادیر مرجع (Baseline Functional Tests)")
                for t in mod["baseline_tests"]:
                    lines.append(f"- • {t}")
                lines.append("")

            if mod.get("clinical_relevance") and isinstance(mod["clinical_relevance"], list):
                lines.append("### ۴. اهمیت بالینی در پاتولوژی و تله‌های امتحانی (Clinical Relevance & Exam Anchors)")
                for cr in mod["clinical_relevance"]:
                    lines.append(f"- ⚡ {cr}")
                lines.append("")

        # Archetype 2: Clinical Diagnostic & Decision Approach (Algorithms)
        elif is_approach:
            inits = mod.get("initial_presentation") or mod.get("chief_complaint_or_scenario")
            if inits:
                lines.append("### ۱. تظاهر بالینی و شکایت اولیه (Initial Clinical Presentation)")
                inits_list = inits if isinstance(inits, list) else [inits]
                for init in inits_list:
                    lines.append(f"- 🔹 {init}")
                lines.append("")

            steps = mod.get("step_by_step_algorithm")
            if steps and isinstance(steps, list):
                lines.append("### ۲. الگوریتم گام‌به‌گام تصمیم‌گیری بالینی (Step-by-Step Triage Algorithm)")
                alg_steps = []
                for s in steps:
                    if isinstance(s, dict):
                        st = s.get("step", "")
                        ac = s.get("action", "")
                        alg_steps.append(f"{st}: {ac}" if st and ac else (st or ac))
                    else:
                        alg_steps.append(str(s).replace('"', "'").replace('[', '(').replace(']', ')').replace("\n", " ").strip())
                if len(alg_steps) > 1:
                    lines.append("```mermaid")
                    lines.append("graph TD")
                    for i in range(len(alg_steps) - 1):
                        lines.append(f'    anode{i}["{alg_steps[i]}"] --> anode{i+1}["{alg_steps[i+1]}"]')
                    lines.append("```")
                for s_idx, s in enumerate(alg_steps, 1):
                    lines.append(f"- **گام {s_idx}:** {s}")
                lines.append("")

            pivots = mod.get("pivot_points") or mod.get("high_yield_decision_points")
            if pivots and isinstance(pivots, list):
                lines.append("### ۳. نقاط انشعاب تشخیصی و تصمیم‌گیری (Decision Bifurcations / Pivot Criteria)")
                for pv in pivots:
                    lines.append(f"- ⚖️ {pv}")
                lines.append("")

            if mod.get("definitive_action"):
                lines.append("### ۴. اقدام قطعی بعدی (Definitive Next Step / Confirmatory Action)")
                actions = mod["definitive_action"] if isinstance(mod["definitive_action"], list) else [mod["definitive_action"]]
                for act in actions:
                    lines.append(f"- ⭐ {act}")
                lines.append("")

        # Archetype 3: Disease Illness Script (Pathologies & Syndromes)
        else:
            # Definition & Etiology & Risk Factors
            risks = mod.get("risk_factors_and_epidemiology") or mod.get("risk_factors_and_predisposing")
            defs = mod.get("definition_and_etiology")
            if defs or risks:
                lines.append("### ۱. تعریف، اتیولوژی و عوامل خطر (Definition, Etiology & Risk Factors)")
                for item in (defs or []):
                    lines.append(f"- {item}")
                risk_items = risks if isinstance(risks, list) else ([risks] if risks else [])
                for r in risk_items:
                    lines.append(f"- 👥 **عامل خطر / اپیدمیولوژی:** {r}")
                lines.append("")

            # Pathophysiology Cascade (with Mermaid diagram)
            if mod.get("pathophysiology_cascade") and isinstance(mod["pathophysiology_cascade"], list):
                lines.append("### ۲. زنجیره پاتوفیزیولوژی (Pathophysiological Cascade)")
                cascade = [
                    str(c).replace('"', "'").replace('[', '(').replace(']', ')').replace("\n", " ").strip()
                    for c in mod["pathophysiology_cascade"] if c and str(c).strip()
                ]
                if len(cascade) > 1:
                    lines.append("```mermaid")
                    lines.append("graph TD")
                    for i in range(len(cascade) - 1):
                        lines.append(f'    node{i}["{cascade[i]}"] --> node{i+1}["{cascade[i+1]}"]')
                    lines.append("```")
                elif len(cascade) == 1:
                    lines.append("```mermaid")
                    lines.append("graph TD")
                    lines.append(f'    node0["{cascade[0]}"]')
                    lines.append("```")
                for item in cascade:
                    lines.append(f"- **گام مکانیسمی:** {item}")
                lines.append("")

            # Clinical Presentation
            if mod.get("clinical_presentation") and isinstance(mod["clinical_presentation"], list):
                lines.append("### ۳. تظاهرات بالینی و ارتباط مکانیسمی (Clinical Presentation)")
                for cp in mod["clinical_presentation"]:
                    if not isinstance(cp, dict):
                        continue
                    sign = cp.get("sign", "")
                    mech = cp.get("mechanism", "")
                    lines.append(f"- **{sign}:** توجیه پاتوفیزیولوژیک ← {mech}")
                lines.append("")

            # Diagnostics
            if mod.get("diagnostics") and isinstance(mod["diagnostics"], dict):
                lines.append("### ۴. بررسی‌های تشخیصی (Diagnostic Workup)")
                diag = mod["diagnostics"]
                best_init = diag.get("best_initial_test")
                if best_init and str(best_init).strip() and str(best_init).lower() not in {"null", "none", "ذکر نشده"}:
                    lines.append(f"- 🔹 **اقدام تشخیصی اولیه (Best Next Step):** {best_init}")
                gold_std = diag.get("gold_standard")
                if gold_std and str(gold_std).strip() and str(gold_std).lower() not in {"null", "none", "ذکر نشده"}:
                    lines.append(f"- ⭐ **استاندارد طلایی (Gold Standard):** {gold_std}")
                for lab in (diag.get("key_labs") or []):
                    if lab and str(lab).strip():
                        lines.append(f"- • {lab}")
                lines.append("")

            # Clinical Semiology & Physical Examination (سیر سمیولوژی و مانورهای بالینی)
            if mod.get("clinical_semiology") and isinstance(mod["clinical_semiology"], list):
                lines.append("### 🩺 سمیولوژی و مانورهای بالینی (Clinical Semiology & Maneuvers)")
                for sem in mod["clinical_semiology"]:
                    if not isinstance(sem, dict):
                        continue
                    man = sem.get("maneuver_or_sign", "")
                    tech = sem.get("technique", "")
                    sig = sem.get("clinical_significance", "")
                    pearl = sem.get("exam_pearl", "")
                    desc = f"- 🩺 **{man}:** "
                    if tech:
                        desc += f"روش انجام ← {tech}"
                    if sig:
                        desc += f" | اهمیت بالینی ← {sig}"
                    if pearl:
                        desc += f" *(نکته آزمون: {pearl})*"
                    lines.append(desc)
                lines.append("")

            # Comprehensive Management
            has_pharma = bool(mod.get("pharmacology_and_management"))
            raw_surg = mod.get("surgical_and_interventional") or mod.get("surgical_and_interventional_indications")
            has_surg = bool(raw_surg)
            has_life = bool(mod.get("lifestyle_and_conservative"))

            if has_pharma or has_surg or has_life:
                lines.append("### ۵. اصول درمان و مدیریت جامع بالینی (Medical, Surgical & Conservative Rx)")
                if has_pharma and isinstance(mod["pharmacology_and_management"], list):
                    for rx in mod["pharmacology_and_management"]:
                        if not isinstance(rx, dict):
                            continue
                        d_class = rx.get("drug_class", "")
                        mechanism = rx.get("mechanism_of_action") or rx.get("mechanism") or ""
                        pearl = rx.get("pearl") or rx.get("exam_pearl") or ""
                        desc = f"مکانیسم عمل: {mechanism}"
                        if pearl:
                            desc += f" (نکته: {pearl})"
                        lines.append(f"- 💊 **{d_class}:** {desc}")
                if has_surg:
                    surg_items = raw_surg if isinstance(raw_surg, list) else [raw_surg]
                    for s in surg_items:
                        lines.append(f"- 🔪 **اندیکاسیون جراحی / مداخله‌ای:** {s}")
                if has_life:
                    life_items = mod["lifestyle_and_conservative"] if isinstance(mod["lifestyle_and_conservative"], list) else [mod["lifestyle_and_conservative"]]
                    for l in life_items:
                        lines.append(f"- 🌱 **درمان غیردارویی و تغییر سبک زندگی:** {l}")
                lines.append("")

            # Complications & Prognosis
            if mod.get("complications_and_prognosis"):
                lines.append("### ۶. عوارض وخیم و پیش‌آگهی (Complications & Prognosis)")
                comps = mod["complications_and_prognosis"] if isinstance(mod["complications_and_prognosis"], list) else [mod["complications_and_prognosis"]]
                for comp in comps:
                    lines.append(f"- 🚨 {comp}")
                lines.append("")

        # Figures
        for fig in (mod.get("figures") or []):
            if isinstance(fig, dict):
                f_path = fig.get("path", "")
                f_cap = fig.get("caption", "")
                lines.append(f"![{f_cap}]({f_path})\n*🔍 تصویر آموزشی: {f_cap}*\n")

        # Differential Table
        if mod.get("differential_table") and isinstance(mod["differential_table"], dict):
            lines.append("### ۶. تشخیص افتراقی (Differential Diagnosis Matrix)")
            dt = mod["differential_table"]
            headers = dt.get("headers", ["بیماری", "علائم مشترک", "نقطه تمایز (Pivot)", "تست قطعی"])
            lines.append(format_md_table_row(headers))
            lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
            for row in dt.get("rows", []):
                lines.append(format_md_table_row(row))
            lines.append("")

        # High Yield Pearls
        if mod.get("high_yield_pearls") and isinstance(mod["high_yield_pearls"], list):
            add_markdown_alert(lines, "WARNING", "نکات دامی آزمونی و مرواریدهای کلیدی (High-Yield Pearls):", mod["high_yield_pearls"], icon="⚡")

        # Instructor Alerts
        if mod.get("instructor_exam_alerts") and isinstance(mod["instructor_exam_alerts"], list):
            add_markdown_alert(lines, "IMPORTANT", "⭐ اولویت اول مرور: تأکیدات مستقیم استاد در کلاس:", mod["instructor_exam_alerts"], icon="⭐")

        # Emergency Management & Immediate Resuscitation
        emerg_items = mod.get("emergency_management") or mod.get("immediate_stabilization")
        if emerg_items:
            emerg_list = format_emergency_items(emerg_items)
            add_markdown_alert(lines, "CAUTION", "🚨 اقدامات اورژانس و احیای فوری بیمار (Emergency Resuscitation & Immediate Stabilization):", emerg_list, icon="🚨")

        # Contraindications & Lethal Withholds
        if mod.get("contraindications_and_warnings") and isinstance(mod["contraindications_and_warnings"], list):
            add_markdown_alert(lines, "CAUTION", "⛔ موارد مصرف ممنوع و خطاهای مرگبار (Absolute Contraindications):", mod["contraindications_and_warnings"], icon="⛔")

        # Special Populations & Pregnancy
        special_items = mod.get("special_populations_and_pregnancy") or mod.get("special_populations") or mod.get("pregnancy_and_pediatrics")
        if special_items:
            spec_list = special_items if isinstance(special_items, list) else [special_items]
            add_markdown_alert(lines, "NOTE", "🤰 ملاحظات گروه‌های خاص، بارداری و سالمندان (Special Populations & Pregnancy):", spec_list, icon="🤰")

        # Diagnostic Formulas in Markdown
        if mod.get("formulas") or mod.get("diagnostic_formulas"):
            f_list = mod.get("formulas") or mod.get("diagnostic_formulas")
            if isinstance(f_list, list):
                lines.append(f"### 🧮 فرمول‌های تشخیصی و محاسبات بالینی {m_name}\n")
                for f_item in f_list:
                    if isinstance(f_item, dict):
                        f_title = f_item.get("title", "فرمول محاسباتی")
                        f_code = sanitize_medical_text_symbols(f_item.get("formula", ""))
                        lines.append(f"> **🧮 {f_title}:** `{f_code}`")
                    elif isinstance(f_item, str):
                        lines.append(f"> **🧮 فرمول:** `{sanitize_medical_text_symbols(f_item)}`")
                lines.append("")

        # Staging & Diagnostic Criteria (Polymorphic)
        criteria_obj = mod.get("staging_and_classification") or mod.get("scoring_criteria") or mod.get("diagnostic_criteria")
        if criteria_obj:
            if isinstance(criteria_obj, dict):
                sc_title = criteria_obj.get("title", f"سیستم نمره‌دهی و طبقه‌بندی {m_name}")
                lines.append(f"### 📋 {sc_title}")
                sc_hdrs = criteria_obj.get("headers", ["معیار", "آستانه", "تفسیر"])
                lines.append(format_md_table_row(sc_hdrs))
                lines.append("| " + " | ".join(["---"] * len(sc_hdrs)) + " |")
                for row in criteria_obj.get("rows", []):
                    lines.append(format_md_table_row(row))
                lines.append("")
            elif isinstance(criteria_obj, list) and criteria_obj:
                if isinstance(criteria_obj[0], dict):
                    lines.append(f"### 📋 جدول مرحله‌بندی و معیارهای بالینی {m_name}")
                    sc_hdrs = list(criteria_obj[0].keys())
                    lines.append(format_md_table_row(sc_hdrs))
                    lines.append("| " + " | ".join(["---"] * len(sc_hdrs)) + " |")
                    for item in criteria_obj:
                        lines.append(format_md_table_row([str(item.get(k, "-")) for k in sc_hdrs]))
                    lines.append("")
                else:
                    lines.append(f"### 📋 سیستم طبقه‌بندی و معیارهای بالینی {m_name} (Staging & Criteria)")
                    for cr in criteria_obj:
                        lines.append(f"- 📋 {cr}")
                    lines.append("")

        # Clinical Vignettes
        if mod.get("clinical_vignettes") and isinstance(mod["clinical_vignettes"], list):
            lines.append("### 📝 خودآزمایی فعال بالینی (Active Recall Clinical Vignettes)")
            for v_idx, v in enumerate(mod["clinical_vignettes"], 1):
                if not isinstance(v, dict):
                    continue
                lines.append(f"#### سناریوی {v_idx}: {v.get('scenario_title', '')}")
                lines.append(f"> {v.get('scenario', '')}\n")
                for opt in v.get("options", []):
                    lines.append(f"- {opt}")
                lines.append("")
                lines.append("<details>")
                lines.append("<summary><b>مشاهده پاسخ صحیح و تحلیل فیزیوپاتولوژی</b></summary>\n")
                lines.append(f"**✓ {v.get('correct_answer', '')}**\n")
                if "rationale" in v:
                    lines.append(f"*تحلیل فیزیوپاتولوژی:* {v['rationale']}\n")
                lines.append("</details>\n")

        # Embedded Figures & Module-level Visual Atlas (اطلس تصویری و تحلیل اسلایدهای بالینی)
        atlas_figs = mod.get("visual_atlas") or mod.get("clinical_figures") or []
        if atlas_figs:
            lines.append(f"### 📸 اطلس تصویری و تحلیل اسلایدهای بالینی {m_name}\n")
            for fig in atlas_figs:
                if not isinstance(fig, dict):
                    continue
                f_title = fig.get("title", "تصویر اسلاید کلاسی")
                f_p = fig.get("image_path") or fig.get("path") or fig.get("src") or ""
                f_cap = fig.get("caption") or fig.get("clinical_findings") or ""
                f_pearl = fig.get("exam_pearl") or ""
                lines.append(f"> **📷 {f_title}**")
                if f_p:
                    lines.append(f"> ![{f_title}]({f_p})")
                if f_cap:
                    lines.append(f"> - **• تحلیل یافته‌های بالینی و رادیولوژیک:** {f_cap}")
                if f_pearl:
                    lines.append(f"> - **⭐ نکته کلیدی استاد / تله امتحانی:** {f_pearl}")
                lines.append("")
        elif mod.get("figures"):
            for fig in (mod.get("figures") or []):
                if isinstance(fig, dict):
                    f_p = fig.get("path", "")
                    f_cap = fig.get("caption", "")
                    if f_p:
                        lines.append(f"![{f_cap}]({f_p})")
                    if f_cap:
                        lines.append(f"*{f_cap}*\n")

    # Track 4: Clinical Semiology & Physical Examination (سمیولوژی و معاینات بالینی)
    if data.get("clinical_semiology") and isinstance(data["clinical_semiology"], list):
        lines.append("## 🩺 سمیولوژی و معاینه بالینی (Track 4: Clinical Semiology & Physical Examination)\n")
        lines.append("| نشانه یا مانور بالینی | نحوه اجرا / تکنیک | اهمیت تشخیصی | بیماری و پاتولوژی مرتبط |")
        lines.append("| --- | --- | --- | --- |")
        for sem in data["clinical_semiology"]:
            if not isinstance(sem, dict):
                continue
            m_sign = sem.get("maneuver_or_sign", "-")
            m_tech = sem.get("technique", "-")
            m_sig = sem.get("clinical_significance", "-")
            m_path = sem.get("pathology_association", "-")
            lines.append(f"| **{m_sign}** | {m_tech} | {m_sig} | *{m_path}* |")
        lines.append("")

    # Master Pharmacology Matrix
    if data.get("master_pharmacology_matrix") and isinstance(data["master_pharmacology_matrix"], dict):
        matrix = data["master_pharmacology_matrix"]
        lines.append(f"## 💊 {matrix.get('title', 'ماتریس جامع مقایسه داروشناسی بالینی (Master Pharmacology Matrix)')}\n")
        headers = matrix.get("headers", [])
        lines.append(format_md_table_row(headers))
        lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
        for row in matrix.get("rows", []):
            lines.append(format_md_table_row(row))
        lines.append("")

    # Crash Sheet
    if data.get("last_minute_crash_sheet") and isinstance(data["last_minute_crash_sheet"], dict):
        lines.append("## 🚀 کپسول طلایی شب امتحان (Last-Minute High-Yield Cheat Sheet)")
        crash = data["last_minute_crash_sheet"]
        headers = crash.get("headers", [])
        lines.append(format_md_table_row(headers))
        lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
        for row in crash.get("rows", []):
            lines.append(format_md_table_row(row))
        lines.append("")

    # Comparative Matrices (ماتریس‌های تطبیقی جامع تخصصی)
    if data.get("comparative_matrices") and isinstance(data["comparative_matrices"], list):
        for mat in data["comparative_matrices"]:
            if not isinstance(mat, dict):
                continue
            title = mat.get("title", "ماتریس مقایسه‌ای تخصصی")
            headers = mat.get("headers", [])
            rows = mat.get("rows", [])
            if headers and rows:
                lines.append(f"## 📊 {title}\n")
                lines.append(format_md_table_row(headers))
                lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                for row in rows:
                    lines.append(format_md_table_row(row))
                lines.append("")

    # 30-Point Morning-of-Exam Capsule (چک‌لیست طلایی ۳۰ نکته‌ای صبح آزمون)
    if data.get("morning_exam_checklist") and isinstance(data["morning_exam_checklist"], list):
        lines.append("## 📋 چک‌لیست طلایی ۳۰ نکته‌ای صبح آزمون (Last-Minute High-Yield Capsule)\n")
        lines.append("> نکات کلیدی جهت مرور سریع ۱۰ دقیقه‌ای قبل از ورود به جلسه امتحان:\n")
        for idx, item in enumerate(data["morning_exam_checklist"], 1):
            lines.append(f"- [ ] **{idx}.** {item}")
        lines.append("")

    # Multimodal Visual Atlas (اطلس تصویری و تحلیل اسلایدهای بالینی تدریس استاد)
    top_atlas = data.get("visual_atlas") or data.get("clinical_figures")
    if top_atlas and isinstance(top_atlas, list):
        lines.append("## 📸 اطلس تصویری و تحلیل اسلایدهای بالینی تدریس استاد\n")
        for fig in top_atlas:
            if not isinstance(fig, dict):
                continue
            f_title = fig.get("title", "تصویر اسلاید کلاسی")
            f_p = fig.get("image_path") or fig.get("path") or fig.get("src") or ""
            f_cap = fig.get("caption") or fig.get("clinical_findings") or ""
            f_pearl = fig.get("exam_pearl") or ""
            lines.append(f"> **📷 {f_title}**")
            if f_p:
                lines.append(f"> ![{f_title}]({f_p})")
            if f_cap:
                lines.append(f"> - **• تحلیل یافته‌های بالینی و رادیولوژیک:** {f_cap}")
            if f_pearl:
                lines.append(f"> - **⭐ نکته کلیدی استاد / تله امتحانی:** {f_pearl}")
            lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[OK] Markdown Summary generated successfully: {output_path}")


# ---------------------------------------------------------------------------
# Anki Flashcard Engine (Wozniak 20 Rules & AnKing Schema Compatible)
# ---------------------------------------------------------------------------

def get_or_synthesize_flashcards(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Retrieves declared flashcards or auto-synthesizes high-yield cards from modules."""
    cards = data.get("anki_flashcards") or []
    if cards:
        return cards

    cards = []

    # -------------------------------------------------------------------
    # Synthesize from Thematic Table Sections (v3.9.0 Table-Row Cloze Engine)
    # -------------------------------------------------------------------
    topic_name = data.get("topic_title", "General")
    sections = data.get("sections") or data.get("thematic_sections") or []
    for sec in sections:
        if not isinstance(sec, dict):
            continue
        sec_title = str(sec.get("title") or sec.get("section_title") or "Table Section").strip()
        for row in (sec.get("rows") or []):
            if not row or not isinstance(row, (list, tuple)):
                continue
            col_0 = str(row[0]).strip().replace("\n", " ") if len(row) > 0 else ""
            col_last = str(row[-1]).strip().replace("\n", " ") if len(row) > 1 else ""

            row_has_badge = False
            for cell_idx, cell in enumerate(row):
                c_str = str(cell).strip()
                for line in c_str.splitlines():
                    l = line.strip().lstrip("• ")
                    if not l:
                        continue
                    if any(b in l for b in ["تله امتحانی", "خطای رایج", "استاندارد طلایی", "مهم"]):
                        row_has_badge = True
                        cards.append({
                            "text": f"در مبحث {sec_title} ({col_0}): {{{{c1::{l}}}}}",
                            "extra": f"نکته کلیدی جدول {sec_title}.",
                            "tag": f"PhysioPath::{topic_name}::TableMatrix"
                        })
            if not row_has_badge and len(row) >= 2 and col_0 and col_last and len(col_last) > 3:
                cards.append({
                    "text": f"در {sec_title}، ویژگی بالینی یا امتحانی {col_0} عبارت است از: {{{{c1::{col_last}}}}}",
                    "extra": f"برگرفته از جدول خلاصه {sec_title}.",
                    "tag": f"PhysioPath::{topic_name}::TableMatrix"
                })

    for mod in (data.get("modules") or []):
        if not isinstance(mod, dict):
            continue
        mod_name = mod.get("name", "")
        # Synthesize from Diagnostics
        diag = mod.get("diagnostics")
        if isinstance(diag, dict):
            best_init = diag.get("best_initial_test")
            if best_init and str(best_init).strip() and str(best_init).lower() not in {"null", "none", "ذکر نشده"}:
                cards.append({
                    "text": f"در بیمار مشکوک به {mod_name}، اقدام تشخیصی نخستین (Best Next Step) عبارت است از {{{{c1::{best_init}}}}}.",
                    "extra": f"بررسی اولیه جهت ارزیابی بالینی سریع و غیرتهاجمی در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Diagnostics"
                })
            gold_std = diag.get("gold_standard")
            if gold_std and str(gold_std).strip() and str(gold_std).lower() not in {"null", "none", "ذکر نشده"}:
                cards.append({
                    "text": f"استاندارد طلایی و قطعی (Gold Standard) برای تشخیص {mod_name} عبارت است از {{{{c1::{gold_std}}}}}.",
                    "extra": f"تست تشخیصی قطعی با بالاترین ویژگی و حساسیت در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Diagnostics"
                })
        # Synthesize from Pearls
        for pearl in (mod.get("high_yield_pearls") or []):
            if pearl and str(pearl).strip():
                pearl_str = str(pearl).strip()
                cards.append({
                    "text": f"نکته های‌یلد {mod_name}: {{{{c1::{pearl_str}}}}}",
                    "extra": f"نکته دامی آزمون پره‌انترنی/دستیاری و فیزیوپاتولوژی مبحث {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::HighYield"
                })

        # Synthesize from Contraindications & Warnings
        for warn in (mod.get("contraindications_and_warnings") or []):
            if warn and str(warn).strip():
                warn_str = str(warn).strip()
                cards.append({
                    "text": f"مورد مصرف ممنوع و خطای مرگبار در {mod_name}: {{{{c1::{warn_str}}}}}",
                    "extra": f"کنترااندیکاسیون قطعی و خطای کشنده پزشکی در مبحث {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Contraindications"
                })

        # Synthesize from Instructor Alerts
        for alert in (mod.get("instructor_exam_alerts") or []):
            if alert and str(alert).strip():
                alert_str = str(alert).strip()
                cards.append({
                    "text": f"تأکید مستقیم استاد در آزمون {mod_name}: {{{{c1::{alert_str}}}}}",
                    "extra": f"نکته اولویت اول و سؤال قطعی امتحانی تدریس‌شده توسط استاد در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::InstructorAlerts"
                })

        # Synthesize from Clinical Semiology (Track 4)
        for sem in (mod.get("clinical_semiology") or []):
            if isinstance(sem, dict) and sem.get("maneuver_or_sign"):
                man = str(sem.get("maneuver_or_sign", "")).strip()
                tech = str(sem.get("technique", "")).strip()
                sig = str(sem.get("clinical_significance", "")).strip()
                cards.append({
                    "text": f"در معاینه فیزیکی و ارزیابی بالینی {mod_name}، مانور یا نشانه {{{{c1::{man}}}}} بررسی می‌شود.",
                    "extra": f"روش و اهمیت بالینی سمیولوژی: {tech} ← {sig}",
                    "tag": f"PhysioPath::{mod_name}::Semiology"
                })

        # Synthesize from Surgical & Interventional Indications
        surg_items = mod.get("surgical_and_interventional") or mod.get("surgical_and_interventional_indications") or []
        if isinstance(surg_items, str):
            surg_items = [surg_items]
        for surg in surg_items:
            if surg and str(surg).strip():
                s_str = str(surg).strip()
                cards.append({
                    "text": f"اندیکاسیون جراحی و اقدام مداخله‌ای در {mod_name}: {{{{c1::{s_str}}}}}",
                    "extra": f"معیارهای تصمیم‌گیری جراحی و اقدام تهاجمی در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Surgery"
                })

        # Synthesize from Complications & Prognosis
        comp_items = mod.get("complications_and_prognosis") or []
        if isinstance(comp_items, str):
            comp_items = [comp_items]
        for comp in comp_items:
            if comp and str(comp).strip():
                c_str = str(comp).strip()
                cards.append({
                    "text": f"عارضه بالینی وخیم و عامل پروگنوستیک در {mod_name}: {{{{c1::{c_str}}}}}",
                    "extra": f"عوارض مهم بیماری و فاکتورهای پیش‌آگهی نامناسب در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Complications"
                })

        # Synthesize from Foundational Feedback Cascades
        f_steps = mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback") or []
        for f_step in f_steps:
            if f_step and str(f_step).strip():
                f_str = str(f_step).strip()
                cards.append({
                    "text": f"در فیزیولوژی و فیدبک تنظیمی {mod_name}: {{{{c1::{f_str}}}}}",
                    "extra": f"مسیر سیگنالینگ و تنظیم هموستاز فیزیولوژیک {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Physiology"
                })

        # Synthesize from Clinical Approach Algorithms & Decision Points
        alg_steps = mod.get("step_by_step_algorithm") or []
        for a_step in alg_steps:
            if a_step:
                if isinstance(a_step, dict):
                    st = a_step.get("step", "")
                    ac = a_step.get("action", "")
                    a_str = f"{st}: {ac}" if st and ac else (st or ac)
                else:
                    a_str = str(a_step).strip()
                if a_str:
                    cards.append({
                        "text": f"در رویکرد تشخیصی و الگوریتم بالینی {mod_name}: {{{{c1::{a_str}}}}}",
                        "extra": f"گام‌های تصمیم‌گیری بالینی و تریاژ تشخیصی در {mod_name}.",
                        "tag": f"PhysioPath::{mod_name}::Approach"
                    })

        dec_points = mod.get("pivot_points") or mod.get("high_yield_decision_points") or []
        for d_pt in dec_points:
            if d_pt and str(d_pt).strip():
                d_str = str(d_pt).strip()
                cards.append({
                    "text": f"نقطه انشعاب تشخیصی و تصمیم‌گیری در {mod_name}: {{{{c1::{d_str}}}}}",
                    "extra": f"معیار تمایز و تصمیم‌گیری قطعی در الگوریتم بالینی {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Approach"
                })

        # Synthesize from Emergency Management & Immediate Stabilization
        emerg_items = format_emergency_items(mod.get("emergency_management") or mod.get("immediate_stabilization"))
        for em in emerg_items:
            if em and str(em).strip():
                e_str = str(em).strip()
                cards.append({
                    "text": f"اقدام اورژانس و احیای فوری در {mod_name}: {{{{c1::{e_str}}}}}",
                    "extra": f"اقدام نجات‌بخش نخستین و پایدارسازی بیمار در شرایط حاد و اورژانس {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Emergency"
                })

        # Synthesize from Staging & Diagnostic Criteria
        crit_obj = mod.get("staging_and_classification") or mod.get("scoring_criteria") or mod.get("diagnostic_criteria") or []
        if isinstance(crit_obj, list):
            for cr in crit_obj:
                if isinstance(cr, dict):
                    cr_str = " | ".join(f"{k}: {v}" for k, v in cr.items() if v)
                else:
                    cr_str = str(cr).strip()
                if cr_str:
                    cards.append({
                        "text": f"معیار بالینی و سیستم مرحله‌بندی {mod_name}: {{{{c1::{cr_str}}}}}",
                        "extra": f"سیستم طبقه‌بندی، نمره‌دهی بالینی و پیش‌آگهی در {mod_name}.",
                        "tag": f"PhysioPath::{mod_name}::Staging"
                    })
        elif isinstance(crit_obj, dict) and crit_obj.get("rows"):
            for row in crit_obj.get("rows", []):
                r_str = " - ".join(str(c) for c in row if c)
                if r_str:
                    cards.append({
                        "text": f"معیار و طبقه‌بندی بالینی {mod_name}: {{{{c1::{r_str}}}}}",
                        "extra": f"سیستم نمره‌دهی و مرحله‌بندی: {crit_obj.get('title', mod_name)}",
                        "tag": f"PhysioPath::{mod_name}::Staging"
                    })

        # Synthesize from Special Populations & Pregnancy
        spec_items = mod.get("special_populations_and_pregnancy") or mod.get("special_populations") or mod.get("pregnancy_and_pediatrics") or []
        if isinstance(spec_items, str):
            spec_items = [spec_items]
        for sp in spec_items:
            if sp and str(sp).strip():
                sp_str = str(sp).strip()
                cards.append({
                    "text": f"ملاحظه در بارداری و گروه‌های خاص در {mod_name}: {{{{c1::{sp_str}}}}}",
                    "extra": f"تعدیل درمانی، عوارض جنینی و ملاحظات پره‌ناتال یا ارگان‌های نارسا در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::SpecialPopulations"
                })

        # Synthesize from Foundational Clinical Relevance
        clin_rels = mod.get("clinical_relevance") or []
        if isinstance(clin_rels, str):
            clin_rels = [clin_rels]
        for cr in clin_rels:
            if cr and str(cr).strip():
                cr_str = str(cr).strip()
                cards.append({
                    "text": f"اهمیت بالینی مفهوم پایه {mod_name}: {{{{c1::{cr_str}}}}}",
                    "extra": f"پل ارتباطی علوم پایه با بالین و سؤالات آزمونی در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::ClinicalRelevance"
                })

        # Synthesize from Module-level Visual Atlas (اطلس تصویری و اسلایدهای بالینی)
        mod_atlas = mod.get("visual_atlas") or mod.get("clinical_figures") or []
        for fig in mod_atlas:
            if isinstance(fig, dict):
                f_title = str(fig.get("title") or "تصویر اسلاید کلاسی").strip()
                f_cap = str(fig.get("caption") or fig.get("clinical_findings") or "").strip()
                f_pearl = str(fig.get("exam_pearl") or "").strip()
                if f_cap:
                    cards.append({
                        "text": f"در تحلیل اسلاید بالینی {f_title} ({mod_name})، یافته کلیدی عبارت است از {{{{c1::{f_cap}}}}}.",
                        "extra": f"نکته تکمیلی / تله امتحانی: {f_pearl}" if f_pearl else f"تحلیل اطلس تصویری و اسلایدهای بالینی مبحث {mod_name}.",
                        "tag": f"PhysioPath::{mod_name}::VisualAtlas"
                    })
                elif f_pearl:
                    cards.append({
                        "text": f"نکته کلیدی استاد در اسلاید بالینی {f_title} ({mod_name}): {{{{c1::{f_pearl}}}}}",
                        "extra": f"نکته دامی آزمون در اطلس تصویری مبحث {mod_name}.",
                        "tag": f"PhysioPath::{mod_name}::VisualAtlas"
                    })

    # Synthesize from Top-Level Clinical Semiology (Track 4)
    topic_name = data.get("topic_title", "General")
    for sem in (data.get("clinical_semiology") or []):
        if isinstance(sem, dict) and sem.get("maneuver_or_sign"):
            man = str(sem.get("maneuver_or_sign", "")).strip()
            tech = str(sem.get("technique", "")).strip()
            sig = str(sem.get("clinical_significance", "")).strip()
            cards.append({
                "text": f"در معاینه فیزیکی و ارزیابی بالینی ({topic_name})، مانور یا نشانه {{{{c1::{man}}}}} ارزیابی می‌شود.",
                "extra": f"روش و اهمیت بالینی سمیولوژی: {tech} ← {sig}",
                "tag": f"PhysioPath::{topic_name}::Semiology"
            })

    # Synthesize from Top-Level Multimodal Visual Atlas (اطلس تصویری تدریس استاد)
    top_atlas = data.get("visual_atlas") or data.get("clinical_figures") or []
    for fig in top_atlas:
        if isinstance(fig, dict):
            f_title = str(fig.get("title") or "تصویر اسلاید کلاسی").strip()
            f_cap = str(fig.get("caption") or fig.get("clinical_findings") or "").strip()
            f_pearl = str(fig.get("exam_pearl") or "").strip()
            if f_cap:
                cards.append({
                    "text": f"در اطلس تصویری و اسلایدهای تدریس استاد ({f_title})، یافته بالینی/رادیولوژیک کلیدی عبارت است از {{{{c1::{f_cap}}}}}.",
                    "extra": f"نکته تکمیلی / تله امتحانی: {f_pearl}" if f_pearl else f"تحلیل اطلس تصویری دوره {topic_name}.",
                    "tag": f"PhysioPath::{topic_name}::VisualAtlas"
                })
            elif f_pearl:
                cards.append({
                    "text": f"نکته طلایی استاد در اسلاید بالینی {f_title}: {{{{c1::{f_pearl}}}}}",
                    "extra": f"نکته تشخیصی اطلس تصویری دوره {topic_name}.",
                    "tag": f"PhysioPath::{topic_name}::VisualAtlas"
                })
    return cards


def build_anki_tsv(data: Dict[str, Any], output_path: str):
    """
    Exports atomic cloze-deletion cards formatted as TSV for direct import into Anki.
    Adheres strictly to Wozniak's 20 Rules:
    1. Atomicity (Single fact per card).
    2. Comprehension before memorization (Mechanism in Extra field).
    3. Cloze deletion syntax: {{c1::cloze_text}}
    """
    cards = get_or_synthesize_flashcards(data)

    tsv_lines = []
    # Header line for Anki import mapping
    # Field 1: Text (Prompt/Cloze), Field 2: Extra, Field 3: Tags
    for card in cards:
        text = str(card.get("text") or card.get("question") or card.get("front") or "").strip()
        if not text:
            continue
        if "{{c" not in text:
            text = f"{{{{c1::{text}}}}}"
        text_clean = sanitize_anki_html(text).replace("\t", " ").replace("\n", "<br>")
        extra_val = card.get("extra") or card.get("answer") or card.get("back") or ""
        extra = sanitize_anki_html(str(extra_val)).replace("\t", " ").replace("\n", "<br>")
        tag = str(card.get("tag", "PhysioPath::General")).replace(" ", "_").replace("\t", "")
        tsv_lines.append(f"{text_clean}\t{extra}\t{tag}")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(tsv_lines))
    print(f"[OK] Anki TSV Deck exported successfully: {output_path} ({len(cards)} cards)")


def build_anki_apkg(data: Dict[str, Any], output_path: str):
    """
    Packages flashcards into an interactive .apkg file if genanki is installed.
    Enforces deterministic deck IDs and card GUIDs anchored exclusively to prompt text
    to prevent card duplication and protect student spaced-repetition schedules.
    """
    try:
        import genanki
    except ImportError:
        print("[INFO] genanki not installed in base python; skipping native .apkg packaging (TSV is ready for Anki import).")
        return

    cards = get_or_synthesize_flashcards(data)
    if not cards:
        print("[INFO] No flashcards found or synthesizable; skipping .apkg export.")
        return

    # Create Cloze Model with model_type=CLOZE for Anki compatibility
    model_id = 1607392319
    cloze_model = genanki.Model(
        model_id,
        'PhysioPath Cloze Model',
        fields=[
            {'name': 'Text'},
            {'name': 'Extra'},
        ],
        templates=[
            {
                'name': 'Cloze Card',
                'qfmt': '<div dir="rtl" style="font-family: Dubai, Arial; font-size: 20px; text-align: right;">{{cloze:Text}}</div>',
                'afmt': '<div dir="rtl" style="font-family: Dubai, Arial; font-size: 20px; text-align: right;">{{cloze:Text}}</div><hr><div dir="rtl" style="font-family: Dubai, Arial; font-size: 16px; color: #78281F; text-align: right;">{{Extra}}</div>',
            },
        ],
        css='.card { font-family: Dubai, Arial; font-size: 18px; text-align: right; color: #262626; background-color: white; } .cloze { font-weight: bold; color: #0E6251; }',
        model_type=genanki.Model.CLOZE
    )

    deck_name = f"PhysioPath::{data.get('topic_title', 'Medical_Summary')}"
    # Deterministic deck ID based on deck name hash (prevents random deck IDs)
    deck_id = int(hashlib.md5(deck_name.encode('utf-8')).hexdigest()[:8], 16)
    deck = genanki.Deck(deck_id, deck_name)

    for c in cards:
        text = str(c.get("text") or c.get("question") or c.get("front") or "").strip()
        if not text:
            continue
        # Ensure cloze syntax exists so Anki creates reviewable cards
        if "{{c" not in text:
            text = f"{{{{c1::{text}}}}}"
        text = sanitize_anki_html(text)
        extra_val = c.get("extra") or c.get("answer") or c.get("back") or ""
        extra = sanitize_anki_html(str(extra_val))
        tags = [str(c.get("tag", "PhysioPath")).replace(" ", "_")]
        # Deterministic GUID anchored strictly to normalized prompt text
        norm_prompt = " ".join(text.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")).replace('\u200c', ' ').split())
        card_guid = genanki.guid_for(norm_prompt)
        note = genanki.Note(
            model=cloze_model,
            fields=[text, extra],
            tags=tags,
            guid=card_guid
        )
        deck.add_note(note)

    genanki.Package(deck).write_to_file(output_path)
    print(f"[OK] Anki Native .apkg Package created successfully: {output_path} (Deterministic Deck ID: {deck_id})")


# ---------------------------------------------------------------------------
# HTML & PDF Generation Engine (Playwright + Persian BiDi Stamping)
# ---------------------------------------------------------------------------

def build_html_summary(data: Dict[str, Any], output_path: Optional[str] = None) -> str:
    """
    Renders an interactive, standalone publication-grade HTML study guide
    optimized for both screen reading and exact A4 PDF printing.
    If output_path is provided, writes to file and returns path; otherwise returns HTML string.
    """
    ensure_master_matrices(data)
    topic_title = data.get("topic_title", "خلاصه مبحث فیزیوپاتولوژی")
    topic_title_en = data.get("topic_title_en", "")
    organ_system = data.get("organ_system", "فیزیوپاتولوژی بالینی")
    institution = data.get("institution", "دانشگاه بق")

    html_parts = []
    html_parts.append("""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>""" + topic_title + """</title>
<style>
@page {
    size: A4 portrait;
    margin-top: 24mm;
    margin-bottom: 20mm;
    margin-left: 17mm;
    margin-right: 17mm;
}
@media print {
    body {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    .page-break {
        page-break-after: always;
        break-after: page;
    }
    .no-break {
        page-break-inside: avoid;
        break-inside: avoid;
    }
}
body {
    font-family: 'Dubai', 'Vazirmatn', Tahoma, 'Segoe UI', Arial, sans-serif;
    direction: rtl;
    text-align: right;
    color: #2E4053;
    background-color: #FFFFFF;
    line-height: 1.6;
    font-size: 9.8pt;
    margin: 0;
    padding: 0;
}
.container {
    max-width: 900px;
    margin: 0 auto;
    padding: 0;
}
.banner {
    border-bottom: 2px solid #007A87;
    padding-bottom: 8px;
    margin-bottom: 18px;
}
.title-main {
    color: #007A87;
    font-size: 17pt;
    font-weight: bold;
    margin: 0 0 4px 0;
}
.title-en {
    color: #0891B2;
    font-size: 12.5pt;
    font-weight: bold;
}
.meta-tag {
    color: #64748B;
    font-size: 9pt;
    margin-top: 4px;
}
.module-header {
    color: #007A87;
    font-size: 13.5pt;
    font-weight: bold;
    border-right: 4px solid #007A87;
    padding-right: 10px;
    margin: 22px 0 10px 0;
}
.section-banner {
    color: #007A87;
    font-size: 13pt;
    font-weight: bold;
    border-right: 4px solid #007A87;
    padding-right: 10px;
    margin: 20px 0 8px 0;
    page-break-after: avoid;
    break-after: avoid;
}
.section-title {
    color: #0B7285;
    font-size: 10.5pt;
    font-weight: bold;
    margin: 12px 0 5px 0;
}
.callout {
    border-radius: 5px;
    margin: 10px 0;
    overflow: hidden;
    page-break-inside: avoid;
    break-inside: avoid;
}
.callout-header {
    font-weight: bold;
    padding: 5px 12px;
    font-size: 10.5pt;
}
.callout-body {
    padding: 8px 12px;
    font-size: 9.5pt;
    line-height: 1.6;
}
.callout-patho { border: 1.5px solid #2980B9; }
.callout-patho .callout-header { background: #EBF5FB; color: #1B4F72; border-bottom: 1px solid #2980B9; }
.callout-patho .callout-body { background: #FAFDFF; }

.callout-semiology { border: 1.5px solid #0D9488; }
.callout-semiology .callout-header { background: #F0FDFA; color: #0F766E; border-bottom: 1px solid #0D9488; }
.callout-semiology .callout-body { background: #FFFFFF; }

.callout-pearl { border: 1.5px solid #D97706; }
.callout-pearl .callout-header { background: #FFFBEB; color: #B45309; border-bottom: 1px solid #D97706; }
.callout-pearl .callout-body { background: #FFFFFD; }

.callout-alert { border: 1.5px solid #F59E0B; }
.callout-alert .callout-header { background: #FEFDE8; color: #7D6608; border-bottom: 1px solid #F59E0B; }
.callout-alert .callout-body { background: #FFFFFB; }

.callout-contra { border: 1.5px solid #DC2626; }
.callout-contra .callout-header { background: #FEF2F2; color: #991B1B; border-bottom: 1px solid #DC2626; }
.callout-contra .callout-body { background: #FFFDFD; }

.callout-emergency { border: 1.5px solid #DC2626; }
.callout-emergency .callout-header { background: #FEF2F2; color: #991B1B; border-bottom: 1.5px solid #DC2626; }
.callout-emergency .callout-body { background: #FFFDFD; font-weight: 500; }

.callout-special { border: 1.5px solid #8E44AD; }
.callout-special .callout-header { background: #F4ECF7; color: #512E5F; border-bottom: 1.5px solid #8E44AD; }
.callout-special .callout-body { background: #FCF9FD; }

.callout-vignette { border: 1.5px solid #6C3483; }
.callout-vignette .callout-header { background: #F5EEF8; color: #512E5F; border-bottom: 1px solid #6C3483; }
.callout-vignette .callout-body { background: #FFFFFF; }

table {
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0;
    font-size: 9.0pt;
    page-break-inside: auto;
    break-inside: auto;
    direction: rtl;
    text-align: right;
}
thead {
    display: table-header-group;
}
tr {
    page-break-inside: avoid;
    break-inside: avoid;
}
th {
    background-color: #007A87;
    color: #FFFFFF;
    font-weight: bold;
    padding: 7px 10px;
    border: 1px solid #007A87;
    text-align: right;
    font-size: 9.2pt;
    direction: rtl;
}
td {
    padding: 6px 9px;
    border: 1px solid #CBD5E1;
    vertical-align: top;
    text-align: right;
    line-height: 1.55;
    direction: rtl;
}
tr:nth-child(even) { background-color: #F0FDFA; }
tr:nth-child(odd) { background-color: #FFFFFF; }

.badge-trap { color: #B91C1C; font-weight: bold; }
.badge-error { color: #C2410C; font-weight: bold; }
.badge-gold { color: #047857; font-weight: bold; }
.badge-note { color: #007A87; font-weight: bold; }
.badge-freq { color: #7C3AED; font-weight: bold; }
.cell-bullet { margin: 2px 0; line-height: 1.5; direction: rtl; text-align: right; }
.cell-line { margin: 2px 0; line-height: 1.5; direction: rtl; text-align: right; }

.checklist-item {
    padding: 5px 8px;
    margin: 3px 0;
    background: #FEFDE8;
    border-bottom: 1px solid #E2E8F0;
    display: flex;
    align-items: center;
    page-break-inside: avoid;
    break-inside: avoid;
}
.box {
    margin-left: 8px;
    font-size: 12pt;
    color: #007A87;
    font-weight: bold;
}
}
.cascade-step {
    padding: 4px 0;
    display: flex;
    align-items: center;
}
.badge-init {
    display: inline-block;
    background: #EBF5FB;
    color: #1B4F72;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: bold;
    margin-left: 6px;
}
.badge-gold {
    display: inline-block;
    background: #FEFDE8;
    color: #7D6608;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: bold;
    margin-left: 6px;
}
.diagram-card {
    margin: 16px 0;
    padding: 14px;
    background: #FDFEFE;
    border: 1px solid #BDC3C7;
    border-radius: 6px;
    text-align: center;
    page-break-inside: avoid;
    break-inside: avoid;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}
.diagram-header {
    font-weight: bold;
    color: #1A5276;
    margin-bottom: 10px;
    font-size: 11pt;
}
.diagram-img {
    max-width: 95%;
    height: auto;
    border-radius: 4px;
    margin: 0 auto;
    display: block;
}
.figure-card {
    margin: 14px 0;
    padding: 12px;
    background: #FDFEFE;
    border: 1px solid #D5D8DC;
    border-right: 4px solid #78281F;
    border-radius: 4px;
    page-break-inside: avoid;
    break-inside: avoid;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.figure-header {
    font-weight: bold;
    color: #78281F;
    font-size: 11pt;
    margin-bottom: 8px;
}
.figure-img {
    max-width: 85%;
    height: auto;
    border-radius: 4px;
    display: block;
    margin: 0 auto 10px auto;
    border: 1px solid #E5E7E9;
}
.figure-caption {
    font-size: 9.5pt;
    color: #2C3E50;
    line-height: 1.6;
    margin-bottom: 4px;
}
.figure-pearl {
    font-size: 9.5pt;
    color: #D35400;
    font-weight: bold;
    line-height: 1.6;
    background: #FEF9E7;
    padding: 4px 8px;
    border-radius: 4px;
    border-right: 3px solid #D35400;
}
.formula-card {
    margin: 16px 0;
    padding: 12px 18px;
    background: #FBFCFD;
    border: 1.5px solid #2980B9;
    border-radius: 6px;
    page-break-inside: avoid;
    break-inside: avoid;
    box-shadow: 0 1px 4px rgba(41, 128, 185, 0.08);
}
.formula-badge {
    font-size: 10pt;
    font-weight: bold;
    color: #1A5276;
    margin-bottom: 8px;
    border-bottom: 1px solid #D4E6F1;
    padding-bottom: 4px;
}
.formula-container {
    display: flex;
    align-items: center;
    justify-content: center;
    direction: ltr;
    font-family: 'Cambria Math', 'Dubai', 'Segoe UI', serif;
    font-size: 11pt;
    color: #1B4F72;
    padding: 8px 0;
    gap: 8px;
}
.formula-fraction {
    display: inline-flex;
    flex-direction: column;
    vertical-align: middle;
    text-align: center;
    padding: 0 4px;
}
.formula-num {
    border-bottom: 1.5px solid #1B4F72;
    padding-bottom: 2px;
}
.formula-denom {
    padding-top: 2px;
}
.formula-caption {
    font-size: 9pt;
    color: #566573;
    margin-top: 6px;
    line-height: 1.5;
    border-top: 1px dashed #E5E7E9;
    padding-top: 4px;
}
.chem-formula {
    display: inline-block;
    direction: ltr;
    font-weight: 600;
    font-family: 'Segoe UI', Tahoma, sans-serif;
}
.med-code {
    background: #F2F4F4;
    color: #C0392B;
    padding: 2px 5px;
    border-radius: 3px;
    font-size: 9pt;
    font-family: Consolas, monospace;
    direction: ltr;
    display: inline-block;
}
</style>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.css">
<script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.js"></script>
<script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/contrib/auto-render.min.js"
    onload="renderMathInElement(document.body, {delimiters: [{left: '$$', right: '$$', display: true}, {left: '$', right: '$', display: false}]});"></script>
<script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
<script>
    document.addEventListener("DOMContentLoaded", function() {
        if (window.mermaid) {
            mermaid.initialize({ startOnLoad: true, theme: 'neutral', securityLevel: 'loose' });
        }
    });
</script>
</head>
<body dir="rtl">
<div class="container">
<div class="banner">
    <div class="title-main">🩺 """ + topic_title + """</div>""")
    if topic_title_en:
        html_parts.append(f'    <div class="title-en">({topic_title_en})</div>')
    html_parts.append(f"""    <div class="meta-tag">🏛️ {institution} | 📌 کورس: {organ_system} | 🎯 سطح: فیزیوپاتولوژی بالینی | ⚡ نوع سند: کتابچه مرور های‌یلد</div>
</div>""")

    # -------------------------------------------------------------------
    # Thematic Table Sections (v3.9.0 Table-First Hybrid Architecture)
    # -------------------------------------------------------------------
    sections = data.get("sections") or data.get("thematic_sections") or []
    if sections:
        for sec_idx, sec in enumerate(sections, 1):
            if not isinstance(sec, dict):
                continue
            sec_title = sec.get("title") or sec.get("section_title") or f"بخش {sec_idx}: مبحث بالینی"
            sec_headers = sec.get("headers") or []
            sec_rows = sec.get("rows") or []

            html_parts.append(f'<div class="section-banner">■ {sec_title}</div>')
            if sec_headers and sec_rows:
                html_parts.append('<table dir="rtl"><thead><tr>')
                for h in sec_headers:
                    html_parts.append(f'<th>{h}</th>')
                html_parts.append('</tr></thead><tbody>')
                for r_idx, row in enumerate(sec_rows):
                    html_parts.append('<tr>')
                    safe_row = row[:len(sec_headers)] if isinstance(row, (list, tuple)) else [str(row)]
                    for col_idx, cell in enumerate(safe_row):
                        html_parts.append(f'<td>{format_cell_html(str(cell))}</td>')
                    html_parts.append('</tr>')
                html_parts.append('</tbody></table>')

            sec_flowchart = sec.get("flowchart") or sec.get("mermaid_code")
            if sec_flowchart:
                html_parts.append(f'<div class="diagram-card"><div class="diagram-header">🧭 الگوریتم بالینی: {sec_title}</div><pre class="mermaid">{clean_mermaid_code(sec_flowchart.strip())}</pre></div>')

            sec_callouts = sec.get("callouts") or []
            if sec_callouts:
                html_parts.append('<div class="callout callout-emergency"><div class="callout-header">🚨 هشدار اورژانسی و بالینی</div><div class="callout-body">')
                for c_item in (sec_callouts if isinstance(sec_callouts, list) else [sec_callouts]):
                    html_parts.append(f'<div>• {format_inline_markdown(str(c_item))}</div>')
                html_parts.append('</div></div>')

    for mod_idx, mod in enumerate((data.get("modules") or []), 1):
        if not isinstance(mod, dict):
            continue
        m_name = mod.get("name", f"مبحث {mod_idx}")
        m_name_en = mod.get("name_en", "")
        mod_type = mod.get("type") or mod.get("module_type") or "illness_script"
        is_foundation = (
            mod_type == "foundational_concept"
            or bool(mod.get("feedback_cascade"))
            or bool(mod.get("regulatory_mechanisms_and_feedback"))
            or bool(mod.get("physiological_role"))
            or bool(mod.get("anatomical_physiological_basis"))
        )
        is_approach = (
            mod_type == "clinical_approach"
            or bool(mod.get("step_by_step_algorithm"))
        )

        type_badge = ""
        if is_foundation:
            type_badge = " [مفاهیم و فیزیولوژی پایه]"
        elif is_approach:
            type_badge = " [رویکرد و الگوریتم تشخیصی]"
        en_tag = f" ({m_name_en})" if m_name_en else ""
        html_parts.append(f'<div class="module-header">■ {m_name}{type_badge}{en_tag}</div>')

        # Archetype 1: Foundational Physiology & Regulatory Loops
        if is_foundation:
            roles = mod.get("physiological_role") or mod.get("anatomical_physiological_basis")
            if roles:
                html_parts.append('<div class="section-title">۱. نقش و تنظیم فیزیولوژیک (Physiological Role & Homeostasis)</div><ul>')
                roles_list = roles if isinstance(roles, list) else [roles]
                for r in roles_list:
                    html_parts.append(f'<li>{r}</li>')
                html_parts.append('</ul>')

            f_casc = mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback")
            if f_casc and isinstance(f_casc, list):
                f_cascade = [str(c).strip() for c in f_casc if c and str(c).strip()]
                if f_cascade:
                    html_parts.append('<div class="callout callout-patho"><div class="callout-header">🔄 ۲. مسیر فیدبک هورمونی و سیگنالینگ تنظیمی (Regulatory Loops)</div><div class="callout-body">')
                    for c_idx, step in enumerate(f_cascade, 1):
                        arrow = " ← " if c_idx < len(f_cascade) else ""
                        html_parts.append(f'<div class="cascade-step"><b>مرحله {c_idx}:</b> {step}{arrow}</div>')
                    html_parts.append('</div></div>')

            if mod.get("baseline_tests") and isinstance(mod["baseline_tests"], list):
                html_parts.append('<div class="section-title">۳. تست‌های ارزیابی عملکرد پایه و مقادیر مرجع (Baseline Functional Tests)</div><ul>')
                for t in mod["baseline_tests"]:
                    html_parts.append(f'<li>• {t}</li>')
                html_parts.append('</ul>')

            if mod.get("clinical_relevance") and isinstance(mod["clinical_relevance"], list):
                html_parts.append('<div class="section-title">۴. اهمیت بالینی در پاتولوژی و تله‌های امتحانی (Clinical Relevance & Exam Anchors)</div><ul>')
                for cr in mod["clinical_relevance"]:
                    html_parts.append(f'<li>⚡ {cr}</li>')
                html_parts.append('</ul>')

        # Archetype 2: Clinical Diagnostic & Decision Approach (Algorithms)
        elif is_approach:
            inits = mod.get("initial_presentation") or mod.get("chief_complaint_or_scenario")
            if inits:
                html_parts.append('<div class="section-title">۱. تظاهر بالینی و شکایت اولیه (Initial Clinical Presentation)</div><ul>')
                inits_list = inits if isinstance(inits, list) else [inits]
                for init in inits_list:
                    html_parts.append(f'<li>🔹 {init}</li>')
                html_parts.append('</ul>')

            steps = mod.get("step_by_step_algorithm")
            if steps and isinstance(steps, list):
                alg_steps = []
                for s in steps:
                    if isinstance(s, dict):
                        st = s.get("step", "")
                        ac = s.get("action", "")
                        alg_steps.append(f"{st}: {ac}" if st and ac else (st or ac))
                    else:
                        alg_steps.append(str(s).strip())
                if alg_steps:
                    html_parts.append('<div class="callout callout-alert"><div class="callout-header">🧭 ۲. الگوریتم گام‌به‌گام تصمیم‌گیری بالینی (Step-by-Step Triage)</div><div class="callout-body">')
                    for s_idx, s in enumerate(alg_steps, 1):
                        arrow = " ← " if s_idx < len(alg_steps) else ""
                        html_parts.append(f'<div class="cascade-step"><b>گام {s_idx}:</b> {s}{arrow}</div>')
                    html_parts.append('</div></div>')

            pivots = mod.get("pivot_points") or mod.get("high_yield_decision_points")
            if pivots and isinstance(pivots, list):
                html_parts.append('<div class="section-title">۳. نقاط انشعاب تشخیصی و تصمیم‌گیری (Decision Bifurcations / Pivot Criteria)</div><ul>')
                for pv in pivots:
                    html_parts.append(f'<li>⚖️ {pv}</li>')
                html_parts.append('</ul>')

            if mod.get("definitive_action"):
                html_parts.append('<div class="section-title">۴. اقدام قطعی بعدی (Definitive Next Step / Confirmatory Action)</div><ul>')
                actions = mod["definitive_action"] if isinstance(mod["definitive_action"], list) else [mod["definitive_action"]]
                for act in actions:
                    html_parts.append(f'<li>⭐ {act}</li>')
                html_parts.append('</ul>')

        # Archetype 3: Disease Illness Script (Pathologies & Syndromes)
        else:
            # Definition & Etiology & Risk Factors
            risks = mod.get("risk_factors_and_epidemiology") or mod.get("risk_factors_and_predisposing")
            defs = mod.get("definition_and_etiology")
            if defs or risks:
                html_parts.append('<div class="section-title">۱. تعریف، اتیولوژی و عوامل خطر (Definition, Etiology & Risk Factors)</div><ul>')
                for item in (defs or []):
                    html_parts.append(f'<li>{item}</li>')
                risk_items = risks if isinstance(risks, list) else ([risks] if risks else [])
                for r in risk_items:
                    html_parts.append(f'<li>👥 <b>عامل خطر / اپیدمیولوژی:</b> {r}</li>')
                html_parts.append('</ul>')

            # Pathophysiology Cascade
            if mod.get("pathophysiology_cascade") and isinstance(mod["pathophysiology_cascade"], list):
                cascade = [str(c).strip() for c in mod["pathophysiology_cascade"] if c and str(c).strip()]
                if cascade:
                    html_parts.append('<div class="callout callout-patho"><div class="callout-header">🔄 ۲. زنجیره پاتوفیزیولوژی (Pathophysiological Cascade)</div><div class="callout-body">')
                    for c_idx, step in enumerate(cascade, 1):
                        arrow = " ← " if c_idx < len(cascade) else ""
                        html_parts.append(f'<div class="cascade-step"><b>گام {c_idx}:</b> {step}{arrow}</div>')
                    html_parts.append('</div></div>')

            # Clinical Presentation
            if mod.get("clinical_presentation") and isinstance(mod["clinical_presentation"], list):
                html_parts.append('<div class="section-title">۳. تظاهرات بالینی و توجیه مکانیسمی (Clinical Presentation)</div><ul>')
                for cp in mod["clinical_presentation"]:
                    if isinstance(cp, dict):
                        sign = cp.get("sign", "")
                        mech = cp.get("mechanism", "")
                        html_parts.append(f'<li><b>{sign}:</b> توجیه مکانیسمی ← {mech}</li>')
                html_parts.append('</ul>')

            # Diagnostics
            if mod.get("diagnostics") and isinstance(mod["diagnostics"], dict):
                diag = mod["diagnostics"]
                html_parts.append('<div class="section-title">۴. بررسی‌های تشخیصی (Diagnostic Workup)</div><ul>')
                b_init = diag.get("best_initial_test")
                if b_init and str(b_init).strip() and str(b_init).lower() not in {"null", "none", "ذکر نشده"}:
                    html_parts.append(f'<li><span class="badge-init">اقدام تشخیصی اولیه</span> {b_init}</li>')
                gold = diag.get("gold_standard")
                if gold and str(gold).strip() and str(gold).lower() not in {"null", "none", "ذکر نشده"}:
                    html_parts.append(f'<li><span class="badge-gold">استاندارد طلایی</span> {gold}</li>')
                for lab in (diag.get("key_labs") or []):
                    if lab and str(lab).strip():
                        html_parts.append(f'<li>• {lab}</li>')
                html_parts.append('</ul>')

            # Track 4: Clinical Semiology
            if mod.get("clinical_semiology") and isinstance(mod["clinical_semiology"], list):
                html_parts.append('<div class="callout callout-semiology"><div class="callout-header">🩺 سمیولوژی و مانورهای بالینی (Clinical Semiology & Maneuvers)</div><div class="callout-body">')
                for sem in mod["clinical_semiology"]:
                    if isinstance(sem, dict):
                        man = sem.get("maneuver_or_sign", "")
                        tech = sem.get("technique", "")
                        sig = sem.get("clinical_significance", "")
                        pearl = sem.get("exam_pearl", "")
                        line = f"<b>🩺 {man}:</b> "
                        if tech:
                            line += f"روش انجام ← {tech}"
                        if sig:
                            line += f" | اهمیت بالینی ← {sig}"
                        if pearl:
                            line += f" <i>(نکته: {pearl})</i>"
                        html_parts.append(f'<div>{line}</div>')
                html_parts.append('</div></div>')

            # Comprehensive Management
            has_pharma = bool(mod.get("pharmacology_and_management"))
            raw_surg = mod.get("surgical_and_interventional") or mod.get("surgical_and_interventional_indications")
            has_surg = bool(raw_surg)
            has_life = bool(mod.get("lifestyle_and_conservative"))

            if has_pharma or has_surg or has_life:
                html_parts.append('<div class="section-title">۵. درمان و مدیریت جامع بالینی (Medical, Surgical & Conservative Rx)</div><ul>')
                if has_pharma and isinstance(mod["pharmacology_and_management"], list):
                    for rx in mod["pharmacology_and_management"]:
                        if isinstance(rx, dict):
                            d_class = rx.get("drug_class", "")
                            mech = rx.get("mechanism_of_action") or rx.get("mechanism") or ""
                            pearl = rx.get("pearl") or rx.get("exam_pearl") or ""
                            p_txt = f" | نکته: {pearl}" if pearl else ""
                            html_parts.append(f'<li>💊 <b>{d_class}:</b> مکانیسم اثر: {mech}{p_txt}</li>')
                if has_surg:
                    surg_items = raw_surg if isinstance(raw_surg, list) else [raw_surg]
                    for s in surg_items:
                        html_parts.append(f'<li>🔪 <b>اندیکاسیون جراحی / مداخله‌ای:</b> {s}</li>')
                if has_life:
                    life_items = mod["lifestyle_and_conservative"] if isinstance(mod["lifestyle_and_conservative"], list) else [mod["lifestyle_and_conservative"]]
                    for l in life_items:
                        html_parts.append(f'<li>🌱 <b>درمان غیردارویی و تغییر سبک زندگی:</b> {l}</li>')
                html_parts.append('</ul>')

            # Complications & Prognosis
            if mod.get("complications_and_prognosis"):
                html_parts.append('<div class="section-title">۶. عوارض وخیم و عوامل پیش‌آگهی (Complications & Prognosis)</div><ul>')
                comps = mod["complications_and_prognosis"] if isinstance(mod["complications_and_prognosis"], list) else [mod["complications_and_prognosis"]]
                for comp in comps:
                    html_parts.append(f'<li>🚨 {comp}</li>')
                html_parts.append('</ul>')

        # Differential Diagnosis Table
        if mod.get("differential_table") and isinstance(mod["differential_table"], dict):
            dt = mod["differential_table"]
            headers = dt.get("headers", ["بیماری", "علائم مشترک", "نقطه تمایز (Pivot)", "تست قطعی"])
            rows = dt.get("rows", [])
            if headers and rows:
                html_parts.append(f'<div class="section-title">۶. تشخیص افتراقی ({m_name})</div><table dir="rtl"><thead><tr>')
                for h in headers:
                    html_parts.append(f'<th>{h}</th>')
                html_parts.append('</tr></thead><tbody>')
                for row in rows:
                    html_parts.append('<tr>')
                    for cell in row:
                        html_parts.append(f'<td>{cell}</td>')
                    html_parts.append('</tr>')
                html_parts.append('</tbody></table>')

        # High Yield Pearls
        if mod.get("high_yield_pearls") and isinstance(mod["high_yield_pearls"], list):
            html_parts.append('<div class="callout callout-pearl"><div class="callout-header">⚡ نکات دامی آزمونی و مرواریدهای کلیدی (High-Yield Pearls)</div><div class="callout-body">')
            for p in mod["high_yield_pearls"]:
                html_parts.append(f'<div>⚡ {format_inline_markdown(p)}</div>')
            html_parts.append('</div></div>')

        # Instructor Alerts
        if mod.get("instructor_exam_alerts") and isinstance(mod["instructor_exam_alerts"], list):
            html_parts.append('<div class="callout callout-alert"><div class="callout-header">⭐ تأکیدات مستقیم استاد در کلاس</div><div class="callout-body">')
            for a in mod["instructor_exam_alerts"]:
                html_parts.append(f'<div>⭐ {format_inline_markdown(a)}</div>')
            html_parts.append('</div></div>')

        # Emergency Management & Immediate Stabilization
        emerg_items = mod.get("emergency_management") or mod.get("immediate_stabilization")
        if emerg_items:
            emerg_list = format_emergency_items(emerg_items)
            html_parts.append('<div class="callout callout-emergency"><div class="callout-header">🚨 اقدامات اورژانس و احیای فوری بیمار (Emergency Resuscitation & Immediate Stabilization)</div><div class="callout-body">')
            for em in emerg_list:
                html_parts.append(f'<div>🚨 {format_inline_markdown(em)}</div>')
            html_parts.append('</div></div>')

        # Contraindications & Warnings
        if mod.get("contraindications_and_warnings") and isinstance(mod["contraindications_and_warnings"], list):
            html_parts.append('<div class="callout callout-contra"><div class="callout-header">⛔ موارد مصرف ممنوع و خطاهای مرگبار (Absolute Contraindications)</div><div class="callout-body">')
            for c in mod["contraindications_and_warnings"]:
                html_parts.append(f'<div>⛔ {format_inline_markdown(c)}</div>')
            html_parts.append('</div></div>')

        # Special Populations & Pregnancy
        special_items = mod.get("special_populations_and_pregnancy") or mod.get("special_populations") or mod.get("pregnancy_and_pediatrics")
        if special_items:
            spec_list = special_items if isinstance(special_items, list) else [special_items]
            html_parts.append('<div class="callout callout-special"><div class="callout-header">🤰 ملاحظات گروه‌های خاص، بارداری و سالمندان (Special Populations & Pregnancy)</div><div class="callout-body">')
            for sp in spec_list:
                html_parts.append(f'<div>🤰 {format_inline_markdown(sp)}</div>')
            html_parts.append('</div></div>')

        # Diagnostic Formulas & Calculations in HTML
        f_list = mod.get("formulas") or mod.get("diagnostic_formulas")
        if f_list and isinstance(f_list, list):
            for f_item in f_list:
                if isinstance(f_item, dict):
                    html_parts.append(render_formula_html(f_item.get("formula", ""), f_item.get("title", "")))
                elif isinstance(f_item, str):
                    html_parts.append(render_formula_html(f_item, f"فرمول تشخیصی {m_name}"))

        # Staging & Diagnostic Criteria (Polymorphic)
        criteria_obj = mod.get("staging_and_classification") or mod.get("scoring_criteria") or mod.get("diagnostic_criteria")
        if criteria_obj:
            if isinstance(criteria_obj, dict):
                sc_title = criteria_obj.get("title", f"سیستم نمره‌دهی و طبقه‌بندی {m_name}")
                headers = criteria_obj.get("headers", ["معیار", "آستانه", "تفسیر"])
                rows = criteria_obj.get("rows", [])
                if headers and rows:
                    html_parts.append(f'<div class="section-title">📋 {sc_title}</div><table dir="rtl"><thead><tr>')
                    for h in headers:
                        html_parts.append(f'<th>{h}</th>')
                    html_parts.append('</tr></thead><tbody>')
                    for row in rows:
                        html_parts.append('<tr>')
                        for cell in row:
                            html_parts.append(f'<td>{format_inline_markdown(cell)}</td>')
                        html_parts.append('</tr>')
                    html_parts.append('</tbody></table>')
            elif isinstance(criteria_obj, list) and criteria_obj:
                if isinstance(criteria_obj[0], dict):
                    headers = list(criteria_obj[0].keys())
                    html_parts.append(f'<div class="section-title">📋 جدول مرحله‌بندی و معیارهای بالینی {m_name}</div><table dir="rtl"><thead><tr>')
                    for h in headers:
                        html_parts.append(f'<th>{h}</th>')
                    html_parts.append('</tr></thead><tbody>')
                    for item in criteria_obj:
                        html_parts.append('<tr>')
                        for k in headers:
                            html_parts.append(f'<td>{format_inline_markdown(item.get(k, "-"))}</td>')
                        html_parts.append('</tr>')
                    html_parts.append('</tbody></table>')
                else:
                    html_parts.append(f'<div class="section-title">📋 معیارهای بالینی و طبقه‌بندی {m_name}</div><ul>')
                    for cr in criteria_obj:
                        html_parts.append(f'<li>📋 {format_inline_markdown(cr)}</li>')
                    html_parts.append('</ul>')

        # Clinical Vignettes
        if mod.get("clinical_vignettes") and isinstance(mod["clinical_vignettes"], list):
            for v_idx, v in enumerate(mod["clinical_vignettes"], 1):
                if isinstance(v, dict):
                    v_title = v.get("scenario_title", f"سناریوی بالینی {v_idx}")
                    html_parts.append(f'<div class="callout callout-vignette"><div class="callout-header">📝 خودآزمایی فعال: {v_title}</div><div class="callout-body">')
                    html_parts.append(f'<p>{v.get("scenario", "")}</p><ul>')
                    for opt in v.get("options", []):
                        html_parts.append(f'<li>{opt}</li>')
                    html_parts.append('</ul>')
                    corr = v.get("correct_answer", "")
                    rat = v.get("rationale", "")
                    html_parts.append(f'<details style="margin-top: 8px; cursor: pointer;"><summary style="font-weight: bold; color: #0E6251;">مشاهده پاسخ صحیح و تحلیل فیزیوپاتولوژی</summary><div style="padding: 8px 12px; background: #FAF9F6; margin-top: 4px; border-right: 3px solid #0E6251;"><b>✓ {corr}</b><br><i>تحلیل مکانیسمی: {rat}</i></div></details>')
                    html_parts.append('</div></div>')

        # Embedded Diagrams (Mermaid / Rendered High-DPI PNG)
        d_img = mod.get("diagram_image_path") or mod.get("diagram_png_path")
        d_code = mod.get("mermaid_code") or mod.get("mermaid_diagram") or (mod.get("diagram") if isinstance(mod.get("diagram"), str) else None)
        d_title = mod.get("diagram_title") or f"دیاگرام الگوریتم بالینی و پاتوفیزیولوژی {m_name}"
        if d_img or d_code:
            html_parts.append('<div class="diagram-card">')
            html_parts.append(f'<div class="diagram-header">🧬 {d_title}</div>')
            if d_img and os.path.exists(d_img):
                html_parts.append(f'<img class="diagram-img" src="{d_img}" alt="{d_title}" />')
            elif d_code:
                clean_mermaid = clean_mermaid_code(d_code)
                html_parts.append(f'<div class="mermaid">{clean_mermaid}</div>')
            html_parts.append('</div>')

        # Module-level Multimodal Visual Atlas (اطلس تصویری و اسلایدهای بالینی)
        mod_atlas = mod.get("visual_atlas") or mod.get("clinical_figures") or []
        if mod_atlas:
            html_parts.append(f'<div class="section-title">📸 اطلس تصویری و اسلایدهای بالینی {m_name}</div>')
            for fig in mod_atlas:
                if not isinstance(fig, dict):
                    continue
                f_title = fig.get("title", "تصویر اسلاید کلاسی")
                f_p = fig.get("image_path") or fig.get("path") or fig.get("src") or ""
                f_cap = fig.get("caption") or fig.get("clinical_findings") or ""
                f_pearl = fig.get("exam_pearl") or ""
                html_parts.append('<div class="figure-card">')
                html_parts.append(f'<div class="figure-header">📷 {f_title}</div>')
                if f_p:
                    html_parts.append(f'<img class="figure-img" src="{f_p}" alt="{f_title}" />')
                if f_cap:
                    html_parts.append(f'<div class="figure-caption"><b>• تحلیل یافته‌های بالینی و رادیولوژیک:</b> {f_cap}</div>')
                if f_pearl:
                    html_parts.append(f'<div class="figure-pearl">⭐ نکته کلیدی استاد / تله امتحانی: {f_pearl}</div>')
                html_parts.append('</div>')

    # Top-Level Track 4: Clinical Semiology
    if data.get("clinical_semiology") and isinstance(data["clinical_semiology"], list):
        html_parts.append('<h2 style="color: #0E6251; border-bottom: 2px solid #16A085; padding-bottom: 6px; margin-top: 30px;">🩺 سمیولوژی و معاینه بالینی (Track 4: Clinical Semiology & Physical Examination)</h2>')
        html_parts.append('<div class="callout callout-semiology"><div class="callout-header">🩺 مانورهای بالینی و نشانه‌های فیزیکی</div><div class="callout-body">')
        for sem in data["clinical_semiology"]:
            if isinstance(sem, dict):
                man = sem.get("maneuver_or_sign", "")
                tech = sem.get("technique", "")
                sig = sem.get("clinical_significance", "")
                path = sem.get("pathology_association", "")
                pearl = sem.get("exam_pearl", "")
                line = f"<b>🩺 {man}:</b> "
                if tech:
                    line += f"روش انجام ← {tech}"
                if sig:
                    line += f" | اهمیت بالینی ← {sig}"
                if path:
                    line += f" | پاتولوژی مرتبط ← <i>{path}</i>"
                if pearl:
                    line += f" <i>(نکته: {pearl})</i>"
                html_parts.append(f'<div style="margin-bottom: 6px;">{format_inline_markdown(line)}</div>')
        html_parts.append('</div></div>')

    # Master Pharmacology Matrix
    if data.get("master_pharmacology_matrix") and isinstance(data["master_pharmacology_matrix"], dict):
        matrix = data["master_pharmacology_matrix"]
        m_title = matrix.get("title", "ماتریس جامع مقایسه داروشناسی بالینی")
        headers = matrix.get("headers", [])
        rows = matrix.get("rows", [])
        if headers and rows:
            html_parts.append(f'<h2 style="color: #78281F; border-bottom: 2px solid #78281F; padding-bottom: 6px; margin-top: 30px;">💊 {m_title}</h2>')
            html_parts.append('<table dir="rtl"><thead><tr>')
            for h in headers:
                html_parts.append(f'<th>{h}</th>')
            html_parts.append('</tr></thead><tbody>')
            for row in rows:
                html_parts.append('<tr>')
                for cell in row:
                    html_parts.append(f'<td>{format_inline_markdown(cell)}</td>')
                html_parts.append('</tr>')
            html_parts.append('</tbody></table>')

    # Comparative Matrices
    if data.get("comparative_matrices") and isinstance(data["comparative_matrices"], list):
        for mat in data["comparative_matrices"]:
            if isinstance(mat, dict):
                c_title = mat.get("title", "ماتریس مقایسه‌ای تخصصی")
                headers = mat.get("headers", [])
                rows = mat.get("rows", [])
                if headers and rows:
                    html_parts.append(f'<h2 style="color: #78281F; border-bottom: 2px solid #78281F; padding-bottom: 6px; margin-top: 30px;">📊 {c_title}</h2>')
                    html_parts.append('<table dir="rtl"><thead><tr>')
                    for h in headers:
                        html_parts.append(f'<th>{h}</th>')
                    html_parts.append('</tr></thead><tbody>')
                    for row in rows:
                        html_parts.append('<tr>')
                        for cell in row:
                            html_parts.append(f'<td>{format_inline_markdown(cell)}</td>')
                        html_parts.append('</tr>')
                    html_parts.append('</tbody></table>')

    # Crash Sheet
    if data.get("last_minute_crash_sheet") and isinstance(data["last_minute_crash_sheet"], dict):
        crash = data["last_minute_crash_sheet"]
        headers = crash.get("headers", [])
        rows = crash.get("rows", [])
        if headers and rows:
            html_parts.append('<h2 style="color: #78281F; border-bottom: 2px solid #78281F; padding-bottom: 6px; margin-top: 30px;">🚀 کپسول طلایی شب امتحان (Last-Minute High-Yield Cheat Sheet)</h2>')
            html_parts.append('<table dir="rtl"><thead><tr>')
            for h in headers:
                html_parts.append(f'<th>{h}</th>')
            html_parts.append('</tr></thead><tbody>')
            for row in rows:
                html_parts.append('<tr>')
                for cell in row:
                    html_parts.append(f'<td>{format_inline_markdown(cell)}</td>')
                html_parts.append('</tr>')
            html_parts.append('</tbody></table>')

    # Morning Exam Checklist
    if data.get("morning_exam_checklist") and isinstance(data["morning_exam_checklist"], list):
        html_parts.append('<h2 style="color: #007A87; border-bottom: 2px solid #007A87; padding-bottom: 6px; margin-top: 30px;">📋 چک‌لیست طلایی ۳۰ نکته‌ای صبح آزمون (Last-Minute High-Yield Capsule)</h2>')
        for idx, item in enumerate(data["morning_exam_checklist"], 1):
            html_parts.append(f'<div class="checklist-item"><span class="box">☐</span> <span><b>نکته {idx}:</b> {format_inline_markdown(item)}</span></div>')

    # Multimodal Visual Atlas (اطلس تصویری و تحلیل اسلایدهای بالینی تدریس استاد)
    top_atlas = data.get("visual_atlas") or data.get("clinical_figures")
    if top_atlas and isinstance(top_atlas, list):
        html_parts.append('<h2 style="color: #007A87; border-bottom: 2px solid #007A87; padding-bottom: 6px; margin-top: 30px;">📸 اطلس تصویری و تحلیل اسلایدهای بالینی تدریس استاد</h2>')
        for fig in top_atlas:
            if not isinstance(fig, dict):
                continue
            f_title = fig.get("title", "تصویر اسلاید کلاسی")
            f_p = fig.get("image_path") or fig.get("path") or fig.get("src") or ""
            f_cap = fig.get("caption") or fig.get("clinical_findings") or ""
            f_pearl = fig.get("exam_pearl") or ""
            html_parts.append('<div class="figure-card">')
            html_parts.append(f'<div class="figure-header">📷 {f_title}</div>')
            if f_p:
                html_parts.append(f'<img class="figure-img" src="{f_p}" alt="{f_title}" />')
            if f_cap:
                html_parts.append(f'<div class="figure-caption"><b>• تحلیل یافته‌های بالینی و رادیولوژیک:</b> {format_inline_markdown(f_cap)}</div>')
            if f_pearl:
                html_parts.append(f'<div class="figure-pearl">⭐ نکته کلیدی استاد / تله امتحانی: {format_inline_markdown(f_pearl)}</div>')
            html_parts.append('</div>')

    html_parts.append('</div></body></html>')
    full_html = "\n".join(html_parts)

    if output_path:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(full_html)
        print(f"[OK] HTML Summary generated successfully: {output_path}")
        return output_path
    return full_html


def stamp_persian_page_numbers(pdf_bytes: bytes, output_path: str, course_name: str = "فیزیوپاتولوژی بالینی", chapter_title: str = "", institution: str = "دانشگاه بق") -> bool:
    """
    Stamps Persian BiDi page numbers ('صفحه X از Y') onto each page of the PDF
    using reportlab, pypdf, arabic_reshaper, and python-bidi.
    Header Left explicitly displays institution ('دانشگاه بق').
    """
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from pypdf import PdfReader, PdfWriter

        def shape_persian(text: str) -> str:
            return get_display(arabic_reshaper.reshape(text))

        font_name = "Helvetica"
        for candidate in ["C:/Windows/Fonts/tahoma.ttf", "C:/Windows/Fonts/arial.ttf"]:
            if os.path.exists(candidate):
                try:
                    pdfmetrics.registerFont(TTFont("CustomPersianFont", candidate))
                    font_name = "CustomPersianFont"
                    break
                except Exception:
                    pass

        reader = PdfReader(io.BytesIO(pdf_bytes))
        writer = PdfWriter()
        total_pages = len(reader.pages)

        bx, by, bw, bh = 34, 26, A4[0] - 68, A4[1] - 52
        hy = A4[1] - 52
        fy = 44.0

        for idx, base_page in enumerate(reader.pages, 1):
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)

            # ۱. کادر بیرونی دور تا دور صفحه با گوشه‌های ملایم
            can.setStrokeColorRGB(0.7, 0.73, 0.76)  # #B2BABB
            can.setLineWidth(0.85)
            can.roundRect(bx, by, bw, bh, radius=4, stroke=1, fill=0)

            # ۲. خط حائل و محتوای هدر فیروزه‌ای رسمی
            can.setStrokeColorRGB(0.0, 0.48, 0.53)  # #007A87
            can.setLineWidth(1.2)
            can.line(bx, hy, bx + bw, hy)

            can.setFont(font_name, 8.5)
            can.setFillColorRGB(0.0, 0.48, 0.53)
            if chapter_title:
                can.drawRightString(bx + bw - 12, hy + 7, shape_persian(chapter_title))

            can.setFont(font_name, 7.5)
            can.setFillColorRGB(0.04, 0.45, 0.52)  # #0B7285
            can.drawString(bx + 12, hy + 7, shape_persian(institution or "دانشگاه بق"))

            # ۳. خط حائل و محتوای فوتر طوسی ملایم
            can.setStrokeColorRGB(0.8, 0.83, 0.86)  # #D5D8DC
            can.setLineWidth(0.75)
            can.line(bx, fy, bx + bw, fy)

            can.setFont(font_name, 8.5)
            can.setFillColorRGB(0.25, 0.3, 0.35)  # #2E4053
            can.drawCentredString(A4[0] / 2.0, 31, shape_persian(f"صفحه {idx} از {total_pages}"))

            can.setFont(font_name, 7.5)
            can.setFillColorRGB(0.5, 0.55, 0.6)
            can.drawRightString(bx + bw - 12, 31, shape_persian("مرور سریع آزمون"))
            can.drawString(bx + 12, 31, shape_persian(course_name or "فیزیوپاتولوژی بالینی"))

            can.save()
            packet.seek(0)

            stamp_pdf = PdfReader(packet)
            base_page.merge_page(stamp_pdf.pages[0])
            writer.add_page(base_page)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        final_written_path = safe_write_pdf(output_path, lambda f_out: writer.write(f_out))
        return final_written_path
    except Exception as e:
        print(f"[WARN] Persian BiDi page stamping skipped ({e}); writing PDF directly.")
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        final_written_path = safe_write_pdf(output_path, pdf_bytes)
        return final_written_path


def build_pdf_summary(data_or_html: Union[Dict[str, Any], str], output_pdf_path: str,
                      html_path: Optional[str] = None,
                      chapter_title: Optional[str] = None,
                      course_name: Optional[str] = None,
                      institution: str = "دانشگاه بق") -> str:
    """
    Renders publication-grade PDF from HTML via Playwright using system Edge/Chrome,
    with custom header templates and Persian BiDi page number stamping.
    Header Left explicitly displays institution ('دانشگاه بق').
    Accepts either data dict or html file path as first argument.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[INFO] Playwright not installed; skipping PDF generation.")
        return ""

    cleanup_temp_html = False
    if isinstance(data_or_html, str):
        if os.path.exists(data_or_html):
            html_path = data_or_html
            data = {}
        else:
            temp_html_path = output_pdf_path.replace(".pdf", "_temp.html")
            with open(temp_html_path, "w", encoding="utf-8") as f_tmp:
                f_tmp.write(data_or_html)
            html_path = temp_html_path
            cleanup_temp_html = True
            data = {}
    elif isinstance(data_or_html, dict):
        data = data_or_html
        if not html_path or not os.path.exists(html_path):
            temp_html_path = output_pdf_path.replace(".pdf", "_temp.html")
            build_html_summary(data, temp_html_path)
            html_path = temp_html_path
            cleanup_temp_html = True
    else:
        data = {}

    topic_title = chapter_title or data.get("topic_title", "خلاصه مبحث فیزیوپاتولوژی")
    inst_name = institution or data.get("institution") or "دانشگاه بق"
    c_name = course_name or data.get("organ_system") or "فیزیوپاتولوژی بالینی"

    header_html = f"""
    <div style="font-size: 8pt; font-family: 'Dubai', 'Vazirmatn', Tahoma, sans-serif; width: 100%;
                border-bottom: 1.5px solid #007A87; padding-bottom: 4px; margin: 0 13mm;
                display: flex; justify-content: space-between; direction: rtl;">
        <span style="font-weight: bold; color: #007A87;">{topic_title}</span>
        <span style="color: #0891B2;">{inst_name}</span>
    </div>
    """

    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    raw_pdf_bytes = None
    with sync_playwright() as p:
        browser = None
        # Try system Edge first, then Chrome, then default Chromium
        for channel in ["msedge", "chrome", None]:
            try:
                if channel:
                    browser = p.chromium.launch(channel=channel)
                else:
                    browser = p.chromium.launch()
                break
            except Exception:
                continue

        if not browser:
            print("[WARN] Could not launch system Edge or Chrome for PDF generation.")
            return ""

        page = browser.new_page()
        page.set_content(html_content, wait_until="networkidle")
        raw_pdf_bytes = page.pdf(
            format="A4",
            print_background=True,
            display_header_footer=False,
            margin={"top": "24mm", "bottom": "20mm", "left": "17mm", "right": "17mm"}
        )
        browser.close()

    final_pdf_path = output_pdf_path
    if raw_pdf_bytes:
        res = stamp_persian_page_numbers(raw_pdf_bytes, output_pdf_path, course_name=c_name, chapter_title=topic_title, institution=inst_name)
        if res and isinstance(res, str):
            final_pdf_path = res
        print(f"[OK] PDF Study Guide generated successfully: {final_pdf_path}")

    if cleanup_temp_html and os.path.exists(html_path):
        try:
            os.remove(html_path)
        except Exception:
            pass

    return final_pdf_path


# ---------------------------------------------------------------------------
# CLI Pipeline Orchestration
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Unified Medical Summary & Revision Pack Builder")
    parser.add_argument("--json", required=True, help="Path to structured input JSON containing summary data")
    parser.add_argument("--out-dir", default="dist", help="Output directory for generated files")
    parser.add_argument("--prefix", default="summary_pack", help="File prefix for outputs")
    parser.add_argument("--institution", default="دانشگاه بق", help="Institution name for header stamping")
    parser.add_argument("--no-pdf", action="store_true", help="Skip PDF generation")
    args = parser.parse_args()

    if not os.path.exists(args.json):
        print(f"[ERROR] Input JSON file not found: {args.json}")
        sys.exit(1)

    with open(args.json, "r", encoding="utf-8") as f:
        data = json.load(f)

    os.makedirs(args.out_dir, exist_ok=True)
    docx_file = os.path.join(args.out_dir, f"{args.prefix}.docx")
    html_file = os.path.join(args.out_dir, f"{args.prefix}.html")
    pdf_file = os.path.join(args.out_dir, f"{args.prefix}.pdf")
    md_file = os.path.join(args.out_dir, f"{args.prefix}.md")
    tsv_file = os.path.join(args.out_dir, f"{args.prefix}_anki.tsv")
    apkg_file = os.path.join(args.out_dir, f"{args.prefix}_anki.apkg")

    print(f"--- Starting Medical Summary Build: {data.get('topic_title', 'Summary')} ---")
    build_docx_summary(data, docx_file)
    build_html_summary(data, html_file)
    if not args.no_pdf:
        build_pdf_summary(data, pdf_file, html_path=html_file, institution=args.institution)
    build_markdown_summary(data, md_file)
    build_anki_tsv(data, tsv_file)
    build_anki_apkg(data, apkg_file)
    print("--- Build Finished Successfully! ---")


if __name__ == "__main__":
    main()
