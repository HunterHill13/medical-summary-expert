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
from typing import List, Dict, Any, Optional, Union, Set

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
COLOR_TEAL_BORDER       = "0D9488"            # Teal Border for Staging & Criteria Callouts
COLOR_TEAL_BG           = "F0FDFA"            # Light Teal Fill for Staging & Criteria Callouts
COLOR_CHECKLIST_NAVY    = "0369A1"            # Navy for Morning-of-Exam Capsule
COLOR_CHECKLIST_BG      = "FEFDE8"            # Warm Amber Tint for Morning Checklist

# Legacy compatibility aliases
COLOR_PRIMARY_BURGUNDY  = COLOR_PRIMARY_TEAL
COLOR_SECONDARY_TEAL    = COLOR_SECONDARY_CYAN
COLOR_TEXT_DARK         = COLOR_TEXT_CHARCOAL

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


def sanitize_slide_references(text: str) -> str:
    """Strips unwanted slide and page citation markers like [اسلاید ۱] or اسلاید ۲: from tables."""
    if not text:
        return ""
    t = str(text)
    t = re.sub(r'\[\s*اسلاید\s*\d+\s*\]', '', t)
    t = re.sub(r'\bاسلاید\s*\d+\s*[:\-–—]\s*', '', t)
    t = re.sub(r'\b(Slide|slide)\s*\d+\s*[:\-–—]?\s*', '', t)
    return t.strip()

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


def ensure_list(val: Any) -> list:
    """
    Safely converts any value to a list of items.
    If val is already a list, returns it. If None or empty, returns [].
    If val is a string, wraps it in [val] instead of character-iterating over it.
    Prevents the catastrophic String Iteration Bug in medical pipelines.
    """
    if val is None:
        return []
    if isinstance(val, list):
        return val
    if isinstance(val, (tuple, set)):
        return list(val)
    if isinstance(val, str):
        s = val.strip()
        return [s] if s else []
    return [val]


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
        # Color highlights (Rule 25)
        l = re.sub(r'==(.*?)==', r'<mark style="background-color: rgba(254, 249, 195, 0.55); border-bottom: 1.5px solid #FDE047; padding: 1px 4px; border-radius: 3px;">\1</mark>', l)
        l = re.sub(r'!!(.*?)!!', r'<mark style="background-color: rgba(255, 241, 242, 0.65); color: #881337; border-bottom: 1.5px solid #FDA4AF; padding: 1px 4px; border-radius: 3px; font-weight: 600;">\1</mark>', l)
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
                  is_rtl: bool = True, highlight_color: Optional[str] = None):
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

    if highlight_color == "yellow":
        rPr.append(parse_xml(r'<w:shd {} w:fill="FEF9C3"/>'.format(nsdecls('w'))))
    elif highlight_color == "red":
        rPr.append(parse_xml(r'<w:shd {} w:fill="FFF1F2"/>'.format(nsdecls('w'))))
        rPr.append(parse_xml(r'<w:color {} w:val="881337"/>'.format(nsdecls('w'))))

    return run


def add_r(p, text: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
          bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL,
          is_rtl: bool = True, highlight_color: Optional[str] = None):
    """Adds a run or parses inline markdown **bold**, yellow ==highlight== and red !!highlight!! into Word runs."""
    clean_text = sanitize_medical_text_symbols(str(text))
    if any(m in clean_text for m in ["==", "!!", "**"]):
        parts = re.split(r'(==.*?==|!!.*?!!|\*\*.*?\*\*)', clean_text)
        last_run = None
        for part in parts:
            if not part:
                continue
            if part.startswith("==") and part.endswith("==") and len(part) >= 4:
                last_run = _add_single_r(p, part[2:-2], font_name=font_name, size_pt=size_pt,
                                         bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl,
                                         highlight_color="yellow")
            elif part.startswith("!!") and part.endswith("!!") and len(part) >= 4:
                last_run = _add_single_r(p, part[2:-2], font_name=font_name, size_pt=size_pt,
                                         bold=True, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl,
                                         highlight_color="red")
            elif part.startswith("**") and part.endswith("**") and len(part) >= 4:
                last_run = _add_single_r(p, part[2:-2], font_name=font_name, size_pt=size_pt,
                                         bold=True, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl,
                                         highlight_color=highlight_color)
            else:
                last_run = _add_single_r(p, part, font_name=font_name, size_pt=size_pt,
                                         bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl,
                                         highlight_color=highlight_color)
        return last_run
    return _add_single_r(p, clean_text, font_name=font_name, size_pt=size_pt,
                         bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=is_rtl,
                         highlight_color=highlight_color)


def add_bidi_text(p, text: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
                  bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL,
                  highlight_color: Optional[str] = None):
    """
    Renders mixed Persian/English scientific text with strict BiDi isolation.
    Supports yellow ==highlight== and red !!highlight!! without extra text badges.
    """
    if not text:
        return

    text = sanitize_medical_text_symbols(text)

    # Handle explicit line breaks
    lines = text.split('\n')
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break()
        # Parse inline markdown **bold** and highlights ==...== and !!...!!
        parts = re.split(r'(==.*?==|!!.*?!!|\*\*.*?\*\*)', line)
        for part in parts:
            if not part:
                continue
            if part.startswith("==") and part.endswith("==") and len(part) >= 4:
                inner_text = part[2:-2]
                _render_bidi_line(p, inner_text, font_name=font_name, size_pt=size_pt,
                                  bold=bold, italic=italic, color_rgb=color_rgb, highlight_color="yellow")
            elif part.startswith("!!") and part.endswith("!!") and len(part) >= 4:
                inner_text = part[2:-2]
                _render_bidi_line(p, inner_text, font_name=font_name, size_pt=size_pt,
                                  bold=True, italic=italic, color_rgb=color_rgb, highlight_color="red")
            elif part.startswith("**") and part.endswith("**") and len(part) >= 4:
                inner_text = part[2:-2]
                _render_bidi_line(p, inner_text, font_name=font_name, size_pt=size_pt,
                                  bold=True, italic=italic, color_rgb=color_rgb, highlight_color=highlight_color)
            else:
                _render_bidi_line(p, part, font_name=font_name, size_pt=size_pt,
                                  bold=bold, italic=italic, color_rgb=color_rgb, highlight_color=highlight_color)


def _render_bidi_line(p, text_line: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
                      bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL,
                      highlight_color: Optional[str] = None):
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
                  bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=False, highlight_color=highlight_color)
        else:
            add_r(p, tok, font_name=font_name, size_pt=size_pt,
                  bold=bold, italic=italic, color_rgb=color_rgb, is_rtl=True, highlight_color=highlight_color)


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
        clean_h = sanitize_slide_references(str(h_text))
        add_r(p, clean_h, font_name=FONT_DUBAI, size_pt=9.0, bold=True,
              color_rgb=(0xFF, 0xFF, 0xFF), is_rtl=True)

    # Data Rows
    for row_idx, r_data in enumerate(valid_rows):
        trPr = table.rows[row_idx + 1]._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))

        bg_fill = COLOR_TABLE_ALT_ROW if row_idx % 2 == 1 else "FFFFFF"
        safe_row = r_data[:len(headers)] if isinstance(r_data, (list, tuple)) else [str(r_data)]
        for col_idx, cell_value in enumerate(safe_row):
            cell = table.cell(row_idx + 1, col_idx)
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            if bg_fill != "FFFFFF":
                set_cell_shading(cell, bg_fill)
            cell._tc.get_or_add_tcPr().append(parse_xml(
                r'<w:tcBorders {}><w:top w:val="none"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="E2E8F0"/><w:right w:val="single" w:sz="4" w:color="E2E8F0"/></w:tcBorders>'.format(nsdecls('w'))
            ))

            lines = str(cell_value).splitlines()
            p_first = cell.paragraphs[0]
            for l_idx, line in enumerate(lines):
                p = p_first if l_idx == 0 else cell.add_paragraph()
                set_p_rtl(p, space_before_pt=0.5, space_after_pt=1.5)
                l = line.strip()
                if not l:
                    continue
                clean_l = sanitize_slide_references(l)
                if not clean_l:
                    continue
                # In-cell badge recognition
                badge_color = None
                if any(k in clean_l for k in ["تله امتحانی", "🔴", "تله بزرگ"]):
                    badge_color = COLOR_RUBY_RED
                elif any(k in clean_l for k in ["خطای رایج", "خطای فاحش", "⚠️"]):
                    badge_color = COLOR_AMBER_ORANGE
                elif any(k in clean_l for k in ["استاندارد طلایی", "💎"]):
                    badge_color = COLOR_EMERALD_GREEN
                elif any(k in clean_l for k in ["نکته", "مهم", "📌"]):
                    badge_color = COLOR_PRIMARY_TEAL

                is_anchor = (col_idx == 0)
                if badge_color:
                    add_bidi_text(p, clean_l, font_name=FONT_DUBAI, size_pt=8.5,
                                  bold=True, color_rgb=badge_color)
                else:
                    add_bidi_text(p, clean_l, font_name=FONT_DUBAI, size_pt=8.5,
                                  bold=is_anchor, color_rgb=COLOR_PRIMARY_TEAL if is_anchor else COLOR_TEXT_CHARCOAL)

    # Set proportional column widths for A4 portrait
    col_widths = []
    if len(headers) == 3:
        col_widths = [Inches(1.5), Inches(2.3), Inches(2.5)]
    elif len(headers) == 4:
        col_widths = [Inches(1.3), Inches(1.8), Inches(1.5), Inches(1.7)]
    elif len(headers) == 5:
        col_widths = [Inches(1.1), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.3)]
    else:
        col_widths = [Inches(6.3 / max(1, len(headers)))] * len(headers)

    for c_idx, width in enumerate(col_widths):
        if c_idx < len(table.columns):
            table.columns[c_idx].width = width
            for row in table.rows:
                if c_idx < len(row.cells):
                    row.cells[c_idx].width = width

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
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        set_cell_shading(cell, COLOR_TABLE_HEADER_BG)
        cell._tc.get_or_add_tcPr().append(parse_xml(
            r'<w:tcBorders {}><w:top w:val="single" w:sz="6" w:color="78281F"/><w:bottom w:val="single" w:sz="12" w:color="78281F"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
        ))
        p = cell.paragraphs[0]
        set_p_rtl(p, space_before_pt=0, space_after_pt=0)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        clean_hdr = sanitize_slide_references(str(h_text))
        add_r(p, clean_hdr, font_name=FONT_DUBAI, size_pt=9.0, bold=True,
              color_rgb=(0xFF, 0xFF, 0xFF), is_rtl=True)

    # Data Rows
    for row_idx, r_data in enumerate(valid_rows):
        bg_fill = COLOR_TABLE_ALT_ROW if row_idx % 2 == 1 else "FFFFFF"
        # Slice r_data to match headers length to prevent IndexError on column overflow
        safe_row = r_data[:len(headers)] if isinstance(r_data, (list, tuple)) else [str(r_data)]
        for col_idx, cell_value in enumerate(safe_row):
            cell = table.cell(row_idx + 1, col_idx)
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            if bg_fill != "FFFFFF":
                set_cell_shading(cell, bg_fill)
            cell._tc.get_or_add_tcPr().append(parse_xml(
                r'<w:tcBorders {}><w:top w:val="none"/><w:bottom w:val="single" w:sz="4" w:color="E0E0E0"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
            ))
            p = cell.paragraphs[0]
            set_p_rtl(p, space_before_pt=1, space_after_pt=2)
            # Make the first column slightly bolder as the condition/drug anchor
            is_anchor = (col_idx == 0)
            clean_cell = sanitize_slide_references(str(cell_value))
            add_bidi_text(p, clean_cell, font_name=FONT_DUBAI, size_pt=8.5,
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



def enrich_mermaid_flowchart(code: str, title: str = "") -> str:
    """
    Enriches Mermaid flowchart code with vibrant, content-aware semantic 5-color classes:
    1. startNode (#E0F2FE blue border #0284C7): Triggers, initial presentations, entry points
    2. organNode (#E6FFFA mint border #0D9488): Organs, glands, hormones, syntheses
    3. decisionNode (#FEF3C7 amber border #D97706): Tests, evaluations, assays, differential dilemmas, '?'
    4. diseaseNode (#FFE4E6 rose border #E11D48): Pathologies, syndromes, abnormal cutoffs, emergencies, alert triggers
    5. actionNode (#DCFCE7 emerald border #16A34A): Treatments, normal results, therapeutic interventions, definitive outcomes
    """
    clean = clean_mermaid_code(code)
    if not clean or not str(clean).strip():
        return clean
    if "classDef" in clean:
        return clean

    classes_header = """
    classDef startNode fill:#E0F2FE,stroke:#0284C7,stroke-width:2.5px,color:#0369A1,rx:12px,ry:12px;
    classDef organNode fill:#E6FFFA,stroke:#0D9488,stroke-width:2.5px,color:#115E59,rx:12px,ry:12px;
    classDef decisionNode fill:#FEF3C7,stroke:#D97706,stroke-width:2.5px,color:#92400E,rx:12px,ry:12px;
    classDef diseaseNode fill:#FFE4E6,stroke:#E11D48,stroke-width:2.5px,color:#9F1239,rx:12px,ry:12px;
    classDef actionNode fill:#DCFCE7,stroke:#16A34A,stroke-width:2.5px,color:#15803D,rx:12px,ry:12px;"""

    node_pattern = re.compile(
        r'(?<![A-Za-z0-9_])([A-Za-z0-9_]+)\s*(\[[^\]]+\]|\{[^\}]+\}|\([^\)]+\)|\[\([^\)]+\)\]|\[\[[^\]]+\]\])'
    )
    matches = node_pattern.findall(clean)
    if not matches:
        return clean

    DECISION_KW = [
        'تست', 'سنجش', 'بررسی', 'آزمایش', 'غربالگری', 'اندازه‌گیری', 'اندازه گیری',
        'معیار', '؟', '?', 'تشخیص افتراقی', 'کتون', 'قند', 'pH', 'TSH', 'ACTH', 'سرم',
        'FT4', 'T3', 'سونوگرافی', 'اسمولالیته', 'بیوپسی', 'سی تی', 'MRI', 'الکترولیت'
    ]
    ACTION_KW = [
        'درمان', 'تجویز', 'جراحی', 'دارو', 'پاسخ', 'طبیعی', 'نرمال', 'رژیم', 'کنترل',
        'موفق', 'رد ', 'منفی', 'بهبود', 'تزریق', 'هیدراتاسیون', 'انسولین', 'استاتین'
    ]
    DISEASE_KW = [
        'بیماری', 'سندروم', 'سندرم', 'پاتولوژی', 'کوشینگ', 'آدیسون', 'هاشیموتو', 'گریوز',
        'تومور', 'آدنوم', 'کارسینوم', 'مقاومت', 'غیرطبیعی', 'مثبت', 'اختلال', 'نارسایی',
        'شوک', 'خطر', 'تیروتوکسیکوز', 'هایپرتیروئیدی', 'هیپوتیروئیدی', 'DKA', 'HHS',
        'مرگبار', 'کما', 'کتواسیدوز'
    ]
    ORGAN_KW = [
        'غده', 'هیپوفیز', 'تیروئید', 'آدرنال', 'پانکراس', 'کبد', 'کلیه', 'هورمون',
        'سنتز', 'ترشح', 'محور', 'پاراتیروئید', 'گناد', 'تخمدان', 'بیضه'
    ]
    START_KW = ['بیمار با', 'مراجعه', 'ارزیابی اولیه', 'ورودی', 'تظاهر اصلی', 'مشکوک']

    palette_cycle = ['organNode', 'decisionNode', 'diseaseNode', 'actionNode']
    class_assignments = []
    for idx, (nid, shape) in enumerate(matches):
        content = shape[1:-1]
        is_decision_shape = shape.startswith('{') and shape.endswith('}')

        if idx == 0 or any(k in content for k in START_KW):
            cls = 'startNode'
        elif any(k in content for k in ACTION_KW):
            cls = 'actionNode'
        elif is_decision_shape or any(k in content for k in DECISION_KW):
            cls = 'decisionNode'
        elif any(k in content for k in DISEASE_KW):
            cls = 'diseaseNode'
        elif any(k in content for k in ORGAN_KW):
            cls = 'organNode'
        else:
            cls = 'decisionNode' if is_decision_shape else (
                'actionNode' if idx == len(matches) - 1 else palette_cycle[idx % len(palette_cycle)]
            )
        class_assignments.append(f"    class {nid} {cls};")

    return clean + "\n" + classes_header + "\n" + "\n".join(class_assignments)


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

    # Enrich Mermaid code with vibrant, content-aware 5-color semantic classes
    enriched_code = enrich_mermaid_flowchart(clean_code, title=title)

    edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    exec_path = edge_exe if os.path.exists(edge_exe) else None

    html_template = """<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
    <meta charset="utf-8">
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
    <script>
        mermaid.initialize({
            startOnLoad: true,
            theme: 'base',
            themeVariables: {
                fontFamily: 'Vazirmatn, Dubai, Segoe UI, Tahoma, sans-serif',
                primaryColor: '#F0FDFA',
                primaryBorderColor: '#007A87',
                primaryTextColor: '#0F172A',
                lineColor: '#007A87',
                secondaryColor: '#FEF3C7',
                tertiaryColor: '#FFE4E6',
                edgeLabelBackground: '#ffffff',
                fontSize: '16px'
            }
        });
    </script>
    <style>
        body {
            background: #ffffff;
            margin: 0;
            padding: 20px;
            font-family: 'Vazirmatn', 'Dubai', 'Segoe UI', Tahoma, sans-serif;
            direction: rtl;
        }
        .card {
            border: 1.5px solid #007A87;
            border-radius: 12px;
            padding: 20px 24px;
            background: #FFFFFF;
            box-shadow: 0 4px 18px rgba(0, 122, 135, 0.1);
            width: 780px;
            max-width: 780px;
            box-sizing: border-box;
            margin: 0 auto;
            text-align: center;
        }
        .card-header {
            font-size: 16px;
            font-weight: bold;
            color: #FFFFFF;
            background: linear-gradient(135deg, #007A87 0%, #005F6A 100%);
            margin: -20px -24px 18px -24px;
            padding: 10px 18px;
            border-radius: 10px 10px 0 0;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            box-shadow: 0 2px 6px rgba(0, 122, 135, 0.2);
        }
        .mermaid {
            display: flex;
            justify-content: center;
        }
        .mermaid svg {
            max-width: 100% !important;
            height: auto !important;
        }
        .node rect, .node circle, .node polygon {
            rx: 10px !important;
            ry: 10px !important;
            filter: drop-shadow(0 4px 10px rgba(0, 122, 135, 0.15)) !important;
        }
        .edgePath path {
            stroke-width: 2.4px !important;
            stroke: #0284C7 !important;
        }
        .edgeLabel {
            font-weight: 600 !important;
            background: #FFFFFF !important;
            border-radius: 4px !important;
            padding: 2px 6px !important;
            border: 1px solid #E2E8F0 !important;
        }
        .cluster rect {
            rx: 12px !important;
            ry: 12px !important;
            stroke: #00A896 !important;
        }
    </style>
</head>
<body dir="rtl">
    <div class="card">
        <div class="card-header">🧬 __TITLE__</div>
        <div class="mermaid">
__CLEAN_CODE__
        </div>
    </div>
</body>
</html>"""
    html_content = html_template.replace("__TITLE__", title).replace("__CLEAN_CODE__", enriched_code)
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

def extract_covered_section_entities(data: Dict[str, Any]) -> Set[str]:
    """
    Extracts all entity and disease names covered in thematic tables (sections)
    to prevent cross-layer triplication across sections, modules, and matrices.
    """
    covered = set()
    for sec in (data.get("sections") or data.get("thematic_sections") or []):
        if not isinstance(sec, dict):
            continue
        title = sec.get("title") or sec.get("section_title") or ""
        if title:
            clean_title = re.sub(r"[^\w\s]", " ", title, flags=re.UNICODE).lower()
            for token in clean_title.split():
                if len(token) >= 3 and token not in {"جدول", "تماتیک", "بررسی", "مقایسه", "جامع", "سندرم", "بخش", "رویکرد"}:
                    covered.add(token)
            covered.add(title.strip().lower())
        for row in (sec.get("rows") or []):
            if isinstance(row, (list, tuple)) and row:
                first_cell = str(row[0]).strip().lower()
                clean_cell = re.sub(r"[^\w\s]", " ", first_cell, flags=re.UNICODE).lower()
                for token in clean_cell.split():
                    if len(token) >= 3 and token not in {"جدول", "تماتیک", "بررسی", "مقایسه", "جامع", "سندرم", "بخش", "رویکرد"}:
                        covered.add(token)
                covered.add(first_cell)
    return covered


def is_module_covered_in_sections(mod: Dict[str, Any], covered_entities: Set[str]) -> bool:
    """
    Returns True if the module's primary clinical subject is already covered in a thematic table.
    """
    if not covered_entities:
        return False
    name = (mod.get("name") or "").strip().lower()
    name_en = (mod.get("name_en") or "").strip().lower()
    if not name and not name_en:
        return False

    if name in covered_entities or name_en in covered_entities:
        return True

    clean_name = re.sub(r"[^\w\s]", " ", name, flags=re.UNICODE).lower()
    for token in clean_name.split():
        if len(token) >= 3 and token not in {"جدول", "تماتیک", "بررسی", "مقایسه", "جامع", "سندرم", "بخش", "رویکرد"} and token in covered_entities:
            return True

    if name_en:
        clean_en = re.sub(r"[^\w\s]", " ", name_en).lower()
        for token in clean_en.split():
            if len(token) >= 4 and token in covered_entities:
                return True

    return False


def ensure_master_matrices(data: Dict[str, Any]):
    """
    Ensures both master_pharmacology_matrix and last_minute_crash_sheet
    are synthesized if missing, ensuring 100% content parity across Word and Markdown.
    Skips duplicate pharmacology matrix if already comprehensively covered in thematic sections.
    """
    sections = data.get("sections") or data.get("thematic_sections") or []
    modules = data.get("modules") or []
    # In pure table mode (sections present and no modules), do NOT synthesize trailing matrices!
    if sections and not modules:
        return

    # 1. Master Pharmacology Matrix
    if not data.get("master_pharmacology_matrix"):
        sections = data.get("sections") or data.get("thematic_sections") or []
        has_pharma_section = any(
            any(k in (sec.get("title") or "") for k in ["دارو", "فارماکولوژی", "درمان دارویی", "Pharmacology"])
            for sec in sections if isinstance(sec, dict)
        )
        if not has_pharma_section:
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
            for a in ensure_list(mod.get("instructor_exam_alerts")):
                if a and str(a).strip():
                    bullets.append(f"تأکید مستقیم استاد ({m_name}): {str(a).strip()}")
        # Priority 2: Absolute Contraindications & Lethal Withholds
        for mod in modules:
            if not isinstance(mod, dict):
                continue
            m_name = mod.get("name", "")
            for c in ensure_list(mod.get("contraindications_and_warnings")):
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
            for p in ensure_list(mod.get("high_yield_pearls")):
                if p and str(p).strip():
                    bullets.append(f"نکته های‌یلد ({m_name}): {str(p).strip()}")
            for pv in ensure_list(mod.get("pivot_points") or mod.get("high_yield_decision_points")):
                if pv and str(pv).strip():
                    bullets.append(f"نقطه انشعاب بالینی ({m_name}): {str(pv).strip()}")
            for cr in ensure_list(mod.get("clinical_relevance")):
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



def to_fa_digits(val: Any) -> str:
    """Converts ASCII digits to Persian digits."""
    return str(val).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))

def extract_chronological_flow(data: Dict[str, Any], chapter_badge: str = "") -> List[Dict[str, Any]]:
    """
    Extracts and pairs sections and modules into a strictly chronological list of topic items.
    Enforces exact chronological sequence of the lecture and prevents the table/pearl segregation
    defect where all tables appear first and all modules second.
    """
    if "topics" in data and isinstance(data["topics"], list):
        flow = []
        for t in data["topics"]:
            t_copy = dict(t)
            t_copy["chapter_badge"] = chapter_badge
            flow.append(t_copy)
        return flow

    raw_sections = data.get("sections") or data.get("thematic_sections") or []
    sections = [s for s in raw_sections if isinstance(s, dict)]
    raw_modules = data.get("modules") or []
    modules = [m for m in raw_modules if isinstance(m, dict)]

    if not sections and not modules:
        return []
    if not sections:
        return [{"type": "module_only", "module": m, "chapter_badge": chapter_badge} for m in modules]
    if not modules:
        return [{"type": "section_only", "section": s, "chapter_badge": chapter_badge} for s in sections]

    matched_modules = set()
    flow = []

    def normalize_name(s: str) -> str:
        s = re.sub(r'[*_#•\(\)\[\]]', ' ', str(s))
        s = re.sub(r'بخش\s*\d+\s*:\s*', '', s)
        s = re.sub(r'سندروم\s+', 'سندرم ', s)
        s = s.replace('\u200c', '')
        return s.strip().lower()

    # Pass 1: Try exact or semantic match between section and module
    for sec_idx, sec in enumerate(sections):
        sec_title = sec.get("title") or sec.get("section_title") or ""
        norm_sec = normalize_name(sec_title)

        best_mod_idx = None
        # Explicit ID or Ref
        for m_idx, mod in enumerate(modules):
            if m_idx in matched_modules:
                continue
            if sec.get("section_id") and sec.get("section_id") == mod.get("section_id"):
                best_mod_idx = m_idx
                break
            if sec.get("module_ref") and sec.get("module_ref") == mod.get("name"):
                best_mod_idx = m_idx
                break

        # Title / Name containment
        if best_mod_idx is None:
            for m_idx, mod in enumerate(modules):
                if m_idx in matched_modules:
                    continue
                m_name = normalize_name(mod.get("name", ""))
                m_name_en = mod.get("name_en", "").lower().strip()
                if m_name and (m_name in norm_sec or norm_sec in m_name):
                    best_mod_idx = m_idx
                    break
                if m_name_en and (m_name_en in norm_sec.lower() or m_name_en in sec_title.lower()):
                    best_mod_idx = m_idx
                    break

        # Check row cells content
        if best_mod_idx is None:
            for m_idx, mod in enumerate(modules):
                if m_idx in matched_modules:
                    continue
                m_name = normalize_name(mod.get("name", ""))
                if m_name and any(m_name in normalize_name(str(cell)) for r in sec.get("rows", []) for cell in (r if isinstance(r, list) else [r])):
                    best_mod_idx = m_idx
                    break

        # Positional alignment fallback if counts match
        if best_mod_idx is None and len(sections) == len(modules):
            if sec_idx not in matched_modules:
                best_mod_idx = sec_idx

        if best_mod_idx is not None:
            matched_modules.add(best_mod_idx)
            flow.append({
                "type": "topic",
                "section": sec,
                "module": modules[best_mod_idx],
                "chapter_badge": chapter_badge
            })
        else:
            flow.append({
                "type": "section_only",
                "section": sec,
                "chapter_badge": chapter_badge
            })

    # Append remaining modules
    for m_idx, mod in enumerate(modules):
        if m_idx not in matched_modules:
            flow.append({
                "type": "module_only",
                "module": mod,
                "chapter_badge": chapter_badge
            })

    return flow


def render_docx_table_of_contents(doc, chapters: List[Dict[str, Any]], title: str = "📑 فهرست فصول و مباحث درسنامه جامع"):
    """Renders an exam-grade Table of Contents table at the top of multi-handout Word documents."""
    p_toc_head = doc.add_paragraph()
    set_p_rtl(p_toc_head, space_before_pt=12, space_after_pt=6)
    add_r(p_toc_head, title, font_name=FONT_DUBAI, size_pt=14.0, bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)

    headers = ["شماره جزوه", "عنوان جزوه و فایل منبع", "سرفصل‌ها و مباحث کلیدی"]
    rows = []
    for idx, ch in enumerate(chapters, 1):
        ch_title = ch.get("title") or ch.get("topic_title") or f"جزوه {idx}"
        src = ch.get("source_file") or ch.get("file_name") or ""
        src_text = f"\n📄 فایل منبع: {src}" if src else ""
        col2 = f"• **{ch_title}**{src_text}"
        topics = []
        for sec in (ch.get("sections") or ch.get("thematic_sections") or []):
            t_name = sec.get("title") or sec.get("section_title") or ""
            if t_name and t_name not in topics:
                topics.append(t_name)
        for mod in (ch.get("modules") or []):
            m_name = mod.get("name", "")
            if m_name and not any(m_name in t for t in topics):
                topics.append(m_name)
        if not topics:
            topics = ["مباحث جامع بالینی"]
        col3 = "\n".join([f"• {t}" for t in topics[:6]])
        rows.append([f"📚 جزوه {to_fa_digits(idx)}]", col2, col3])

    add_thematic_table(doc, headers=headers, rows=rows)
    p_sp = doc.add_paragraph()
    set_p_rtl(p_sp, space_before_pt=0, space_after_pt=8)


def render_docx_section_block(doc, sec: Dict[str, Any], sec_idx: int = 1, chapter_badge: str = "", out_dir: str = ""):
    """Renders a single thematic section (table, flowchart, callouts) into Word."""
    if not isinstance(sec, dict):
        return
    raw_title = sec.get("title") or sec.get("section_title") or f"بخش {sec_idx}: مبحث بالینی"
    if chapter_badge and chapter_badge not in raw_title:
        sec_title = f"{chapter_badge} {raw_title}"
    else:
        sec_title = raw_title
    sec_headers = sec.get("headers") or []
    sec_rows = sec.get("rows") or []

    p_sec = doc.add_paragraph()
    set_p_rtl(p_sec, space_before_pt=14, space_after_pt=4)
    add_r(p_sec, f"■ {sec_title}", font_name=FONT_DUBAI, size_pt=10.5, bold=True,
          color_rgb=COLOR_PRIMARY_TEAL, is_rtl=True)

    if sec_headers and sec_rows:
        add_thematic_table(doc, headers=sec_headers, rows=sec_rows)

    sec_flowchart = sec.get("flowchart") or sec.get("mermaid_code")
    if sec_flowchart:
        target_dir = out_dir if out_dir else os.path.join(os.getcwd(), "خلاصه")
        diag_dir = os.path.join(target_dir, "diagrams")
        os.makedirs(diag_dir, exist_ok=True)
        m_hash = hashlib.md5(sec_flowchart.strip().encode("utf-8")).hexdigest()[:10]
        candidate_png = os.path.join(diag_dir, f"diag_sec_{m_hash}.png")
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


def render_docx_module_block(doc, mod: Dict[str, Any], mod_idx: int = 1, covered_entities: Optional[Set[str]] = None, chapter_badge: str = ""):
    """Renders a single high-yield module/illness script into Word."""
    if not isinstance(mod, dict):
        return
    covered_entities = covered_entities or set()
    raw_mod_name = mod.get("name", f"مبحث شماره {mod_idx}")
    if chapter_badge and chapter_badge not in raw_mod_name:
        mod_name = f"{chapter_badge} {raw_mod_name}"
    else:
        mod_name = raw_mod_name
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
    is_covered = is_module_covered_in_sections(mod, covered_entities)

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

    # Cross-Layer Deduplication Notice (v3.9.3)
    if is_covered:
        p_cov = doc.add_paragraph()
        set_p_rtl(p_cov, space_before_pt=3, space_after_pt=5)
        add_r(p_cov, "⚡ ارجاع به جدول تماتیک: ", font_name=FONT_DUBAI, size_pt=10.5, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
        add_bidi_text(p_cov, "جزئیات تشخیصی، کات‌آف‌ها، تظاهرات و درمان در جدول تماتیک فوق به صورت مقایسه‌ای ارائه شده است. در ادامه تنها نکات دامی آزمون، مدیریت اورژانس و هشدارهای اختصاصی درج می‌گردد.",
                      font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_MUTED_GRAY)

    # Module-level Thematic Tables (v3.9.0 Table-First Hybrid)
    if mod.get("tables") and isinstance(mod["tables"], list):
        for t_item in mod["tables"]:
            if isinstance(t_item, dict):
                add_thematic_table(doc, headers=t_item.get("headers", []), rows=t_item.get("rows", []), title=t_item.get("title"))
    elif mod.get("thematic_table") and isinstance(mod["thematic_table"], dict):
        t_item = mod["thematic_table"]
        add_thematic_table(doc, headers=t_item.get("headers", []), rows=t_item.get("rows", []), title=t_item.get("title"))

    if is_covered:
        pass  # Suppress duplicate boilerplate paragraphs; jump straight to high-yield callouts!
    # -------------------------------------------------------------------
    # Archetype 1: Foundational Physiology & Regulatory Loops
    # -------------------------------------------------------------------
    elif is_foundation:
        roles = mod.get("physiological_role") or mod.get("anatomical_physiological_basis")
        if roles:
            p_rh = doc.add_paragraph()
            set_p_rtl(p_rh, space_before_pt=6, space_after_pt=2)
            add_r(p_rh, "۱. نقش و تنظیم فیزیولوژیک (Physiological Role & Homeostasis)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
            roles_list = roles if isinstance(roles, list) else [roles]
            for r in roles_list:
                p_r = doc.add_paragraph()
                set_p_rtl(p_r, space_before_pt=1, space_after_pt=2)
                add_r(p_r, f"• {r}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

        f_casc = mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback")
        if f_casc and isinstance(f_casc, list):
            add_callout_box(
                doc,
                title="🔄 آبشار تنظیمی و فیدبک هورمونی (Feedback & Signaling Cascade)",
                content_items=f_casc,
                border_color="007A87",
                header_bg="E6FFFA",
                body_bg="FFFFFF",
                icon="🔄"
            )

        if mod.get("baseline_tests") and isinstance(mod["baseline_tests"], list):
            p_th = doc.add_paragraph()
            set_p_rtl(p_th, space_before_pt=6, space_after_pt=2)
            add_r(p_th, "۳. سنجش آزمایشگاهی و مقادیر پایه (Baseline & Dynamic Tests)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
            for t in mod["baseline_tests"]:
                p_t = doc.add_paragraph()
                set_p_rtl(p_t, space_before_pt=1, space_after_pt=2)
                add_r(p_t, f"• {t}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

        if mod.get("clinical_relevance") and isinstance(mod["clinical_relevance"], list):
            p_ch = doc.add_paragraph()
            set_p_rtl(p_ch, space_before_pt=6, space_after_pt=2)
            add_r(p_ch, "۴. ارتباط بالینی و عوارض اولیه اختلال (Clinical Correlate)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
            for cr in mod["clinical_relevance"]:
                p_c = doc.add_paragraph()
                set_p_rtl(p_c, space_before_pt=1, space_after_pt=2)
                add_r(p_c, f"• {cr}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

    # -------------------------------------------------------------------
    # Archetype 2: Clinical Approach & Diagnostic Algorithms
    # -------------------------------------------------------------------
    elif is_approach:
        inits = mod.get("initial_assessment") or mod.get("triage_and_stabilization")
        if inits:
            p_ih = doc.add_paragraph()
            set_p_rtl(p_ih, space_before_pt=6, space_after_pt=2)
            add_r(p_ih, "۱. ارزیابی اولیه و پایداری همودینامیک (Triage & Initial Workup)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
            inits_list = inits if isinstance(inits, list) else [inits]
            for init in inits_list:
                p_i = doc.add_paragraph()
                set_p_rtl(p_i, space_before_pt=1, space_after_pt=2)
                add_r(p_i, f"• {init}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

        steps = mod.get("step_by_step_algorithm") or mod.get("sequential_steps")
        if steps and isinstance(steps, list):
            add_callout_box(
                doc,
                title="🧭 الگوریتم گام‌به‌گام تصمیم‌گیری بالینی (Step-by-Step Clinical Decision)",
                content_items=steps,
                border_color="007A87",
                header_bg="E6FFFA",
                body_bg="FFFFFF",
                icon="🧭"
            )

        pivots = mod.get("pivot_points") or mod.get("high_yield_decision_points")
        if pivots and isinstance(pivots, list):
            p_ph = doc.add_paragraph()
            set_p_rtl(p_ph, space_before_pt=6, space_after_pt=2)
            add_r(p_ph, "۳. نقاط انشعاب کلیدی و کات‌آف‌ها (Clinical Pivot Points & Cutoffs)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
            for pv in pivots:
                p_p = doc.add_paragraph()
                set_p_rtl(p_p, space_before_pt=1, space_after_pt=2)
                add_r(p_p, f"• {pv}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

        actions = mod.get("action_plan") or mod.get("first_line_management")
        if actions and isinstance(actions, list):
            p_ah = doc.add_paragraph()
            set_p_rtl(p_ah, space_before_pt=6, space_after_pt=2)
            add_r(p_ah, "۴. پلن اقدام بالینی و مداخله (Clinical Action Plan)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
            for act in actions:
                p_a = doc.add_paragraph()
                set_p_rtl(p_a, space_before_pt=1, space_after_pt=2)
                add_r(p_a, f"• {act}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

    # -------------------------------------------------------------------
    # Archetype 3: Disease Illness Script (Full Standard)
    # -------------------------------------------------------------------
    else:
        # Part 1: Definition & Etiology
        p_h2_1 = doc.add_paragraph()
        set_p_rtl(p_h2_1, space_before_pt=8, space_after_pt=2)
        add_r(p_h2_1, "۱. تعریف و سبب‌شناسی (Definition & Etiology)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
        defs = mod.get("definition_and_etiology")
        if isinstance(defs, str):
            defs = [defs]
        for b in (defs or []):
            p_b = doc.add_paragraph()
            set_p_rtl(p_b, space_before_pt=1, space_after_pt=2)
            add_r(p_b, f"• {b}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

        risk_factors = mod.get("risk_factors")
        if risk_factors:
            risk_items = risk_factors if isinstance(risk_factors, list) else [risk_factors]
            for r in risk_items:
                p_r = doc.add_paragraph()
                set_p_rtl(p_r, space_before_pt=1, space_after_pt=2)
                add_r(p_r, f"• ریسک‌فاکتورها: {r}", font_name=FONT_DUBAI, size_pt=9.5, color_rgb=COLOR_MUTED_GRAY, is_rtl=True)

        # Part 2: Pathophysiology Cascade (Callout Box)
        p_casc = mod.get("pathophysiology_cascade")
        if p_casc:
            items = p_casc if isinstance(p_casc, list) else [p_casc]
            add_callout_box(
                doc,
                title="⚙️ آبشار پاتوفیزیولوژی و مکانیسم بیماری (Pathophysiology Cascade)",
                content_items=items,
                border_color="007A87",
                header_bg="E6FFFA",
                body_bg="FFFFFF",
                icon="⚙️"
            )

        # Part 3: Clinical Presentation
        if mod.get("clinical_presentation"):
            p_h2_3 = doc.add_paragraph()
            set_p_rtl(p_h2_3, space_before_pt=8, space_after_pt=2)
            add_r(p_h2_3, "۳. تظاهرات بالینی، نشانه‌ها و معاینه فیزیکی (Clinical Presentation)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

            pres_items = mod["clinical_presentation"]
            if isinstance(pres_items, str):
                pres_items = [pres_items]
            for item in pres_items:
                p_p = doc.add_paragraph()
                set_p_rtl(p_p, space_before_pt=1, space_after_pt=2)
                add_r(p_p, f"• {item}", font_name=FONT_DUBAI, size_pt=10.0, color_rgb=COLOR_TEXT_DARK, is_rtl=True)

        # Diagnostic Criteria & Staging
        if mod.get("diagnostic_criteria") and isinstance(mod["diagnostic_criteria"], list):
            add_callout_box(
                doc,
                title="📋 معیارهای تشخیصی و گایدلاین‌ها (Diagnostic Criteria & Guidelines)",
                content_items=mod["diagnostic_criteria"],
                border_color="007A87",
                header_bg="F0FDFA",
                body_bg="FFFFFF",
                icon="📋"
            )

        if mod.get("staging_and_classification") and isinstance(mod["staging_and_classification"], list):
            add_callout_box(
                doc,
                title="📊 استیجینگ و طبقه‌بندی شدت (Staging & Severity Classification)",
                content_items=mod["staging_and_classification"],
                border_color="007A87",
                header_bg="F0FDFA",
                body_bg="FFFFFF",
                icon="📊"
            )

        # Part 4: Diagnostics
        if mod.get("diagnostics") and isinstance(mod["diagnostics"], dict):
            diag = mod["diagnostics"]
            p_h2_4 = doc.add_paragraph()
            set_p_rtl(p_h2_4, space_before_pt=8, space_after_pt=2)
            add_r(p_h2_4, "۴. رویکرد تشخیصی و تست‌های پاراکلینیک (Diagnostic Workup)", font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

            if diag.get("best_initial_test"):
                p_init = doc.add_paragraph()
                set_p_rtl(p_init, space_before_pt=1, space_after_pt=2)
                add_r(p_init, "• اقدام تشخیصی نخستین (Best Initial Test): ", font_name=FONT_DUBAI, size_pt=10.0, bold=True, color_rgb=COLOR_PRIMARY_TEAL, is_rtl=True)
                add_r(p_init, str(diag['best_initial_test']), font_name=FONT_DUBAI, size_pt=10.0, is_rtl=True)

            if diag.get("gold_standard"):
                p_gold = doc.add_paragraph()
                set_p_rtl(p_gold, space_before_pt=1, space_after_pt=2)
                add_r(p_gold, "• استاندارد طلایی و قطعی (Gold Standard / Confirmatory): ", font_name=FONT_DUBAI, size_pt=10.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                add_r(p_gold, str(diag['gold_standard']), font_name=FONT_DUBAI, size_pt=10.0, is_rtl=True)

            labs = diag.get("key_laboratory_findings") or diag.get("key_findings")
            if labs:
                labs_list = labs if isinstance(labs, list) else [labs]
                for lab in labs_list:
                    p_l = doc.add_paragraph()
                    set_p_rtl(p_l, space_before_pt=1, space_after_pt=2)
                    add_r(p_l, f"• یافته آزمایشگاهی کلیدی: {lab}", font_name=FONT_DUBAI, size_pt=9.5, is_rtl=True)

        # Part 5: Comprehensive Management (Medical, Surgical & Conservative Rx)
        has_pharma = bool(mod.get("pharmacology_and_management"))
        raw_surg = mod.get("surgical_and_interventions") or mod.get("surgical_and_interventional_indications")
        has_surg = bool(raw_surg)
        raw_life = mod.get("lifestyle_and_followup") or mod.get("follow_up_and_monitoring") or mod.get("lifestyle_and_conservative")
        has_life = bool(raw_life)

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
                    drug_class = rx.get("drug_class") or rx.get("class") or rx.get("agent") or rx.get("drug") or "-"
                    mechanism = rx.get("mechanism_of_action") or rx.get("mechanism") or rx.get("moa") or "-"
                    pearl = rx.get("pearl") or rx.get("exam_pearl") or rx.get("high_yield_pearl") or ""
                    add_r(p_rx, f"💊 داروی انتخابی ({drug_class}): ", font_name=FONT_DUBAI, size_pt=10.5,
                          bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                    text_desc = f"مکانیسم: {mechanism}"
                    if pearl:
                        text_desc += f" | نکته: {pearl}"
                    add_bidi_text(p_rx, text_desc, font_name=FONT_DUBAI, size_pt=10.5)

            if has_surg:
                surg_items = raw_surg if isinstance(raw_surg, list) else [raw_surg]
                p_sh = doc.add_paragraph()
                set_p_rtl(p_sh, space_before_pt=4, space_after_pt=2)
                add_r(p_sh, "• اندیکاسیون‌های مداخله و جراحی (Surgical & Interventional Indications):", font_name=FONT_DUBAI, size_pt=10.0, bold=True, color_rgb=COLOR_PRIMARY_TEAL, is_rtl=True)
                for s in surg_items:
                    p_s = doc.add_paragraph()
                    set_p_rtl(p_s, space_before_pt=1, space_after_pt=2)
                    add_r(p_s, f"  - {s}", font_name=FONT_DUBAI, size_pt=9.5, is_rtl=True)

            if has_life:
                life_items = raw_life if isinstance(raw_life, list) else [raw_life]
                p_lh = doc.add_paragraph()
                set_p_rtl(p_lh, space_before_pt=4, space_after_pt=2)
                add_r(p_lh, "• اصلاح سبک زندگی، پیگیری و پایش (Lifestyle & Follow-up):", font_name=FONT_DUBAI, size_pt=10.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                for l in life_items:
                    p_l = doc.add_paragraph()
                    set_p_rtl(p_l, space_before_pt=1, space_after_pt=2)
                    add_r(p_l, f"  - {l}", font_name=FONT_DUBAI, size_pt=9.5, is_rtl=True)

        # Complications & Prognosis
        if mod.get("complications_and_prognosis"):
            p_h2_comp = doc.add_paragraph()
            set_p_rtl(p_h2_comp, space_before_pt=6, space_after_pt=2)
            add_r(p_h2_comp, "• عوارض و پیش‌آگهی (Complications & Prognosis):", font_name=FONT_DUBAI, size_pt=10.0, bold=True, color_rgb=COLOR_PRIMARY_TEAL, is_rtl=True)
            comps = mod["complications_and_prognosis"]
            if isinstance(comps, str):
                comps = [comps]
            for comp in comps:
                p_c = doc.add_paragraph()
                set_p_rtl(p_c, space_before_pt=1, space_after_pt=2)
                add_r(p_c, f"  - {comp}", font_name=FONT_DUBAI, size_pt=9.5, is_rtl=True)

    # -------------------------------------------------------------------
    # High-Yield Callout Boxes (Universally rendered across all archetypes)
    # -------------------------------------------------------------------
    pearls = mod.get("high_yield_pearls") or mod.get("clinical_pearls") or []
    if pearls:
        pearls_list = pearls if isinstance(pearls, list) else [pearls]
        add_callout_box(
            doc,
            title=f"💎 مرواریدهای بالینی و نکات کلیدی آزمون: {mod_name}",
            content_items=pearls_list,
            border_color="007A87",
            header_bg="E6FFFA",
            body_bg="FFFFFF",
            icon="💎"
        )

    traps = mod.get("exam_traps") or mod.get("common_pitfalls") or []
    if traps:
        traps_list = traps if isinstance(traps, list) else [traps]
        add_callout_box(
            doc,
            title=f"⚠️ تله‌های تستی و خطاهای رایج پزشکی: {mod_name}",
            content_items=traps_list,
            border_color="B91C1C",
            header_bg="FEF2F2",
            body_bg="FFFFFF",
            icon="⚠️"
        )

    emerg = mod.get("emergency_management") or mod.get("emergency_action") or []
    if emerg:
        emerg_list = emerg if isinstance(emerg, list) else [emerg]
        add_callout_box(
            doc,
            title=f"🚨 اقدامات اورژانسی و پروتکل کرایزیس: {mod_name}",
            content_items=emerg_list,
            border_color="C0392B",
            header_bg="FDEDEC",
            body_bg="FFFFFF",
            icon="🚨"
        )

    spec = mod.get("special_populations") or mod.get("special_considerations") or []
    if spec:
        spec_list = spec if isinstance(spec, list) else [spec]
        add_callout_box(
            doc,
            title=f"👥 ملاحظات جمعیت‌های ویژه (بارداری، اطفال، سالمندان، CKD): {mod_name}",
            content_items=spec_list,
            border_color="0284C7",
            header_bg="F0F9FF",
            body_bg="FFFFFF",
            icon="👥"
        )

    emphasis = mod.get("instructor_emphasis") or mod.get("prof_notes") or []
    if emphasis:
        emp_list = emphasis if isinstance(emphasis, list) else [emphasis]
        add_callout_box(
            doc,
            title=f"⭐ تأکید ویژه استاد در تدریس: {mod_name}",
            content_items=emp_list,
            border_color="D97706",
            header_bg="FFFBEB",
            body_bg="FFFFFF",
            icon="⭐"
        )

    # Diagnostic & Physiological Formulas
    formulas = mod.get("formulas") or mod.get("clinical_formulas") or []
    if formulas:
        f_list = formulas if isinstance(formulas, list) else [formulas]
        for f_item in f_list:
            if isinstance(f_item, dict):
                f_name = f_item.get("name", f"فرمول بالینی {mod_name}")
                f_code = f_item.get("latex") or f_item.get("code") or f_item.get("formula", "")
                f_cap = f_item.get("interpretation") or f_item.get("notes", "")
                add_formula_card(doc, f_code, f_name, f_cap)
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


def render_docx_flow_item(doc, item: Dict[str, Any], idx: int, covered_entities: Set[str], out_dir: str = ""):
    """Renders a chronological flow item (topic, section_only, or module_only) into Word."""
    i_type = item.get("type")
    c_badge = item.get("chapter_badge", "")
    if i_type == "topic":
        if item.get("section"):
            render_docx_section_block(doc, item["section"], sec_idx=idx, chapter_badge=c_badge, out_dir=out_dir)
        if item.get("module"):
            render_docx_module_block(doc, item["module"], mod_idx=idx, covered_entities=covered_entities, chapter_badge=c_badge)
    elif i_type == "section_only":
        render_docx_section_block(doc, item["section"], sec_idx=idx, chapter_badge=c_badge, out_dir=out_dir)
    elif i_type == "module_only":
        render_docx_module_block(doc, item["module"], mod_idx=idx, covered_entities=covered_entities, chapter_badge=c_badge)


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

    chapters = data.get("chapters")
    is_multi = bool(data.get("is_multi_chapter") or (chapters and len(chapters) > 1))

    if is_multi and chapters:
        render_docx_table_of_contents(doc, chapters)
        for ch_idx, ch in enumerate(chapters, 1):
            doc.add_page_break()
            ch_title = ch.get("title") or ch.get("topic_title") or f"جزوه {ch_idx}"
            src_file = ch.get("source_file") or ch.get("file_name") or "منبع نامشخص"
            ch_organ = ch.get("organ_system") or organ_system
            add_callout_box(
                doc,
                title=f"📚 جزوه {to_fa_digits(ch_idx)}: {ch_title}",
                content_items=[
                    f"📄 فایل منبع تدریس: {src_file}",
                    f"🏛️ کورس آموزشی: {ch_organ}  |  🎯 مقطع: فیزیوپاتولوژی بالینی"
                ],
                border_color="007A87",
                header_bg="E6FFFA",
                body_bg="FFFFFF",
                icon="📚"
            )
            ch_flow = extract_chronological_flow(ch, chapter_badge=f"[جزوه {to_fa_digits(ch_idx)}]")
            ch_covered = extract_covered_section_entities(ch)
            for it_idx, item in enumerate(ch_flow, 1):
                render_docx_flow_item(doc, item, it_idx, ch_covered)
    else:
        flow = extract_chronological_flow(data, chapter_badge="")
        covered_entities = extract_covered_section_entities(data)
        for it_idx, item in enumerate(flow, 1):
            render_docx_flow_item(doc, item, it_idx, covered_entities)

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



def render_markdown_table_of_contents(lines: List[str], chapters: List[Dict[str, Any]], title: str = "📑 فهرست فصول و مباحث درسنامه جامع"):
    """Renders GitHub/Obsidian-compatible Table of Contents for multi-handout bundles."""
    lines.append(f"# {title}\n")
    lines.append("| شماره جزوه | عنوان جزوه / سرفصل | فایل منبع | تعداد مباحث |")
    lines.append("| :--- | :--- | :--- | :--- |")
    for idx, ch in enumerate(chapters, 1):
        ch_title = ch.get("title") or ch.get("topic_title") or f"جزوه {idx}"
        src = ch.get("source_file") or ch.get("file_name") or "منبع نامشخص"
        n_topics = len(ch.get("sections", []) or ch.get("thematic_sections", [])) + len(ch.get("modules", []))
        slug = re.sub(r'[^\\w\\s-]', '', ch_title).strip().replace(' ', '-')
        lines.append(f"| جزوه {to_fa_digits(idx)} | [📚 {ch_title}](#{slug}) | `{src}` | {max(1, n_topics)} مبحث |")
    lines.append("\n")

    for idx, ch in enumerate(chapters, 1):
        ch_title = ch.get("title") or ch.get("topic_title") or f"جزوه {idx}"
        src = ch.get("source_file") or ch.get("file_name") or ""
        slug = re.sub(r'[^\\w\\s-]', '', ch_title).strip().replace(' ', '-')
        src_badge = f" (منبع: `{src}`)" if src else ""
        lines.append(f"- **[📚 {ch_title}](#{slug})**{src_badge}")
        for sec in ch.get("sections", []) or ch.get("thematic_sections", []):
            s_name = sec.get("title") or sec.get("section_title") or ""
            if s_name:
                s_slug = re.sub(r'[^\\w\\s-]', '', s_name).strip().replace(' ', '-')
                lines.append(f"  - [■ [جزوه {to_fa_digits(idx)}] {s_name}](#{s_slug})")
        for mod in ch.get("modules", []):
            m_name = mod.get("name", "")
            if m_name and not any(m_name in (s.get("title") or s.get("section_title") or "") for s in (ch.get("sections") or [])):
                m_slug = re.sub(r'[^\\w\\s-]', '', m_name).strip().replace(' ', '-')
                lines.append(f"  - [■ [جزوه {to_fa_digits(idx)}] {m_name}](#{m_slug})")
    lines.append("\n---\n")


def render_md_section_block(lines: List[str], sec: Dict[str, Any], sec_idx: int = 1, chapter_badge: str = ""):
    """Renders a single thematic section (table, flowchart, callouts) into Markdown."""
    if not isinstance(sec, dict):
        return
    raw_title = sec.get("title") or sec.get("section_title") or f"بخش {sec_idx}: مبحث بالینی"
    if chapter_badge and chapter_badge not in raw_title:
        sec_title = f"{chapter_badge} {raw_title}"
    else:
        sec_title = raw_title
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


def render_md_module_block(lines: List[str], mod: Dict[str, Any], mod_idx: int = 1, covered_entities: Optional[Set[str]] = None, chapter_badge: str = ""):
    """Renders a single high-yield module/illness script into Markdown."""
    if not isinstance(mod, dict):
        return
    covered_entities = covered_entities or set()
    raw_m_name = mod.get("name", f"مبحث شماره {mod_idx}")
    if chapter_badge and chapter_badge not in raw_m_name:
        m_name = f"{chapter_badge} {raw_m_name}"
    else:
        m_name = raw_m_name
    if not isinstance(mod, dict):
        return
    # m_name already computed with chapter_badge
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
    is_covered = is_module_covered_in_sections(mod, covered_entities)

    type_badge = ""
    if is_foundation:
        type_badge = " [مفاهیم و فیزیولوژی پایه]"
    elif is_approach:
        type_badge = " [رویکرد و الگوریتم تشخیصی]"

    lines.append(f"## ■ {m_name}{type_badge} ({m_name_en})\n")

    # Cross-Layer Deduplication Notice (v3.9.3)
    if is_covered:
        lines.append("> [!NOTE]\n> **⚡ ارجاع به جدول تماتیک:** جزئیات تشخیصی، کات‌آف‌ها، تظاهرات بالینی و درمان در جدول تماتیک فوق به صورت مقایسه‌ای ارائه شده است. در ادامه تنها نکات دامی آزمون، مدیریت اورژانس و هشدارهای اختصاصی درج می‌گردد.\n")

    if is_covered:
        pass  # Suppress duplicate boilerplate paragraphs; jump straight to high-yield callouts!
    # Archetype 1: Foundational Physiology & Regulatory Loops
    elif is_foundation:
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
            raw_labs = diag.get("key_labs") or []
            labs_list = raw_labs if isinstance(raw_labs, list) else ([raw_labs] if raw_labs else [])
            for lab in labs_list:
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

    # Clinical Semiology & Physical Examination (retained for covered modules)
    if is_covered and mod.get("clinical_semiology") and isinstance(mod["clinical_semiology"], list):
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



def render_md_flow_item(lines: List[str], item: Dict[str, Any], idx: int, covered_entities: Set[str]):
    """Renders a chronological flow item (topic, section_only, or module_only) into Markdown."""
    i_type = item.get("type")
    c_badge = item.get("chapter_badge", "")
    if i_type == "topic":
        if item.get("section"):
            render_md_section_block(lines, item["section"], sec_idx=idx, chapter_badge=c_badge)
        if item.get("module"):
            render_md_module_block(lines, item["module"], mod_idx=idx, covered_entities=covered_entities, chapter_badge=c_badge)
    elif i_type == "section_only":
        render_md_section_block(lines, item["section"], sec_idx=idx, chapter_badge=c_badge)
    elif i_type == "module_only":
        render_md_module_block(lines, item["module"], mod_idx=idx, covered_entities=covered_entities, chapter_badge=c_badge)

def build_markdown_summary(data: Dict[str, Any], output_path: str):
    """Generates a clean, GitHub/Obsidian-compatible Markdown study guide."""
    ensure_master_matrices(data)
    topic_title = data.get("topic_title", "خلاصه مبحث فیزیوپاتولوژی")
    topic_title_en = data.get("topic_title_en", "")
    organ_system = data.get("organ_system", "فیزیوپاتولوژی")

    lines = []
    lines.append(f"# 🩺 {topic_title} ({topic_title_en})")
    lines.append(f"> **کورس آموزشی:** {organ_system} | **مقطع:** فیزیوپاتولوژی بالینی | **نوع سند:** راهنمای مرور های‌یلد\n")
    chapters = data.get("chapters")
    is_multi = bool(data.get("is_multi_chapter") or (chapters and len(chapters) > 1))

    if is_multi and chapters:
        render_markdown_table_of_contents(lines, chapters)
        for ch_idx, ch in enumerate(chapters, 1):
            ch_title = ch.get("title") or ch.get("topic_title") or f"جزوه {ch_idx}"
            src_file = ch.get("source_file") or ch.get("file_name") or "منبع نامشخص"
            ch_organ = ch.get("organ_system") or organ_system
            lines.append(f"---\n\n# 📚 جزوه {ch_idx}: {ch_title}\n")
            lines.append(f"> **📄 فایل منبع:** `{src_file}` | **🏛️ کورس:** {ch_organ} | **🎯 سطح:** فیزیوپاتولوژی بالینی\n\n")
            ch_flow = extract_chronological_flow(ch, chapter_badge=f"[جزوه {to_fa_digits(ch_idx)}]")
            ch_covered = extract_covered_section_entities(ch)
            for it_idx, item in enumerate(ch_flow, 1):
                render_md_flow_item(lines, item, it_idx, ch_covered)
    else:
        flow = extract_chronological_flow(data, chapter_badge="")
        covered_entities = extract_covered_section_entities(data)
        for it_idx, item in enumerate(flow, 1):
            render_md_flow_item(lines, item, it_idx, covered_entities)

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

ANKI_MEDICAL_CSS = """
.card {
  font-family: 'Vazirmatn', 'Shabnam', 'IRANSans', 'Dubai', 'Segoe UI', Tahoma, sans-serif;
  font-size: 16px;
  line-height: 1.7;
  text-align: right;
  direction: rtl;
  color: #1F2937;
  background-color: #F8FAFC;
  padding: 15px;
  margin: 0 auto;
}

.anki-card-container {
  max-width: 680px;
  margin: 0 auto;
  background: #FFFFFF;
  border: 1.5px solid #007A87;
  border-radius: 12px;
  padding: 22px 26px;
  box-shadow: 0 4px 16px rgba(0, 122, 135, 0.08);
  direction: rtl;
  text-align: right;
}

.anki-header-badge {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  background: linear-gradient(135deg, #007A87 0%, #005F6A 100%);
  color: #FFFFFF;
  padding: 5px 14px;
  border-radius: 20px;
  font-size: 13px;
  font-weight: 700;
  margin-bottom: 16px;
  box-shadow: 0 2px 6px rgba(0, 122, 135, 0.25);
}

.anki-header-cloze {
  background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%);
  box-shadow: 0 2px 6px rgba(2, 132, 199, 0.25);
}

.anki-header-back {
  background: #F1F5F9;
  color: #475569;
  box-shadow: none;
  border: 1px solid #CBD5E1;
}

.anki-front-question {
  font-size: 19px;
  font-weight: 700;
  color: #0F172A;
  line-height: 1.8;
  margin: 10px 0 16px 0;
}

.anki-front-question-compact {
  font-size: 16px;
  font-weight: 600;
  color: #64748B;
  line-height: 1.6;
  margin-bottom: 12px;
}

.anki-divider {
  height: 2px;
  background: linear-gradient(to left, #007A87, #5EEAD4, transparent);
  margin: 14px 0 18px 0;
  border-radius: 2px;
}

.anki-answer-badge {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: #0E6251;
  font-size: 14px;
  font-weight: 700;
  margin-bottom: 8px;
}

.anki-back-answer {
  font-size: 20px;
  font-weight: 800;
  color: #0E6251;
  background: #E8F8F5;
  border-right: 4px solid #007A87;
  padding: 12px 18px;
  border-radius: 8px;
  line-height: 1.7;
  margin-bottom: 16px;
}

.anki-cloze-prompt {
  font-size: 19px;
  font-weight: 600;
  color: #0F172A;
  line-height: 1.8;
  margin: 10px 0 16px 0;
}

.cloze {
  font-weight: 800;
  color: #007A87;
  background: #E0F2F1;
  padding: 2px 8px;
  border-radius: 6px;
  border-bottom: 2px solid #007A87;
}

.anki-extra-box {
  background: #F0FDFA;
  border: 1px solid #99F6E4;
  border-right: 4px solid #0D9488;
  border-radius: 8px;
  padding: 12px 16px;
  margin-top: 16px;
  text-align: right;
}

.anki-extra-header {
  font-size: 13px;
  font-weight: 700;
  color: #0F766E;
  margin-bottom: 6px;
}

.anki-extra-content {
  font-size: 15px;
  color: #134E4A;
  line-height: 1.6;
}

.anki-footer-sub {
  font-size: 12px;
  color: #94A3B8;
  margin-top: 16px;
  text-align: center;
  border-top: 1px dashed #E2E8F0;
  padding-top: 10px;
}

.anki-footer {
  font-size: 12px;
  color: #94A3B8;
  margin-top: 18px;
  text-align: left;
  border-top: 1px solid #E2E8F0;
  padding-top: 8px;
  direction: ltr;
}

/* Night Mode (Dark Mode) Support */
.nightMode .card {
  background-color: #0F172A;
  color: #F1F5F9;
}
.nightMode .anki-card-container {
  background: #1E293B;
  border-color: #14B8A6;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.4);
}
.nightMode .anki-front-question {
  color: #F8FAFC;
}
.nightMode .anki-front-question-compact {
  color: #94A3B8;
}
.nightMode .anki-header-back {
  background: #334155;
  color: #CBD5E1;
  border-color: #475569;
}
.nightMode .anki-back-answer {
  background: #064E3B;
  color: #A7F3D0;
  border-right-color: #2DD4BF;
}
.nightMode .cloze {
  color: #2DD4BF;
  background: #134E4A;
  border-bottom-color: #2DD4BF;
}
.nightMode .anki-extra-box {
  background: #11292B;
  border-color: #134E4A;
  border-right-color: #2DD4BF;
}
.nightMode .anki-extra-header {
  color: #5EEAD4;
}
.nightMode .anki-extra-content {
  color: #CCFBF1;
}
.nightMode .anki-footer-sub,
.nightMode .anki-footer {
  color: #64748B;
  border-color: #334155;
}
"""


def get_or_synthesize_flashcards(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Retrieves declared flashcards or auto-synthesizes high-yield cards from modules."""
    cards = data.get("anki_flashcards") or []
    if cards:
        normalized = []
        for c in cards:
            if not isinstance(c, dict):
                continue
            card_dict = dict(c)
            # Ensure proper separation of Question and Answer for user-provided flashcards
            if card_dict.get("question") and card_dict.get("answer"):
                if not card_dict.get("text"):
                    q = str(card_dict["question"]).strip()
                    a = str(card_dict["answer"]).strip()
                    card_dict["text"] = f"{q}: {{{{c1::{a}}}}}"
                card_dict.setdefault("card_type", "basic")
            normalized.append(card_dict)
        return normalized

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
                            "card_type": "cloze",
                            "text": f"در مبحث {sec_title} ({col_0}): {{{{c1::{l}}}}}",
                            "question": f"در مبحث {sec_title}، نکته کلیدی مربوط به «{col_0}» چیست؟",
                            "answer": f"{l}",
                            "extra": f"نکته کلیدی جدول {sec_title}.",
                            "tag": f"PhysioPath::{topic_name}::TableMatrix"
                        })
            if not row_has_badge and len(row) >= 2 and col_0 and col_last and len(col_last) > 3:
                cards.append({
                    "card_type": "basic",
                    "text": f"در {sec_title}، ویژگی بالینی یا امتحانی {col_0} عبارت است از: {{{{c1::{col_last}}}}}",
                    "question": f"در مبحث {sec_title}، ویژگی بالینی یا امتحانی «{col_0}» چیست؟",
                    "answer": f"{col_last}",
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
                    "card_type": "basic",
                    "text": f"در بیمار مشکوک به {mod_name}، اقدام تشخیصی نخستین (Best Next Step) عبارت است از {{{{c1::{best_init}}}}}.",
                    "question": f"در بیمار مشکوک به {mod_name}، اقدام تشخیصی نخستین (Best Next Step) چیست؟",
                    "answer": f"{best_init}",
                    "extra": f"بررسی اولیه جهت ارزیابی بالینی سریع و غیرتهاجمی در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Diagnostics"
                })
            gold_std = diag.get("gold_standard")
            if gold_std and str(gold_std).strip() and str(gold_std).lower() not in {"null", "none", "ذکر نشده"}:
                cards.append({
                    "card_type": "basic",
                    "text": f"استاندارد طلایی و قطعی (Gold Standard) برای تشخیص {mod_name} عبارت است از {{{{c1::{gold_std}}}}}.",
                    "question": f"استاندارد طلایی و تست قطعی (Gold Standard) برای تشخیص {mod_name} چیست؟",
                    "answer": f"{gold_std}",
                    "extra": f"تست تشخیصی قطعی با بالاترین ویژگی و حساسیت در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Diagnostics"
                })
        # Synthesize from Pearls
        for pearl in ensure_list(mod.get("high_yield_pearls")):
            if pearl and str(pearl).strip():
                pearl_str = str(pearl).strip()
                cards.append({
                    "card_type": "basic",
                    "text": f"نکته های‌یلد {mod_name}: {{{{c1::{pearl_str}}}}}",
                    "question": f"نکته های‌یلد و مروارید آزمونی پیرامون {mod_name} چیست؟",
                    "answer": f"{pearl_str}",
                    "extra": f"نکته دامی آزمون پره‌انترنی/دستیاری و فیزیوپاتولوژی مبحث {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::HighYield"
                })

        # Synthesize from Contraindications & Warnings
        for warn in ensure_list(mod.get("contraindications_and_warnings")):
            if warn and str(warn).strip():
                warn_str = str(warn).strip()
                cards.append({
                    "card_type": "basic",
                    "text": f"مورد مصرف ممنوع و خطای مرگبار در {mod_name}: {{{{c1::{warn_str}}}}}",
                    "question": f"مهم‌ترین مورد مصرف ممنوع (کنترااندیکاسیون) و خطای پرخطر در {mod_name} چیست؟",
                    "answer": f"{warn_str}",
                    "extra": f"کنترااندیکاسیون قطعی و خطای کشنده پزشکی در مبحث {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Contraindications"
                })

        # Synthesize from Instructor Alerts
        for alert in ensure_list(mod.get("instructor_exam_alerts")):
            if alert and str(alert).strip():
                alert_str = str(alert).strip()
                cards.append({
                    "card_type": "basic",
                    "text": f"تأکید مستقیم استاد در آزمون {mod_name}: {{{{c1::{alert_str}}}}}",
                    "question": f"تأکید ویژه و نکته مستقیم استاد در آزمون پیرامون {mod_name} چیست؟",
                    "answer": f"{alert_str}",
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
                    "card_type": "basic",
                    "text": f"در معاینه فیزیکی و ارزیابی بالینی {mod_name}، مانور یا نشانه {{{{c1::{man}}}}} بررسی می‌شود.",
                    "question": f"در معاینه فیزیکی و ارزیابی بالینی {mod_name}، مانور یا نشانه «{man}» چگونه انجام می‌شود و چه اهمیتی دارد؟",
                    "answer": f"تکنیک: {tech} ← اهمیت: {sig}",
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
                    "card_type": "basic",
                    "text": f"اندیکاسیون جراحی و اقدام مداخله‌ای در {mod_name}: {{{{c1::{s_str}}}}}",
                    "question": f"مهم‌ترین اندیکاسیون‌های جراحی و اقدام مداخله‌ای در {mod_name} چیست؟",
                    "answer": f"{s_str}",
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
                    "card_type": "basic",
                    "text": f"عارضه بالینی وخیم و عامل پروگنوستیک در {mod_name}: {{{{c1::{c_str}}}}}",
                    "question": f"مهم‌ترین عارضه بالینی و عامل پروگنوستیک نامناسب در {mod_name} چیست؟",
                    "answer": f"{c_str}",
                    "extra": f"عوارض مهم بیماری و فاکتورهای پیش‌آگهی نامناسب در {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Complications"
                })

        # Synthesize from Foundational Feedback Cascades
        f_steps = mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback") or []
        for f_step in f_steps:
            if f_step and str(f_step).strip():
                f_str = str(f_step).strip()
                cards.append({
                    "card_type": "basic",
                    "text": f"در فیزیولوژی و فیدبک تنظیمی {mod_name}: {{{{c1::{f_str}}}}}",
                    "question": f"در فیزیولوژی و فیدبک تنظیمی {mod_name}، مکانیسم کنترلی چگونه است؟",
                    "answer": f"{f_str}",
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
                        "card_type": "basic",
                        "text": f"در رویکرد تشخیصی و الگوریتم بالینی {mod_name}: {{{{c1::{a_str}}}}}",
                        "question": f"در الگوریتم بالینی {mod_name}، گام‌های تصمیم‌گیری تشخیصی چیست؟",
                        "answer": f"{a_str}",
                        "extra": f"گام‌های تصمیم‌گیری بالینی و تریاژ تشخیصی در {mod_name}.",
                        "tag": f"PhysioPath::{mod_name}::Approach"
                    })

        dec_points = mod.get("pivot_points") or mod.get("high_yield_decision_points") or []
        for d_pt in dec_points:
            if d_pt and str(d_pt).strip():
                d_str = str(d_pt).strip()
                cards.append({
                    "card_type": "basic",
                    "text": f"نقطه انشعاب تشخیصی و تصمیم‌گیری در {mod_name}: {{{{c1::{d_str}}}}}",
                    "question": f"نقطه انشعاب تشخیصی و معیار تمایز بالینی در {mod_name} چیست؟",
                    "answer": f"{d_str}",
                    "extra": f"معیار تمایز و تصمیم‌گیری قطعی در الگوریتم بالینی {mod_name}.",
                    "tag": f"PhysioPath::{mod_name}::Approach"
                })

        # Synthesize from Emergency Management & Immediate Stabilization
        emerg_items = format_emergency_items(mod.get("emergency_management") or mod.get("immediate_stabilization"))
        for em in emerg_items:
            if em and str(em).strip():
                e_str = str(em).strip()
                cards.append({
                    "card_type": "basic",
                    "text": f"اقدام اورژانس و احیای فوری در {mod_name}: {{{{c1::{e_str}}}}}",
                    "question": f"اقدام اورژانس و احیای فوری نجات‌بخش در مواجهه با {mod_name} چیست؟",
                    "answer": f"{e_str}",
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
                        "card_type": "basic",
                        "text": f"معیار بالینی و سیستم مرحله‌بندی {mod_name}: {{{{c1::{cr_str}}}}}",
                        "question": f"معیار بالینی و سیستم مرحله‌بندی / نمره‌دهی در {mod_name} چیست؟",
                        "answer": f"{cr_str}",
                        "extra": f"سیستم طبقه‌بندی، نمره‌دهی بالینی و پیش‌آگهی در {mod_name}.",
                        "tag": f"PhysioPath::{mod_name}::Staging"
                    })
        elif isinstance(crit_obj, dict) and crit_obj.get("rows"):
            for row in crit_obj.get("rows", []):
                r_str = " - ".join(str(c) for c in row if c)
                if r_str:
                    cards.append({
                        "card_type": "basic",
                        "text": f"معیار و طبقه‌بندی بالینی {mod_name}: {{{{c1::{r_str}}}}}",
                        "question": f"معیار و طبقه‌بندی بالینی در {mod_name} چیست؟",
                        "answer": f"{r_str}",
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
                    "card_type": "basic",
                    "text": f"ملاحظه در بارداری و گروه‌های خاص در {mod_name}: {{{{c1::{sp_str}}}}}",
                    "question": f"مهم‌ترین ملاحظه درمانی در بارداری یا گروه‌های خاص برای بیماری {mod_name} چیست؟",
                    "answer": f"{sp_str}",
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
                    "card_type": "basic",
                    "text": f"اهمیت بالینی مفهوم پایه {mod_name}: {{{{c1::{cr_str}}}}}",
                    "question": f"اهمیت بالینی مفهوم پایه {mod_name} چیست؟",
                    "answer": f"{cr_str}",
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
                        "card_type": "basic",
                        "text": f"در تحلیل اسلاید بالینی {f_title} ({mod_name})، یافته کلیدی عبارت است از {{{{c1::{f_cap}}}}}.",
                        "question": f"در تحلیل اسلاید بالینی «{f_title}» ({mod_name})، یافته کلیدی چیست؟",
                        "answer": f"{f_cap}",
                        "extra": f"نکته تکمیلی / تله امتحانی: {f_pearl}" if f_pearl else f"تحلیل اطلس تصویری و اسلایدهای بالینی مبحث {mod_name}.",
                        "tag": f"PhysioPath::{mod_name}::VisualAtlas"
                    })
                elif f_pearl:
                    cards.append({
                        "card_type": "basic",
                        "text": f"نکته کلیدی استاد در اسلاید بالینی {f_title} ({mod_name}): {{{{c1::{f_pearl}}}}}",
                        "question": f"نکته کلیدی استاد در اسلاید بالینی «{f_title}» ({mod_name}) چیست؟",
                        "answer": f"{f_pearl}",
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
                "card_type": "basic",
                "text": f"در معاینه فیزیکی و ارزیابی بالینی ({topic_name})، مانور یا نشانه {{{{c1::{man}}}}} ارزیابی می‌شود.",
                "question": f"در معاینه فیزیکی و ارزیابی بالینی ({topic_name})، مانور یا نشانه «{man}» چگونه انجام می‌شود و چه اهمیتی دارد؟",
                "answer": f"تکنیک: {tech} ← اهمیت: {sig}",
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
                    "card_type": "basic",
                    "text": f"در اطلس تصویری و اسلایدهای تدریس استاد ({f_title})، یافته بالینی/رادیولوژیک کلیدی عبارت است از {{{{c1::{f_cap}}}}}.",
                    "question": f"در اطلس تصویری تدریس استاد ({f_title})، یافته بالینی/رادیولوژیک کلیدی چیست؟",
                    "answer": f"{f_cap}",
                    "extra": f"نکته تکمیلی / تله امتحانی: {f_pearl}" if f_pearl else f"تحلیل اطلس تصویری دوره {topic_name}.",
                    "tag": f"PhysioPath::{topic_name}::VisualAtlas"
                })
            elif f_pearl:
                cards.append({
                    "card_type": "basic",
                    "text": f"نکته طلایی استاد در اسلاید بالینی {f_title}: {{{{c1::{f_pearl}}}}}",
                    "question": f"نکته طلایی استاد در اسلاید بالینی «{f_title}» چیست؟",
                    "answer": f"{f_pearl}",
                    "extra": f"نکته تشخیصی اطلس تصویری دوره {topic_name}.",
                    "tag": f"PhysioPath::{topic_name}::VisualAtlas"
                })
    return cards


def build_anki_tsv(data: Dict[str, Any], output_path: str):
    """
    Exports atomic cloze-deletion cards formatted as TSV for direct import into Anki.
    Includes Anki 2.1+ file directives (#separator:tab, #html:true, #notetype:Cloze).
    Also exports companion _basic.tsv with (#notetype:Basic) for dedicated Q&A import.
    """
    cards = get_or_synthesize_flashcards(data)

    cloze_lines = [
        "#separator:tab",
        "#html:true",
        "#notetype:Cloze",
        "#tags column:3"
    ]
    basic_lines = [
        "#separator:tab",
        "#html:true",
        "#notetype:Basic",
        "#tags column:4"
    ]

    has_basic = False
    for card in cards:
        text = str(card.get("text") or card.get("question") or card.get("front") or "").strip()
        if not text:
            continue
        if "{{c" not in text:
            # Atomic split on colon if available, otherwise wrap
            if ":" in text or "：" in text:
                parts = text.split(":", 1) if ":" in text else text.split("：", 1)
                text = f"{parts[0].strip()}: {{{{c1::{parts[1].strip()}}}}}"
            else:
                text = f"{{{{c1::{text}}}}}"

        text_clean = sanitize_anki_html(text).replace("\t", " ").replace("\n", "<br>")
        extra_val = card.get("extra") or card.get("answer") or card.get("back") or ""
        extra = sanitize_anki_html(str(extra_val)).replace("\t", " ").replace("\n", "<br>")
        tag = str(card.get("tag", "PhysioPath::General")).replace(" ", "_").replace("\t", "")
        cloze_lines.append(f"{text_clean}\t{extra}\t{tag}")

        # Basic card line
        q = str(card.get("question") or card.get("front") or "").strip()
        a = str(card.get("answer") or card.get("back") or "").strip()
        if q and a:
            has_basic = True
            q_clean = sanitize_anki_html(q).replace("\t", " ").replace("\n", "<br>")
            a_clean = sanitize_anki_html(a).replace("\t", " ").replace("\n", "<br>")
            e_clean = sanitize_anki_html(str(card.get("extra") or "")).replace("\t", " ").replace("\n", "<br>")
            basic_lines.append(f"{q_clean}\t{a_clean}\t{e_clean}\t{tag}")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(cloze_lines))
    print(f"[OK] Anki Cloze TSV Deck exported successfully: {output_path} ({len(cards)} cards)")

    if has_basic and output_path.endswith(".tsv"):
        basic_path = output_path[:-4] + "_basic.tsv"
        with open(basic_path, "w", encoding="utf-8") as f:
            f.write("\n".join(basic_lines))
        print(f"[OK] Anki Basic Q&A TSV Deck exported successfully: {basic_path}")


def build_anki_apkg(data: Dict[str, Any], output_path: str):
    """
    Packages flashcards into an interactive .apkg file if genanki is installed.
    Supports Smart Dual Note Type Architecture:
    1. PhysioPath Basic Model (Front/Back) with Question/Answer separation & Medical styling.
    2. PhysioPath Cloze Model (Cloze) with atomic deletions & Emerald highlighting.
    Enforces deterministic deck IDs and card GUIDs anchored exclusively to prompt text.
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

    # 1. PhysioPath Basic Model (Front / Back / Extra / Source)
    basic_model = genanki.Model(
        1607392320,
        'PhysioPath Basic Model',
        fields=[
            {'name': 'Front'},
            {'name': 'Back'},
            {'name': 'Extra'},
            {'name': 'Source'},
        ],
        templates=[
            {
                'name': 'PhysioPath Clinical Q&A',
                'qfmt': '''<div class="anki-card-container" dir="rtl">
  <div class="anki-header-badge">
    <span>🩺</span> <span>سوال بالینی و سناریوی تشخیصی</span>
  </div>
  <div class="anki-front-question">{{Front}}</div>
  <div class="anki-footer-sub">💡 برای مشاهده پاسخ تشخیصی و تصمیم بالینی کلیک کنید</div>
</div>''',
                'afmt': '''<div class="anki-card-container" dir="rtl">
  <div class="anki-header-badge anki-header-back">
    <span>🩺</span> <span>سوال بالینی و سناریوی تشخیصی</span>
  </div>
  <div class="anki-front-question-compact">{{Front}}</div>
  <div class="anki-divider"></div>
  <div class="anki-answer-badge">
    <span>✅</span> <span>پاسخ تشخیصی و اقدام درمانی:</span>
  </div>
  <div class="anki-back-answer">{{Back}}</div>
  {{#Extra}}
  <div class="anki-extra-box">
    <div class="anki-extra-header">💡 مکانیزم پاتوفیزیولوژی / تله امتحانی:</div>
    <div class="anki-extra-content">{{Extra}}</div>
  </div>
  {{/Extra}}
  <div class="anki-footer">📚 دک هوشمند درسنامه پزشکی | Medical Summary Expert</div>
</div>''',
            },
        ],
        css=ANKI_MEDICAL_CSS,
    )

    # 2. PhysioPath Cloze Model (Text / Extra / Source)
    cloze_model = genanki.Model(
        1607392319,
        'PhysioPath Cloze Model',
        fields=[
            {'name': 'Text'},
            {'name': 'Extra'},
            {'name': 'Source'},
        ],
        templates=[
            {
                'name': 'PhysioPath Cloze Card',
                'qfmt': '''<div class="anki-card-container" dir="rtl">
  <div class="anki-header-badge anki-header-cloze">
    <span>⚡</span> <span>جای‌خالی های‌یلد آزمونی</span>
  </div>
  <div class="anki-cloze-prompt">{{cloze:Text}}</div>
  <div class="anki-footer-sub">🎯 کلیدواژه یا کات‌آف بالینی پنهان‌شده را حدس بزنید</div>
</div>''',
                'afmt': '''<div class="anki-card-container" dir="rtl">
  <div class="anki-header-badge anki-header-cloze">
    <span>⚡</span> <span>پاسخ جای‌خالی و نکته آزمونی</span>
  </div>
  <div class="anki-cloze-prompt">{{cloze:Text}}</div>
  {{#Extra}}
  <div class="anki-extra-box">
    <div class="anki-extra-header">💡 مکانیزم پاتوفیزیولوژی / تله امتحانی:</div>
    <div class="anki-extra-content">{{Extra}}</div>
  </div>
  {{/Extra}}
  <div class="anki-footer">📚 دک هوشمند درسنامه پزشکی | Medical Summary Expert</div>
</div>''',
            },
        ],
        css=ANKI_MEDICAL_CSS,
        model_type=genanki.Model.CLOZE
    )

    deck_name = f"PhysioPath::{data.get('topic_title', 'Medical_Summary')}"
    # Deterministic deck ID based on deck name hash (prevents random deck IDs)
    deck_id = int(hashlib.md5(deck_name.encode('utf-8')).hexdigest()[:8], 16)
    deck = genanki.Deck(deck_id, deck_name)

    for c in cards:
        card_type = str(c.get("card_type") or "").strip().lower()
        has_qa = bool(c.get("question") and c.get("answer")) or bool(c.get("front") and c.get("back"))
        has_cloze = "{{c" in str(c.get("text") or "")

        tags = [str(c.get("tag", "PhysioPath")).replace(" ", "_")]

        # If card explicitly declared basic or has Q&A without cloze in text, use basic model
        if card_type == "basic" and has_qa:
            q_text = str(c.get("question") or c.get("front") or "").strip()
            a_text = str(c.get("answer") or c.get("back") or "").strip()
            e_text = str(c.get("extra") or "").strip()

            q_clean = sanitize_anki_html(q_text)
            a_clean = sanitize_anki_html(a_text)
            e_clean = sanitize_anki_html(e_text)

            norm_prompt = " ".join(q_clean.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")).replace('\u200c', ' ').split())
            card_guid = genanki.guid_for("basic_" + norm_prompt)
            note = genanki.Note(
                model=basic_model,
                fields=[q_clean, a_clean, e_clean, "PhysioPath"],
                tags=tags,
                guid=card_guid
            )
            deck.add_note(note)
        else:
            # Cloze model note
            text = str(c.get("text") or "").strip()
            if not text:
                if has_qa:
                    q_text = str(c.get("question") or c.get("front") or "").strip()
                    a_text = str(c.get("answer") or c.get("back") or "").strip()
                    text = f"{q_text}: {{{{c1::{a_text}}}}}"
                else:
                    continue

            if "{{c" not in text:
                if ":" in text or "：" in text:
                    parts = text.split(":", 1) if ":" in text else text.split("：", 1)
                    text = f"{parts[0].strip()}: {{{{c1::{parts[1].strip()}}}}}"
                else:
                    text = f"{{{{c1::{text}}}}}"

            text_clean = sanitize_anki_html(text)
            extra_val = c.get("extra") or c.get("answer") or c.get("back") or ""
            extra_clean = sanitize_anki_html(str(extra_val))

            norm_prompt = " ".join(text_clean.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")).replace('\u200c', ' ').split())
            card_guid = genanki.guid_for("cloze_" + norm_prompt)
            note = genanki.Note(
                model=cloze_model,
                fields=[text_clean, extra_clean, "PhysioPath"],
                tags=tags,
                guid=card_guid
            )
            deck.add_note(note)

    genanki.Package(deck).write_to_file(output_path)
    print(f"[OK] Anki Native .apkg Package created successfully: {output_path} (Deterministic Deck ID: {deck_id})")


# HTML & PDF Generation Engine (Playwright + Persian BiDi Stamping)
# ---------------------------------------------------------------------------


def render_html_table_of_contents(html_parts: List[str], chapters: List[Dict[str, Any]], title: str = "📑 فهرست فصول و سرفصل‌های درسنامه جامع"):
    """Renders interactive HTML Table of Contents with jump anchors and source metadata."""
    html_parts.append(f"""
    <div class="toc-container">
        <div class="toc-header">
            <div class="toc-main-title">{title}</div>
            <div class="toc-sub-title">تفکیک جامع سرفصل‌ها و مباحث به تفکیک جزوات منبع</div>
        </div>
        <div class="toc-grid">
    """)
    for idx, ch in enumerate(chapters, 1):
        ch_title = ch.get("title") or ch.get("topic_title") or f"جزوه {idx}"
        src = ch.get("source_file") or ch.get("file_name") or "منبع نامشخص"
        organ = ch.get("organ_system") or "فیزیوپاتولوژی"
        topics = []
        for sec in (ch.get("sections") or ch.get("thematic_sections") or []):
            t_name = sec.get("title") or sec.get("section_title") or ""
            if t_name and t_name not in topics:
                topics.append(t_name)
        for mod in (ch.get("modules") or []):
            m_name = mod.get("name", "")
            if m_name and not any(m_name in t for t in topics):
                topics.append(m_name)
        if not topics:
            topics = ["مباحث جامع بالینی"]

        topic_items_html = "".join([f"<li>• [جزوه {to_fa_digits(idx)}] {format_inline_markdown(t)}</li>" for t in topics[:8]])

        html_parts.append(f"""
            <div class="toc-card">
                <div class="toc-card-header">
                    <a href="#handout-{idx}" class="toc-link">📚 {ch_title}</a>
                    <div class="toc-organ-badge">{organ}</div>
                </div>
                <div class="toc-meta">
                    <span>📄 فایل منبع: <code dir="ltr">{src}</code></span>
                    <span class="badge-gold">{len(topics)} مبحث تدریس‌شده</span>
                </div>
                <ul class="toc-topic-list">
                    {topic_items_html}
                </ul>
            </div>
        """)
    html_parts.append("""
        </div>
    </div>
    """)


def render_html_section_block(html_parts: List[str], sec: Dict[str, Any], sec_idx: int = 1, chapter_badge: str = "", out_dir: str = ""):
    """Renders a single thematic section (table, flowchart, callouts) into HTML."""
    if not isinstance(sec, dict):
        return
    raw_title = sec.get("title") or sec.get("section_title") or f"بخش {sec_idx}: مبحث بالینی"
    if chapter_badge and chapter_badge not in raw_title:
        sec_title = f"{chapter_badge} {raw_title}"
    else:
        sec_title = raw_title
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
        target_dir = out_dir if out_dir else os.path.join(os.getcwd(), "خلاصه")
        diag_dir = os.path.join(target_dir, "diagrams")
        os.makedirs(diag_dir, exist_ok=True)
        m_hash = hashlib.md5(sec_flowchart.strip().encode("utf-8")).hexdigest()[:10]
        png_path = os.path.join(diag_dir, f"diag_sec_{m_hash}.png")
        if not os.path.exists(png_path) or os.path.getsize(png_path) < 1000:
            render_mermaid_diagram_to_png(sec_flowchart, png_path, title=f"الگوریتم {sec_title}")

        enriched_m = enrich_mermaid_flowchart(sec_flowchart, title=sec_title)
        if os.path.exists(png_path) and os.path.getsize(png_path) > 1000:
            rel_png = os.path.relpath(png_path, target_dir) if target_dir else png_path
            rel_png = rel_png.replace("\\", "/")
            html_parts.append(f'<div class="diagram-card"><div class="diagram-header">🧭 الگوریتم بالینی: {sec_title}</div><img class="diagram-img" src="{rel_png}" alt="الگوریتم بالینی {sec_title}"><div style="display:none;" class="mermaid">{enriched_m}</div></div>')
        else:
            html_parts.append(f'<div class="diagram-card"><div class="diagram-header">🧭 الگوریتم بالینی: {sec_title}</div><div class="mermaid">{enriched_m}</div></div>')

    sec_callouts = sec.get("callouts") or []
    if sec_callouts:
        html_parts.append('<div class="callout callout-emergency"><div class="callout-header">🚨 هشدار اورژانسی و بالینی</div><div class="callout-body">')
        for c_item in (sec_callouts if isinstance(sec_callouts, list) else [sec_callouts]):
            html_parts.append(f'<div>• {format_inline_markdown(str(c_item))}</div>')
        html_parts.append('</div></div>')


def render_html_module_block(html_parts: List[str], mod: Dict[str, Any], mod_idx: int = 1, covered_entities: Optional[Set[str]] = None, chapter_badge: str = ""):
    """Renders a single high-yield module/illness script into HTML."""
    if not isinstance(mod, dict):
        return
    covered_entities = covered_entities or set()
    raw_m_name = mod.get("name", f"مبحث شماره {mod_idx}")
    if chapter_badge and chapter_badge not in raw_m_name:
        m_name = f"{chapter_badge} {raw_m_name}"
    else:
        m_name = raw_m_name
    if not isinstance(mod, dict):
        return
    # m_name already computed with chapter_badge
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
    is_covered = is_module_covered_in_sections(mod, covered_entities)

    type_badge = ""
    if is_foundation:
        type_badge = " [مفاهیم و فیزیولوژی پایه]"
    elif is_approach:
        type_badge = " [رویکرد و الگوریتم تشخیصی]"
    en_tag = f" ({m_name_en})" if m_name_en else ""
    html_parts.append(f'<div class="module-header">■ {m_name}{type_badge}{en_tag}</div>')

    # Cross-Layer Deduplication Notice (v3.9.3)
    if is_covered:
        html_parts.append('<div class="callout callout-semiology"><div class="callout-header">⚡ ارجاع به جدول تماتیک</div><div class="callout-body">جزئیات تشخیصی، کات‌آف‌ها، تظاهرات بالینی و درمان در جدول تماتیک فوق به صورت مقایسه‌ای ارائه شده است. در ادامه تنها نکات دامی آزمون، مدیریت اورژانس و هشدارهای اختصاصی درج می‌گردد.</div></div>')

    if is_covered:
        pass  # Suppress duplicate boilerplate paragraphs; jump straight to high-yield callouts!
    # Archetype 1: Foundational Physiology & Regulatory Loops
    elif is_foundation:
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
            raw_labs = diag.get("key_labs") or []
            labs_list = raw_labs if isinstance(raw_labs, list) else ([raw_labs] if raw_labs else [])
            for lab in labs_list:
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

    # Clinical Semiology & Physical Examination (retained for covered modules)
    if is_covered and mod.get("clinical_semiology") and isinstance(mod["clinical_semiology"], list):
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



def render_html_flow_item(html_parts: List[str], item: Dict[str, Any], idx: int, covered_entities: Set[str]):
    """Renders a chronological flow item (topic, section_only, or module_only) into HTML."""
    i_type = item.get("type")
    c_badge = item.get("chapter_badge", "")
    if i_type == "topic":
        if item.get("section"):
            render_html_section_block(html_parts, item["section"], sec_idx=idx, chapter_badge=c_badge)
        if item.get("module"):
            render_html_module_block(html_parts, item["module"], mod_idx=idx, covered_entities=covered_entities, chapter_badge=c_badge)
    elif i_type == "section_only":
        render_html_section_block(html_parts, item["section"], sec_idx=idx, chapter_badge=c_badge)
    elif i_type == "module_only":
        render_html_module_block(html_parts, item["module"], mod_idx=idx, covered_entities=covered_entities, chapter_badge=c_badge)

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
    font-size: 8.5pt;
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
    font-size: 8.8pt;
    direction: rtl;
}
td {
    padding: 6px 9px;
    border: 1px solid #CBD5E1;
    vertical-align: top;
    text-align: right;
    line-height: 1.38;
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
.toc-container {
    background: #F8FAFC;
    border: 2px solid #007A87;
    border-radius: 8px;
    padding: 16px;
    margin: 18px 0 24px 0;
    page-break-after: always;
}
.toc-main-title {
    color: #007A87;
    font-size: 14pt;
    font-weight: bold;
    margin-bottom: 4px;
}
.toc-sub-title {
    color: #64748B;
    font-size: 9pt;
    margin-bottom: 12px;
}
.toc-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 12px;
}
.toc-card {
    background: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-right: 4px solid #007A87;
    border-radius: 6px;
    padding: 10px 12px;
}
.toc-card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.toc-link {
    color: #007A87;
    font-weight: bold;
    text-decoration: none;
    font-size: 10.5pt;
}
.toc-link:hover {
    text-decoration: underline;
}
.toc-organ-badge {
    display: inline-block;
    background: #E0F2FE;
    color: #0369A1;
    font-size: 8pt;
    padding: 2px 6px;
    border-radius: 4px;
    font-weight: bold;
}
.toc-meta {
    font-size: 8.5pt;
    color: #64748B;
    margin: 6px 0;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.toc-topic-list {
    margin: 6px 0 0 0;
    padding-right: 14px;
    font-size: 8.5pt;
    color: #334155;
    list-style-type: none;
}
.toc-topic-list li {
    margin: 2px 0;
}
.handout-banner {
    background: #F0FDFA;
    border: 2px solid #007A87;
    border-right: 6px solid #007A87;
    border-radius: 6px;
    padding: 12px 16px;
    margin: 24px 0 14px 0;
    page-break-before: always;
}
.handout-banner-title {
    font-size: 14pt;
    font-weight: bold;
    color: #007A87;
    margin-bottom: 4px;
}
.handout-banner-meta {
    font-size: 9pt;
    color: #475569;
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
            mermaid.initialize({
            startOnLoad: true,
            theme: 'base',
            themeVariables: {
                fontFamily: 'Dubai, Vazirmatn, Segoe UI, Tahoma, sans-serif',
                primaryColor: '#F0FDFA',
                primaryBorderColor: '#007A87',
                primaryTextColor: '#0F172A',
                lineColor: '#0284C7',
                secondaryColor: '#FEF3C7',
                tertiaryColor: '#FFE4E6',
                edgeLabelBackground: '#FFFFFF',
                fontSize: '15px'
            },
            securityLevel: 'loose'
        });
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

    chapters = data.get("chapters")
    is_multi = bool(data.get("is_multi_chapter") or (chapters and len(chapters) > 1))

    if is_multi and chapters:
        render_html_table_of_contents(html_parts, chapters)
        for ch_idx, ch in enumerate(chapters, 1):
            ch_title = ch.get("title") or ch.get("topic_title") or f"جزوه {ch_idx}"
            src_file = ch.get("source_file") or ch.get("file_name") or "منبع نامشخص"
            ch_organ = ch.get("organ_system") or organ_system
            html_parts.append(f"""
    <div class="handout-banner" id="handout-{ch_idx}">
        <div class="handout-banner-title">📚 جزوه {ch_idx}: {ch_title}</div>
        <div class="handout-banner-meta">📄 فایل منبع: <code dir="ltr">{src_file}</code> | 🏛️ کورس: {ch_organ} | 🎯 سطح: فیزیوپاتولوژی بالینی</div>
    </div>
    """)
            ch_flow = extract_chronological_flow(ch, chapter_badge=f"[جزوه {to_fa_digits(ch_idx)}]")
            ch_covered = extract_covered_section_entities(ch)
            for it_idx, item in enumerate(ch_flow, 1):
                render_html_flow_item(html_parts, item, it_idx, ch_covered)
    else:
        flow = extract_chronological_flow(data, chapter_badge="")
        covered_entities = extract_covered_section_entities(data)
        for it_idx, item in enumerate(flow, 1):
            render_html_flow_item(html_parts, item, it_idx, covered_entities)

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
