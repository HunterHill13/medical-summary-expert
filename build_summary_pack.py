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
import argparse
import hashlib
from typing import List, Dict, Any, Optional

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

# Color Palette Specifications
COLOR_PRIMARY_BURGUNDY = (0x78, 0x28, 0x1F)  # #78281F - Main Headings & Accent
COLOR_SECONDARY_TEAL   = (0x0E, 0x62, 0x51)  # #0E6251 - Subheadings (H2)
COLOR_TEXT_CHARCOAL    = (0x26, 0x26, 0x26)  # #262626 - Main Body Text
COLOR_BLUE_BORDER      = "85C1E9"            # Soft Sky Blue Border for Pathophysiology
COLOR_BLUE_BG          = "EBF5FB"            # Soft Sky Blue Header Fill
COLOR_AMBER_BORDER     = "D35400"            # Amber/Orange Border for High-Yield Pearls
COLOR_AMBER_BG         = "FEF9E7"            # Soft Yellow/Cream Fill for Pearls
COLOR_TABLE_HEADER_BG  = "78281F"            # Burgundy Table Header
COLOR_TABLE_HEADER_TXT = "FFFFFF"            # White Table Header Text
COLOR_TABLE_ALT_ROW    = "FDFEFE"            # Alternating Row Light Fill
COLOR_MUTED_GRAY       = (0x7F, 0x8C, 0x8D)  # Gray for timestamps/meta

FONT_DUBAI = "Dubai"

# ---------------------------------------------------------------------------
# BiDi & OpenXML Typography Helpers for Word (.docx)
# ---------------------------------------------------------------------------

def add_r(p, text: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
          bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL,
          is_rtl: bool = True):
    """Adds a run with explicit Complex Script (w:szCs, w:bCs, w:rtl) settings for Word."""
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


def add_bidi_text(p, text: str, font_name: str = FONT_DUBAI, size_pt: float = 11.0,
                  bold: bool = False, italic: bool = False, color_rgb: tuple = COLOR_TEXT_CHARCOAL):
    """
    Renders mixed Persian/English scientific text with strict BiDi isolation.
    Preserves arrows (→, ⇌), chemical symbols (Ca²⁺, HCO₃⁻), inequalities (≥, ≤, ±)
    and Latin abbreviations without reversing English parentheses.
    """
    if not text:
        return

    # Handle explicit line breaks
    lines = text.split('\n')
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break()
        _render_bidi_line(p, line, font_name=font_name, size_pt=size_pt,
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

    # Force table RTL
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(r'<w:bidiVisual {} w:val="1"/>'.format(nsdecls('w'))))

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

    # Force table RTL
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(r'<w:bidiVisual {} w:val="1"/>'.format(nsdecls('w'))))

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

def add_clinical_vignettes(doc, vignettes: List[Dict[str, Any]]):
    """Renders active recall case scenarios with choices, correct answers and pathophysiological rationale."""
    if not vignettes:
        return
    for idx, v in enumerate(vignettes, 1):
        table = doc.add_table(rows=2, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = table._tbl.tblPr
        tblPr.append(parse_xml(r'<w:bidiVisual {} w:val="1"/>'.format(nsdecls('w'))))

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
            etio_items = mod.get("definition_and_etiology") or ["-"]
            m_etio = " ".join(str(x) for x in etio_items if x)[:60] or "-"
            mech_items = mod.get("pathophysiology_cascade") or ["-"]
            m_mech = str(mech_items[0] if mech_items else "-")[:60]
            pres_items = mod.get("clinical_presentation") or []
            m_sign = str(pres_items[0].get("sign", "-") if pres_items and isinstance(pres_items[0], dict) else "-")[:50]
            diag_dict = mod.get("diagnostics") or {}
            m_dx = str(diag_dict.get("best_initial_test") or "-")[:50]
            rx_items = mod.get("pharmacology_and_management") or []
            m_rx = str(rx_items[0].get("drug_class", "-") if rx_items and isinstance(rx_items[0], dict) else "-")[:50]
            crash_rows.append([m_name, m_etio, m_mech, m_sign, m_dx, m_rx])
        if crash_rows:
            data["last_minute_crash_sheet"] = {
                "headers": ["بیماری / مبحث", "اتیولوژی اصلی", "مکانیسم پاتوفیزیولوژی کلیدی", "علامت شاخص / تریاد", "اقدام تشخیصی نخستین", "درمان انتخابی"],
                "rows": crash_rows
            }


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

    # Iterate over modules/diseases in the data
    modules = data.get("modules", [])
    for mod_idx, mod in enumerate(modules, 1):
        mod_name = mod.get("name", f"مبحث شماره {mod_idx}")
        mod_name_en = mod.get("name_en", "")

        # Heading 1 (Module Title)
        p_h1 = doc.add_paragraph()
        set_p_rtl(p_h1, space_before_pt=14, space_after_pt=6)
        add_r(p_h1, f"■ {mod_name}", font_name=FONT_DUBAI, size_pt=13.5, bold=True,
              color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
        if mod_name_en:
            add_r(p_h1, f"  ({mod_name_en})", font_name=FONT_DUBAI, size_pt=12.0, bold=True,
                  color_rgb=COLOR_SECONDARY_TEAL, is_rtl=False)

        # 1. Definition & Etiology
        if mod.get("definition_and_etiology") and isinstance(mod["definition_and_etiology"], list):
            p_sub = doc.add_paragraph()
            set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
            add_r(p_sub, "۱. تعریف و اتیولوژی (Definition & Etiology)",
                  font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

            for b in mod["definition_and_etiology"]:
                p_b = doc.add_paragraph()
                set_p_rtl(p_b, space_before_pt=1, space_after_pt=3)
                add_r(p_b, "• ", font_name=FONT_DUBAI, size_pt=10.5, bold=True,
                      color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)
                add_bidi_text(p_b, str(b), font_name=FONT_DUBAI, size_pt=10.5, bold=False,
                              color_rgb=COLOR_TEXT_CHARCOAL)

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

        # 5. Pharmacology & Mechanism-Based Treatment
        if mod.get("pharmacology_and_management") and isinstance(mod["pharmacology_and_management"], list):
            p_sub = doc.add_paragraph()
            set_p_rtl(p_sub, space_before_pt=6, space_after_pt=3)
            add_r(p_sub, "۵. اصول فارماکولوژی و درمان مبتنی بر مکانیسم (Mechanism-Based Rx)",
                  font_name=FONT_DUBAI, size_pt=11.0, bold=True, color_rgb=COLOR_SECONDARY_TEAL, is_rtl=True)

            for rx in mod["pharmacology_and_management"]:
                if not isinstance(rx, dict):
                    continue
                p_rx = doc.add_paragraph()
                set_p_rtl(p_rx, space_before_pt=1, space_after_pt=3)
                drug_class = rx.get("drug_class", "")
                mechanism = rx.get("mechanism_of_action") or rx.get("mechanism") or ""
                pearl = rx.get("pearl") or rx.get("exam_pearl") or ""
                add_r(p_rx, f"💊 دسته دارویی {drug_class}: ", font_name=FONT_DUBAI, size_pt=10.5,
                      bold=True, color_rgb=COLOR_PRIMARY_BURGUNDY, is_rtl=True)
                text_desc = f"مکانیسم عمل: {mechanism}"
                if pearl:
                    text_desc += f" | نکته کلیدی: {pearl}"
                add_bidi_text(p_rx, text_desc, font_name=FONT_DUBAI, size_pt=10.5)

        # 6. Differential Diagnosis Table
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

        # Clinical Scoring Criteria (if any)
        if mod.get("scoring_criteria") and isinstance(mod["scoring_criteria"], dict):
            sc = mod["scoring_criteria"]
            sc_headers = sc.get("headers", ["معیار بالینی", "آستانه / امتیاز", "تفسیر بالینی"])
            sc_rows = sc.get("rows", [])
            add_comparison_table(doc, title=sc.get("title", f"سیستم نمره‌دهی و طبقه‌بندی {mod_name}"), headers=sc_headers, rows=sc_rows)

        # Embedded Figures (if any)
        for fig in (mod.get("figures") or []):
            if isinstance(fig, dict):
                add_figure_block(doc, fig.get("path", ""), fig.get("caption", ""))

        # Clinical Vignettes (Active Recall MCQs)
        if mod.get("clinical_vignettes") and isinstance(mod["clinical_vignettes"], list):
            add_clinical_vignettes(doc, mod["clinical_vignettes"])

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

def build_markdown_summary(data: Dict[str, Any], output_path: str):
    """Generates a clean, GitHub/Obsidian-compatible Markdown study guide."""
    ensure_master_matrices(data)
    topic_title = data.get("topic_title", "خلاصه مبحث فیزیوپاتولوژی")
    topic_title_en = data.get("topic_title_en", "")
    organ_system = data.get("organ_system", "فیزیوپاتولوژی")

    lines = []
    lines.append(f"# 🩺 {topic_title} ({topic_title_en})")
    lines.append(f"> **کورس آموزشی:** {organ_system} | **مقطع:** فیزیوپاتولوژی بالینی | **نوع سند:** راهنمای مرور های‌یلد\n")

    for mod in (data.get("modules") or []):
        if not isinstance(mod, dict):
            continue
        m_name = mod.get("name", "")
        m_name_en = mod.get("name_en", "")
        lines.append(f"## ■ {m_name} ({m_name_en})\n")

        # Definition
        if mod.get("definition_and_etiology") and isinstance(mod["definition_and_etiology"], list):
            lines.append("### ۱. تعریف و اتیولوژی (Definition & Etiology)")
            for item in mod["definition_and_etiology"]:
                lines.append(f"- {item}")
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

        # Pharmacology
        if mod.get("pharmacology_and_management") and isinstance(mod["pharmacology_and_management"], list):
            lines.append("### ۵. اصول فارماکولوژی مبتنی بر مکانیسم (Pharmacology & Rx)")
            for rx in mod["pharmacology_and_management"]:
                if not isinstance(rx, dict):
                    continue
                d_class = rx.get("drug_class", "")
                mechanism = rx.get("mechanism_of_action") or rx.get("mechanism") or ""
                pearl = rx.get("pearl") or rx.get("exam_pearl") or ""
                desc = f"مکانیسم عمل: {mechanism}"
                if pearl:
                    desc += f" (نکته: {pearl})"
                lines.append(f"- **{d_class}:** {desc}")
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
            lines.append("> [!WARNING] **نکات دامی آزمونی و مرواریدهای کلیدی (High-Yield Pearls):**")
            for pearl in mod["high_yield_pearls"]:
                lines.append(f"> - ⚡ {pearl}")
            lines.append("")

        # Instructor Alerts
        if mod.get("instructor_exam_alerts") and isinstance(mod["instructor_exam_alerts"], list):
            lines.append("> [!IMPORTANT] **⭐ اولویت اول مرور: تأکیدات مستقیم استاد در کلاس:**")
            for alert in mod["instructor_exam_alerts"]:
                lines.append(f"> - ⭐ {alert}")
            lines.append("")

        # Contraindications & Lethal Withholds
        if mod.get("contraindications_and_warnings") and isinstance(mod["contraindications_and_warnings"], list):
            lines.append("> [!CAUTION] **⛔ موارد مصرف ممنوع و خطاهای مرگبار (Absolute Contraindications):**")
            for warning in mod["contraindications_and_warnings"]:
                lines.append(f"> - ⛔ {warning}")
            lines.append("")

        # Scoring Criteria
        if mod.get("scoring_criteria") and isinstance(mod["scoring_criteria"], dict):
            sc = mod["scoring_criteria"]
            sc_title = sc.get("title", f"سیستم نمره‌دهی و طبقه‌بندی {m_name}")
            lines.append(f"### 📋 {sc_title}")
            sc_hdrs = sc.get("headers", ["معیار", "آستانه", "تفسیر"])
            lines.append(format_md_table_row(sc_hdrs))
            lines.append("| " + " | ".join(["---"] * len(sc_hdrs)) + " |")
            for row in sc.get("rows", []):
                lines.append(format_md_table_row(row))
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
        text = str(card.get("text", "")).strip()
        if not text:
            continue
        if "{{c" not in text:
            text = f"{{{{c1::{text}}}}}"
        text_clean = text.replace("\t", " ").replace("\n", "<br>")
        extra = str(card.get("extra", "")).replace("\t", " ").replace("\n", "<br>")
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
        text = str(c.get("text", "")).strip()
        if not text:
            continue
        # Ensure cloze syntax exists so Anki creates reviewable cards
        if "{{c" not in text:
            text = f"{{{{c1::{text}}}}}"
        extra = str(c.get("extra", ""))
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
# CLI Pipeline Orchestration
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Unified Medical Summary & Revision Pack Builder")
    parser.add_argument("--json", required=True, help="Path to structured input JSON containing summary data")
    parser.add_argument("--out-dir", default="dist", help="Output directory for generated files")
    parser.add_argument("--prefix", default="summary_pack", help="File prefix for outputs")
    args = parser.parse_args()

    if not os.path.exists(args.json):
        print(f"[ERROR] Input JSON file not found: {args.json}")
        sys.exit(1)

    with open(args.json, "r", encoding="utf-8") as f:
        data = json.load(f)

    os.makedirs(args.out_dir, exist_ok=True)
    docx_file = os.path.join(args.out_dir, f"{args.prefix}.docx")
    md_file = os.path.join(args.out_dir, f"{args.prefix}.md")
    tsv_file = os.path.join(args.out_dir, f"{args.prefix}_anki.tsv")
    apkg_file = os.path.join(args.out_dir, f"{args.prefix}_anki.apkg")

    print(f"--- Starting Medical Summary Build: {data.get('topic_title', 'Summary')} ---")
    build_docx_summary(data, docx_file)
    build_markdown_summary(data, md_file)
    build_anki_tsv(data, tsv_file)
    build_anki_apkg(data, apkg_file)
    print("--- Build Finished Successfully! ---")


if __name__ == "__main__":
    main()
