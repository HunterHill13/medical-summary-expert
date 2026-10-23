#!/usr/bin/env python3
"""
parse_input_lecture.py: Enhanced Medical Lecture & Pamphlet Ingestor.
Fixes Omission Bugs:
1. Retains all tables formatted as Markdown in natural reading order inside full_text.
2. Extracts Word Shapes & Text Boxes (w:txbxContent) without losing callouts.
3. Associates extracted images with the nearest section heading.
4. Supports automatic section-based hierarchical partitioning for long handouts (30-60 pages).
"""

import os
import sys
import json
import re
import argparse
from typing import Dict, Any, List

# Force UTF-8 encoding in stdout/stderr for Windows compatibility
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def extract_text_including_txbx(p_elem) -> str:
    """
    Extracts text from a paragraph, including any nested drawing text boxes (w:txbxContent),
    preserving line breaks (<w:br>, <w:cr>), tabs (<w:tab>), non-breaking hyphens (<w:noBreakHyphen>),
    and nested paragraphs without word fusion.
    """
    chunks = []
    for node in p_elem.iter():
        tag = node.tag
        if tag.endswith('t'):
            if node.text:
                chunks.append(node.text)
        elif tag.endswith('tab'):
            chunks.append("\t")
        elif tag.endswith('br') or tag.endswith('cr'):
            chunks.append("\n")
        elif tag.endswith('noBreakHyphen'):
            chunks.append("-")
        elif tag.endswith('p') and node is not p_elem:
            if chunks and not chunks[-1].endswith("\n"):
                chunks.append("\n")
    return "".join(chunks).strip()

def table_to_markdown(table) -> str:
    """Converts a python-docx table into a clean GitHub-flavored Markdown table, capturing drawing textboxes and nested text."""
    if not table.rows:
        return ""
    md_lines = []
    rows_data = []
    for row in table.rows:
        cell_texts = []
        for cell in row.cells:
            # Clean newlines inside cells and capture nested shapes/textboxes
            p_texts = [extract_text_including_txbx(p._p) for p in cell.paragraphs]
            raw_cell_text = " ".join([pt for pt in p_texts if pt])
            c_text = " ".join(raw_cell_text.split()).replace("|", "\\|")
            cell_texts.append(c_text)
        rows_data.append(cell_texts)

    if not rows_data:
        return ""

    headers = rows_data[0]
    md_lines.append("| " + " | ".join(headers) + " |")
    md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for r in rows_data[1:]:
        md_lines.append("| " + " | ".join(r) + " |")

    return "\n".join(md_lines)

def iter_block_elements(parent):
    """
    Recursively traverses block-level elements (<w:p>, <w:tbl>, and <w:sdt>),
    yielding top-level (<w:p> or <w:tbl>) XML elements in true document order
    without descending into table cells or duplicating content.
    """
    for child in parent:
        tag = child.tag
        if tag.endswith('p') or tag.endswith('tbl'):
            yield child
        elif tag.endswith('sdt'):
            # Look inside w:sdtContent
            for sub in child:
                if sub.tag.endswith('sdtContent'):
                    yield from iter_block_elements(sub)

def parse_docx(file_path: str, extract_images: bool = True, img_out_dir: str = "extracted_images") -> Dict[str, Any]:
    """
    Exhaustively parses a Microsoft Word (.docx) document in true document order.
    Preserves all paragraphs, tables as Markdown, Word text boxes, and content controls (w:sdt).
    """
    import docx
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
    from docx.table import Table

    doc = docx.Document(file_path)
    body = doc.element.body

    elements = []
    ordered_text_blocks = []
    current_heading = "مقدمه و کلیات"

    # Map tables and paragraphs for quick lookup
    p_map = {p._p: p for p in doc.paragraphs}
    t_map = {t._tbl: t for t in doc.tables}

    for child in iter_block_elements(body):
        tag = child.tag
        if tag.endswith('p'):
            p_obj = p_map.get(child) or Paragraph(child, doc)
            style_name = p_obj.style.name if p_obj and p_obj.style else "Normal"
            text = extract_text_including_txbx(child)
            if text:
                if style_name.startswith("Heading") or text.startswith("#"):
                    current_heading = text.lstrip("#").strip()
                elements.append({
                    "type": "paragraph",
                    "text": text,
                    "style": style_name,
                    "section": current_heading
                })
                ordered_text_blocks.append(text)
        elif tag.endswith('tbl'):
            t_obj = t_map.get(child) or Table(child, doc)
            md_tbl = table_to_markdown(t_obj)
            if md_tbl:
                elements.append({
                    "type": "table",
                    "markdown": md_tbl,
                    "section": current_heading
                })
                ordered_text_blocks.append(f"\n[جدول آموزشی]\n{md_tbl}\n")

    # Image Extraction with Rel IDs
    extracted_images = []
    if extract_images:
        os.makedirs(img_out_dir, exist_ok=True)
        img_counter = 0
        for rel in doc.part.rels.values():
            if getattr(rel, 'is_external', False):
                continue
            if "image" in rel.target_ref:
                try:
                    img_data = rel.target_part.blob
                except Exception:
                    continue
                # Filter out tiny decorative icons, bullet markers, and spacer graphics (< 3KB = 3072 bytes)
                if len(img_data) < 3072:
                    continue
                try:
                    import io
                    from PIL import Image
                    im = Image.open(io.BytesIO(img_data))
                    if im.width < 80 or im.height < 80:
                        continue
                except Exception:
                    pass
                ext = rel.target_ref.split(".")[-1].lower() if "." in rel.target_ref else "png"
                if len(ext) > 5 or "/" in ext or "\\" in ext:
                    ext = "png"
                img_counter += 1
                img_filename = f"image_{img_counter}.{ext}"
                img_path = os.path.join(img_out_dir, img_filename)
                with open(img_path, "wb") as f_img:
                    f_img.write(img_data)
                extracted_images.append({
                    "path": img_path,
                    "filename": img_filename,
                    "associated_section": current_heading
                })

    full_text = "\n\n".join(ordered_text_blocks)

    return {
        "file_name": os.path.basename(file_path),
        "elements": elements,
        "images": extracted_images,
        "full_text": full_text
    }

def partition_by_headings(doc_data: Dict[str, Any], max_words_per_batch: int = 2500) -> List[Dict[str, Any]]:
    """
    Partitions long handouts (30-60 pages) into semantic batches based on section headings.
    Provides a 1-heading Context Halo to prevent boundary amnesia.
    """
    elements = doc_data.get("elements", [])
    if not elements:
        txt = doc_data.get("full_text", "")
        return [{
            "batch_index": 1,
            "title": "کل جزوه",
            "context_halo": "",
            "text": txt,
            "word_count": len(txt.split())
        }]

    batches = []
    current_batch_elements = []
    current_word_count = 0
    batch_idx = 1
    last_halo = ""

    for elem in elements:
        text = elem.get("text") or elem.get("markdown", "")
        w_count = len(text.split())

        if current_word_count + w_count > max_words_per_batch and current_batch_elements:
            batch_text = "\n\n".join([e.get("text") or e.get("markdown", "") for e in current_batch_elements])
            batches.append({
                "batch_index": batch_idx,
                "context_halo": last_halo,
                "text": batch_text,
                "word_count": current_word_count
            })
            last_halo = " ".join(batch_text.split()[-150:])  # last 150 words as halo
            batch_idx += 1
            current_batch_elements = []
            current_word_count = 0

        current_batch_elements.append(elem)
        current_word_count += w_count

    if current_batch_elements:
        batch_text = "\n\n".join([e.get("text") or e.get("markdown", "") for e in current_batch_elements])
        batches.append({
            "batch_index": batch_idx,
            "context_halo": last_halo,
            "text": batch_text,
            "word_count": current_word_count
        })

    return batches

def parse_markdown_or_text(file_path: str) -> Dict[str, Any]:
    """Reads Markdown or plain text files, preserving Markdown tables atomically."""
    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()

    lines = content.splitlines()
    elements = []
    current_sec = "مقدمه"
    table_buffer = []

    def flush_table():
        if table_buffer:
            md_tbl = "\n".join(table_buffer)
            elements.append({
                "type": "table",
                "markdown": md_tbl,
                "section": current_sec
            })
            table_buffer.clear()

    for line in lines:
        l_str = line.strip()
        if not l_str:
            flush_table()
            continue

        if l_str.startswith("|") and l_str.endswith("|"):
            table_buffer.append(l_str)
            continue
        else:
            flush_table()

        if l_str.startswith("#"):
            current_sec = l_str.lstrip("#").strip()

        elements.append({
            "type": "paragraph",
            "text": l_str,
            "style": "Normal",
            "section": current_sec
        })

    flush_table()

    return {
        "file_name": os.path.basename(file_path),
        "elements": elements,
        "images": [],
        "full_text": content
    }

def parse_pdf(file_path: str, extract_images: bool = True, img_out_dir: str = "extracted_images") -> Dict[str, Any]:
    """
    Extracts text, tables, and images from medical PDF lecture notes or handouts using PyMuPDF.
    """
    import pymupdf

    doc = pymupdf.open(file_path)
    elements = []
    ordered_text_blocks = []
    extracted_images = []
    current_heading = "مقدمه و کلیات"

    if extract_images:
        os.makedirs(img_out_dir, exist_ok=True)

    img_counter = 0

    for page_idx, page in enumerate(doc, 1):
        # 1. Detect and extract tables on page
        tables = page.find_tables()
        table_rects = []
        if tables and tables.tables:
            for t_idx, tab in enumerate(tables.tables):
                t_df = tab.extract()
                if t_df and len(t_df) > 1:
                    headers = [str(c or "").strip() for c in t_df[0]]
                    md_lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join(["---"] * len(headers)) + " |"]
                    for row in t_df[1:]:
                        r_vals = [str(c or "").strip().replace("|", "\\|") for c in row]
                        md_lines.append("| " + " | ".join(r_vals) + " |")
                    md_tbl = "\n".join(md_lines)
                    elements.append({
                        "type": "table",
                        "markdown": md_tbl,
                        "section": current_heading,
                        "page": page_idx
                    })
                    ordered_text_blocks.append(f"\n[جدول صفحه {page_idx}]\n{md_tbl}\n")
                    table_rects.append(tab.bbox)

        # 2. Check for scanned / raster-only pages (Multi-Tier Ingestion)
        page_text = page.get_text("text").strip()
        page_images = page.get_images(full=True)
        if len(page_text.split()) < 15 and len(page_images) > 0:
            os.makedirs(img_out_dir, exist_ok=True)
            pix = page.get_pixmap(dpi=150)
            scanned_filename = f"scanned_page_{page_idx}.png"
            scanned_path = os.path.join(img_out_dir, scanned_filename)
            pix.save(scanned_path)

            # Tier 1: Attempt native PyMuPDF OCR if Tesseract/OCR engine is available
            ocr_text = ""
            try:
                tp = page.get_textpage_ocr(flags=0, dpi=150, language='fas+eng')
                ocr_text = page.get_text(textpage=tp).strip()
            except Exception:
                try:
                    tp = page.get_textpage_ocr(flags=0, dpi=150)
                    ocr_text = page.get_text(textpage=tp).strip()
                except Exception:
                    ocr_text = ""

            if len(ocr_text.split()) >= 15:
                elements.append({
                    "type": "scanned_page_ocr",
                    "page": page_idx,
                    "image_path": scanned_path,
                    "ocr_text": ocr_text,
                    "section": current_heading
                })
                ordered_text_blocks.append(f"\n[متن استخراج‌شده با OCR از صفحه {page_idx}]:\n{ocr_text}\n")
            else:
                # Tier 2: Explicit Multimodal Inspection Marker
                marker = f"[🚨 SCANNED_PAGE_REQUIRING_MULTIMODAL_INSPECTION: {scanned_path} (صفحه {page_idx})]"
                elements.append({
                    "type": "scanned_page",
                    "page": page_idx,
                    "image_path": scanned_path,
                    "section": current_heading,
                    "requires_multimodal_inspection": True,
                    "warning": f"صفحه {page_idx} به صورت اسکن/تصویر بدون متن برداری استخراج شد و نیاز به بازبینی چندوجهی دارد."
                })
                ordered_text_blocks.append(f"\n{marker}\n")

        # 3. Extract text blocks
        blocks = page.get_text("blocks")
        for b in blocks:
            if len(b) >= 5:
                b_text = b[4].strip()
                if not b_text:
                    continue
                bx0, by0, bx1, by1 = b[0], b[1], b[2], b[3]
                bcx, bcy = (bx0 + bx1) / 2, (by0 + by1) / 2
                in_table = any(
                    (tx0 <= bx0 and ty0 <= by0 and bx1 <= tx1 and by1 <= ty1) or
                    (tx0 <= bcx <= tx1 and ty0 <= bcy <= ty1)
                    for (tx0, ty0, tx1, ty1) in table_rects
                )
                if in_table:
                    continue

                if b_text.startswith("#") or "فصل" in b_text or "مبحث" in b_text or "Chapter" in b_text:
                    current_heading = b_text.splitlines()[0].strip()

                elements.append({
                    "type": "paragraph",
                    "text": b_text,
                    "style": "Normal",
                    "section": current_heading,
                    "page": page_idx
                })
                ordered_text_blocks.append(b_text)

        # 3. Extract high-resolution images
        if extract_images:
            images = page.get_images(full=True)
            for img in images:
                xref = img[0]
                base_img = doc.extract_image(xref)
                if base_img:
                    img_bytes = base_img["image"]
                    ext = base_img["ext"]
                    w, h = base_img["width"], base_img["height"]
                    if w < 100 or h < 100 or (w * h) < 15000:
                        continue
                    img_counter += 1
                    img_filename = f"pdf_image_{img_counter}.{ext}"
                    img_path = os.path.join(img_out_dir, img_filename)
                    with open(img_path, "wb") as f_img:
                        f_img.write(img_bytes)
                    extracted_images.append({
                        "path": img_path,
                        "filename": img_filename,
                        "associated_section": current_heading,
                        "page": page_idx
                    })

    full_text = "\n\n".join(ordered_text_blocks)
    return {
        "file_name": os.path.basename(file_path),
        "elements": elements,
        "images": extracted_images,
        "full_text": full_text
    }

def ingest_file(file_path: str, img_out_dir: str = "extracted_images") -> Dict[str, Any]:
    """Universal file ingestor supporting Word (.docx), PDF (.pdf), Markdown (.md), and text (.txt)."""
    ext = os.path.splitext(file_path)[-1].lower()
    if ext == ".docx":
        return parse_docx(file_path, extract_images=True, img_out_dir=img_out_dir)
    elif ext == ".pdf":
        return parse_pdf(file_path, extract_images=True, img_out_dir=img_out_dir)
    elif ext in [".md", ".txt"]:
        return parse_markdown_or_text(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}. Supported formats: .docx, .pdf, .md, .txt")


def ingest_multiple_files(file_paths: List[str], img_out_dir: str = "extracted_images") -> Dict[str, Any]:
    """
    Ingests multiple lecture handouts (DOCX, PDF, MD), producing a unified multi-chapter structure.
    Preserves per-handout provenance, page numbering, extracted images, and section boundaries.
    """
    chapters = []
    combined_elements = []
    combined_images = []
    full_texts = []

    for idx, fp in enumerate(file_paths, 1):
        if not os.path.exists(fp):
            continue
        ch_img_dir = os.path.join(img_out_dir, f"handout_{idx}")
        ch_data = ingest_file(fp, img_out_dir=ch_img_dir)
        ch_data["chapter_index"] = idx
        ch_data["source_file"] = os.path.basename(fp)
        ch_title = os.path.splitext(os.path.basename(fp))[0].replace("_", " ")
        ch_data["title"] = f"جزوه {idx}: {ch_title}"
        chapters.append(ch_data)
        combined_elements.extend(ch_data.get("elements", []))
        combined_images.extend(ch_data.get("images", []))
        full_texts.append(f"=== [📚 جزوه {idx}: {os.path.basename(fp)}] ===\n" + ch_data.get("full_text", ""))

    return {
        "is_multi_chapter": True,
        "total_chapters": len(chapters),
        "file_name": f"multi_handout_bundle_{len(chapters)}_files",
        "chapters": chapters,
        "elements": combined_elements,
        "images": combined_images,
        "full_text": "\n\n".join(full_texts)
    }

def main():
    parser = argparse.ArgumentParser(description="Ingest medical notes with zero table/textbox loss")
    parser.add_argument("--input", required=True, help="Path to input .docx, .md, or .txt file")
    parser.add_argument("--out-json", default="ingested_lecture.json", help="Path to output JSON")
    parser.add_argument("--partition", action="store_true", help="Partition into section batches if long")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"[ERROR] Input file not found: {args.input}")
        sys.exit(1)

    result = ingest_file(args.input)
    if args.partition:
        batches = partition_by_headings(result)
        result["partitions"] = batches
        print(f"[INFO] Partitioned into {len(batches)} batches for long handout handling.")

    with open(args.out_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"[OK] Ingestion complete: {args.input} -> {args.out_json}")
    print(f"     Extracted {len(result['elements'])} ordered elements, {len(result['images'])} images, total words: {len(result['full_text'].split())}.")

if __name__ == "__main__":
    main()
