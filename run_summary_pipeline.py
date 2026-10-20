#!/usr/bin/env python3
"""
run_summary_pipeline.py: Unified CLI Orchestration Pipeline for Medical Summary Expert.
Provides end-to-end command dispatching:
  1. ingest       - Parses Word (.docx), PDF (.pdf), or Markdown into clean structured JSON.
  2. partition    - Slices long lectures (>15 pages) into semantic batches with context halo.
  3. merge-batches- Merges worker subagent summary shards into a unified summary_data.json.
  4. audit        - Executes the deterministic zero-omission and cutoff coverage verification gate.
  5. build        - Compiles the Triple-Pack: Word (.docx), Markdown (.md), and Anki (.tsv/.apkg).
  6. auto         - Runs ingest, audit, and build in a single end-to-end command.
"""

import os
import sys
import json
import argparse
import subprocess
from typing import Dict, Any, List

# Force UTF-8 encoding in stdout/stderr for Windows compatibility
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from parse_input_lecture import ingest_file, partition_by_headings
from audit_summary_coverage import audit_coverage
from build_summary_pack import (
    build_docx_summary,
    build_markdown_summary,
    build_anki_tsv,
    build_anki_apkg,
    build_html_summary,
    build_pdf_summary
)

def cmd_ingest(args):
    print(f"--- [INGEST] Ingesting source file: {args.input} ---")
    data = ingest_file(args.input, img_out_dir=args.img_dir)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OK] Ingestion complete -> {args.out}")
    print(f"     Elements: {len(data['elements'])}, Images: {len(data['images'])}, Words: {len(data['full_text'].split())}")

BATCH_PROMPT_TEMPLATE = """# دستور کار سنتز خلاصه پزشکی فیزیوپاتولوژی (بچ شماره {batch_idx:02d})

شما یک متخصص ارشد سنتز خلاصه جزوات پزشکی برای دانشجویان فیزیوپاتولوژی (Physiopathology) و آزمون‌های بالینی هستید.
وظیفه شما تبدیل متن خام این بخش به یک فایل JSON معتبر ساختاریافته به نام `summary_shard_{batch_idx:02d}.json` است.

## 📋 اصول هفت‌گانه تلخیص، انسجام و پرهیز از حشو (High-Yield Editorial Directives)
1. **پرهیز مطلق از مقدمه‌چینی، توضیح درباره روند کار، نتیجه‌گیری، نظر شخصی، تکرار مطالب و حشو:** از هرگونه تعارف، فرا-کلام، نتیجه‌گیری کلیشه‌ای و نظر شخصی به شدت خودداری کنید.
2. **استخراج انحصاری عصاره علمی و نکات آزمونی:** فقط عصاره علمی و نکات قابل سؤال/قابل آزمون پزشکی را استخراج کنید.
3. **وفاداری ۱۰۰٪ به منبع و منع مطلق افزودن اطلاعات (Strict Grounding):** هیچ اطلاعاتی از خودتان، کتب رفرنس خارجی یا حافظه قبلی اضافه نکنید. اگر مطلبی در فایل ورودی وجود ندارد، آن را به خلاصه اضافه نکنید و فیلد آن را `null` یا خالی بگذارید.
4. **حفظ فشرده و منظم ساختارهای بالینی:** اگر در متن دسته‌بندی، علل، عوامل خطر، پاتوفیزیولوژی، علائم و نشانه‌ها، یافته‌های آزمایشگاهی، تشخیص، درمان، عوارض، پیش‌آگهی یا افتراق‌ها وجود دارد، آنها را در فشرده‌ترین، منظم‌ترین و جدول‌بندی‌شده‌ترین شکل ممکن حفظ کنید.
5. **حفظ روابط علت و معلولی و زنجیره پاتوفیزیولوژی:** روابط علت و معلولی و ارتباط منطقی بین مطالب را حفظ کنید؛ به‌گونه‌ای که خلاصه فقط فهرستی خشک از کلمات نباشد و برای فهم عمیق بالینی، یادگیری و مرور سریع کاملاً قابل استفاده باشد.
6. **پرهیز از بولت‌سازی افراطی و تجمیع مطالب مرتبط:** مطالبی که به یکدیگر مرتبط هستند را تا حد امکان در یک بولت یا عبارت ترکیبی منسجم تجمیع کنید و از بولت‌سازی افراطی و تقطیع بیهوده جملات خودداری کنید.
7. **خروجی یکدست، منسجم، کم‌حجم و پُرمحتوا:** خروجی باید کاملاً یکدست، منسجم، کم‌حجم، پُرمحتوا و متناسب با مرور سریع برای امتحان پزشکی باشد.

## قوانین ساختاری و گیت‌های تکمیلی
- **قانون موازنه جامعیت بالینی و فارماکولوژی (Anti-Pharmacology-Skew Law):** قبل از ورود به مباحث دارویی، تمام بیماری‌ها و سندرم‌های ذکرشده در متن (مانند شیهان، نلسون، کارنی، FHH، خطوط لوزر، آشرمن، CAIS، ۵-آلفا ردوکتاز و آپوپلکسی) با حفظ ۴ رکن (اتیولوژی، تظاهرات، تست اولیه در برابر طلایی، و تله‌های آزمونی) استخراج شوند.
- **ترک تخصصی سمیولوژی و معاینات فیزیکی (Track 4: Clinical Semiology):** مانورهای بالینی و معاینات فیزیکی (مانور بلع، رباط Berry، نشانه‌های دالریمپل و فن‌گرافه، شوستک و تروسو، دیاپازون ۱۲۸ هرتز و...) در کلید `clinical_semiology` استخراج شوند.
- **حفظ دقیق مقادیر عددی و کات‌آف‌ها (Zero Cutoff Omission):** تمام اعداد بالینی، درصدها (مانند EF <= 40%)، کات‌آف‌های آزمایشگاهی (مانند BNP < 100 pg/mL) و دوره‌های زمانی باید دقیقاً حفظ شوند.
- **فلش‌کارت‌های اتمی آنکی:** برای هر نکته کلیدی یک کارت Cloze Deletion بسازید ({{{{c1::...}}}}). تحلیل مکانیسمی را در فیلد extra قرار دهید.
- **کپسول طلایی شب امتحان و جداول تطبیقی:** نکات متمایزکننده تشخیصی را در قالب جدول تطبیقی (`comparative_matrices`) و نکات فوق‌سریع را در `morning_exam_checklist` درج کنید.

---

### [CONTEXT HALO - متن بخش قبلی صرفاً جهت آگاهی از پیوستگی مطلب - این بخش را خلاصه نکنید]
{context_halo}
### [پایان CONTEXT HALO]

---

### [متن اصلی این پارت برای خلاصه‌سازی]
{batch_text}
### [پایان متن اصلی پارت]
"""

def cmd_partition(args):
    print(f"--- [PARTITION] Slicing lecture into semantic batches ---")
    if not os.path.exists(args.input):
        print(f"[ERROR] Ingested file not found: {args.input}")
        sys.exit(1)
    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    batches = partition_by_headings(data, max_words_per_batch=args.batch_words)
    os.makedirs(args.out_dir, exist_ok=True)

    manifest = {
        "file_name": data.get("file_name", "lecture"),
        "total_batches": len(batches),
        "batches": []
    }

    for b in batches:
        b_idx = b["batch_index"]
        b_filename = f"batch_{b_idx:02d}.json"
        b_path = os.path.join(args.out_dir, b_filename)
        with open(b_path, "w", encoding="utf-8") as fb:
            json.dump(b, fb, ensure_ascii=False, indent=2)

        # Generate Subagent Worker Prompt file
        prompt_filename = f"prompt_batch_{b_idx:02d}.md"
        prompt_path = os.path.join(args.out_dir, prompt_filename)
        prompt_content = BATCH_PROMPT_TEMPLATE.format(
            batch_idx=b_idx,
            context_halo=b.get("context_halo", "ندارد (بخش اول جزوه)"),
            batch_text=b.get("text", "")
        )
        with open(prompt_path, "w", encoding="utf-8") as fp:
            fp.write(prompt_content)

        manifest["batches"].append({
            "batch_index": b_idx,
            "input_file": b_filename,
            "prompt_file": prompt_filename,
            "target_shard": f"summary_shard_{b_idx:02d}.json",
            "word_count": b["word_count"]
        })

    manifest_path = os.path.join(args.out_dir, "manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as fm:
        json.dump(manifest, fm, ensure_ascii=False, indent=2)

    print(f"[OK] Partitioned into {len(batches)} batches in '{args.out_dir}/'. Worker prompts and manifest created.")

def cmd_merge_batches(args):
    print(f"--- [MERGE] Merging summary shards from '{args.batch_dir}/' ---")
    manifest_path = os.path.join(args.batch_dir, "manifest.json")
    if not os.path.exists(manifest_path):
        print(f"[ERROR] Manifest file not found in {args.batch_dir}. Did you run partition?")
        sys.exit(1)

    with open(manifest_path, "r", encoding="utf-8") as fm:
        manifest = json.load(fm)

    merged_modules = []
    merged_flashcards = []
    merged_semiology = []
    merged_matrices = []
    merged_checklist = []
    topic_title = ""
    topic_title_en = ""
    organ_system = ""

    for b_info in manifest["batches"]:
        shard_file = os.path.join(args.batch_dir, b_info["target_shard"])
        if not os.path.exists(shard_file):
            print(f"[ERROR] Missing shard file: {shard_file}. Please ensure subagents generated all shards.")
            sys.exit(1)

        with open(shard_file, "r", encoding="utf-8") as fs:
            shard_data = json.load(fs)

        if not topic_title and "topic_title" in shard_data:
            topic_title = shard_data["topic_title"]
            topic_title_en = shard_data.get("topic_title_en", "")
            organ_system = shard_data.get("organ_system", "")

        for mod in shard_data.get("modules", []):
            merged_modules.append(mod)

        for card in shard_data.get("anki_flashcards", []):
            merged_flashcards.append(card)

        for sem in shard_data.get("clinical_semiology", []):
            merged_semiology.append(sem)

        for mat in shard_data.get("comparative_matrices", []):
            merged_matrices.append(mat)

        for chk in shard_data.get("morning_exam_checklist", []):
            merged_checklist.append(chk)

    # Master Pharmacology Matrix synthesis across all modules
    master_pharm_rows = []
    for mod in merged_modules:
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
            master_pharm_rows.append([m_name, d_class, d_mech, d_ind, d_contra, d_pearl])

    # Global Synthesis: Build unified Last-Minute Crash Sheet
    crash_rows = []
    for mod in merged_modules:
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

    final_data = {
        "topic_title": topic_title or "خلاصه مباحث جزوه",
        "topic_title_en": topic_title_en,
        "organ_system": organ_system or "فیزیوپاتولوژی بالینی",
        "modules": merged_modules,
        "last_minute_crash_sheet": {
            "headers": ["بیماری / مبحث", "اتیولوژی اصلی", "مکانیسم پاتوفیزیولوژی کلیدی", "علامت شاخص / تریاد", "اقدام تشخیصی نخستین", "درمان انتخابی"],
            "rows": crash_rows
        },
        "anki_flashcards": merged_flashcards
    }

    if merged_semiology:
        final_data["clinical_semiology"] = merged_semiology

    if merged_matrices:
        final_data["comparative_matrices"] = merged_matrices

    if merged_checklist:
        final_data["morning_exam_checklist"] = merged_checklist

    if len(master_pharm_rows) >= 2:
        final_data["master_pharmacology_matrix"] = {
            "title": "ماتریس جامع مقایسه داروشناسی بالینی (Master Pharmacology Matrix)",
            "headers": ["مبحث / بیماری", "دسته دارویی و نمونه‌ها", "مکانیسم اثر سلولی", "اندیکاسیون دقیق", "کنترااندیکاسیون / عوارض مرگبار", "نکته کلیدی استاد / دام آزمون"],
            "rows": master_pharm_rows
        }

    with open(args.out, "w", encoding="utf-8") as f_out:
        json.dump(final_data, f_out, ensure_ascii=False, indent=2)

    print(f"[OK] Successfully merged {len(merged_modules)} modules, {len(merged_semiology)} semiology sections, and {len(merged_flashcards)} flashcards -> {args.out}")

def cmd_audit(args):
    print(f"--- [AUDIT] Running deterministic zero-omission and cutoff coverage verification ---")
    if not os.path.exists(args.source) or not os.path.exists(args.summary):
        print(f"[ERROR] Source or summary file not found.")
        sys.exit(1)

    with open(args.source, "r", encoding="utf-8", errors="replace") as fs:
        src_raw = fs.read()
    try:
        src_data = json.loads(src_raw)
        source_text = src_data.get("full_text", src_raw)
    except Exception:
        source_text = src_raw

    with open(args.summary, "r", encoding="utf-8") as fsum:
        summary_data = json.load(fsum)

    report = audit_coverage(source_text, summary_data)
    with open(args.report, "w", encoding="utf-8") as frep:
        json.dump(report, frep, ensure_ascii=False, indent=2)

    print(f"  • Verdict: {report['verdict']}")
    print(f"  • Numerical Cutoffs Retained: {report['scores']['numerical_coverage_pct']}%")
    print(f"  • Clinical Entities Retained: {report['scores']['entity_recall_pct']}%")
    print(f"  • Diagnostic Bifurcation:     {'Passed' if report['scores']['diagnostic_bifurcation_compliant'] else 'Failed'}")
    print(f"  • Illness Script Integrity:   {report['scores']['illness_script_compliance_pct']}% ({'Passed' if report['scores']['illness_script_compliant'] else 'Failed'})")
    if report["scores"].get("pharmacology_skew_detected"):
        print(f"  ⚠️ Warning: Severe pharmacology skew detected! Clinical scripts eclipsed by drug tables.")
    if report.get("illness_script_audit", {}).get("omitted_subtle_syndromes"):
        print(f"  ⛔ Alert: Subtle clinical syndromes omitted: {report['illness_script_audit']['omitted_subtle_syndromes']}")
    if report["omitted_clinical_numbers"]:
        print(f"  ⚠️ Warning: {len(report['omitted_clinical_numbers'])} cutoffs were omitted!")
    if report.get("omitted_contraindications"):
        print(f"  ⛔ Severe Warning: {len(report['omitted_contraindications'])} contraindications were omitted!")
    print(f"[OK] Audit report written to: {args.report}")

def cmd_build(args):
    print(f"--- [BUILD] Compiling Publication-Grade Deliverables Suite ---")
    if not os.path.exists(args.summary):
        print(f"[ERROR] Summary JSON file not found: {args.summary}")
        sys.exit(1)

    with open(args.summary, "r", encoding="utf-8") as f:
        data = json.load(f)

    os.makedirs(args.out_dir, exist_ok=True)
    docx_path = os.path.join(args.out_dir, f"{args.prefix}.docx")
    html_path = os.path.join(args.out_dir, f"{args.prefix}.html")
    pdf_path = os.path.join(args.out_dir, f"{args.prefix}.pdf")
    md_path = os.path.join(args.out_dir, f"{args.prefix}.md")
    tsv_path = os.path.join(args.out_dir, f"{args.prefix}_anki.tsv")
    apkg_path = os.path.join(args.out_dir, f"{args.prefix}_anki.apkg")

    build_docx_summary(data, docx_path)
    build_html_summary(data, html_path)
    build_pdf_summary(data, pdf_path, html_path=html_path, course_name=data.get("organ_system", "فیزیوپاتولوژی بالینی"))
    build_markdown_summary(data, md_path)
    build_anki_tsv(data, tsv_path)
    build_anki_apkg(data, apkg_path)
    print(f"[OK] Build complete! Deliverables (DOCX, HTML, PDF, MD, TSV, APKG) ready in '{args.out_dir}/'.")

def cmd_auto(args):
    print(f"=== [AUTO] Running End-to-End Medical Summary Pipeline ===")
    out_dir = args.out_dir or "dist"
    os.makedirs(out_dir, exist_ok=True)
    ingested_file = os.path.join(out_dir, "ingested_lecture.json")
    audit_report = os.path.join(out_dir, "coverage_audit_report.json")

    # Step 1: Ingest
    args.img_dir = os.path.join(out_dir, "extracted_images")
    args.out = ingested_file
    cmd_ingest(args)

    # Step 2: Audit (if summary-json is provided)
    if args.summary_json and os.path.exists(args.summary_json):
        args.source = ingested_file
        args.summary = args.summary_json
        args.report = audit_report
        cmd_audit(args)

        # Step 3: Build
        args.out_dir = out_dir
        args.prefix = args.prefix or "summary_pack"
        cmd_build(args)
    else:
        print(f"[INFO] Input ingested to '{ingested_file}'. Next step: Synthesize summary_data.json and run 'run_summary_pipeline.py build'.")

def main():
    parser = argparse.ArgumentParser(description="Medical Summary Expert CLI Pipeline")
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    # Ingest
    p_ingest = subparsers.add_parser("ingest", help="Ingest DOCX, PDF, or Markdown")
    p_ingest.add_argument("--input", required=True, help="Input lecture file path")
    p_ingest.add_argument("--out", default="ingested_lecture.json", help="Output JSON path")
    p_ingest.add_argument("--img-dir", default="extracted_images", help="Directory for extracted figures")

    # Partition
    p_part = subparsers.add_parser("partition", help="Partition ingested lecture into batches")
    p_part.add_argument("--input", default="ingested_lecture.json", help="Ingested lecture JSON")
    p_part.add_argument("--batch-words", type=int, default=2500, help="Max words per batch")
    p_part.add_argument("--out-dir", default="batches", help="Output directory for shards")

    # Merge
    p_merge = subparsers.add_parser("merge-batches", help="Merge worker summary shards into summary_data.json")
    p_merge.add_argument("--batch-dir", default="batches", help="Directory containing manifest and shards")
    p_merge.add_argument("--out", default="summary_data.json", help="Merged output JSON")

    # Audit
    p_audit = subparsers.add_parser("audit", help="Audit summary against source for omissions")
    p_audit.add_argument("--source", required=True, help="Source ingested JSON or raw text")
    p_audit.add_argument("--summary", required=True, help="Summary JSON")
    p_audit.add_argument("--report", default="coverage_audit_report.json", help="Audit report path")

    # Build
    p_build = subparsers.add_parser("build", help="Compile Word, Markdown, and Anki files")
    p_build.add_argument("--summary", required=True, help="Summary JSON path")
    p_build.add_argument("--out-dir", default="dist", help="Output directory")
    p_build.add_argument("--prefix", default="summary_pack", help="File prefix")

    # Auto
    p_auto = subparsers.add_parser("auto", help="Run end-to-end pipeline")
    p_auto.add_argument("--input", required=True, help="Input lecture file path")
    p_auto.add_argument("--summary-json", default=None, help="Pre-synthesized summary JSON")
    p_auto.add_argument("--out-dir", default="dist", help="Output directory")
    p_auto.add_argument("--prefix", default="summary_pack", help="File prefix")

    args = parser.parse_args()

    if args.subcommand == "ingest":
        cmd_ingest(args)
    elif args.subcommand == "partition":
        cmd_partition(args)
    elif args.subcommand == "merge-batches":
        cmd_merge_batches(args)
    elif args.subcommand == "audit":
        cmd_audit(args)
    elif args.subcommand == "build":
        cmd_build(args)
    elif args.subcommand == "auto":
        cmd_auto(args)

if __name__ == "__main__":
    main()
