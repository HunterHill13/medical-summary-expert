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

import glob
from parse_input_lecture import ingest_file, partition_by_headings, ingest_multiple_files
from audit_summary_coverage import audit_coverage
from build_summary_pack import (
    build_docx_summary,
    build_markdown_summary,
    build_anki_tsv,
    build_anki_apkg,
    build_html_summary,
    build_pdf_summary,
    ensure_master_matrices
)

def cmd_ingest(args):
    raw_inputs = []
    if getattr(args, "inputs", None):
        raw_inputs.extend(args.inputs)
    if getattr(args, "input", None):
        raw_inputs.append(args.input)

    input_files = []
    for item in raw_inputs:
        matched = glob.glob(item)
        if matched:
            input_files.extend(matched)
        elif os.path.exists(item):
            input_files.append(item)
    input_files = sorted(list(set(input_files)))

    if len(input_files) > 1:
        print(f"--- [INGEST] Ingesting {len(input_files)} source files as multi-handout course ---")
        data = ingest_multiple_files(input_files, img_out_dir=args.img_dir)
    elif input_files:
        print(f"--- [INGEST] Ingesting source file: {input_files[0]} ---")
        data = ingest_file(input_files[0], img_out_dir=args.img_dir)
    else:
        print(f"[ERROR] No valid source file found from: {raw_inputs}")
        sys.exit(1)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OK] Ingestion complete -> {args.out}")
    print(f"     Elements: {len(data['elements'])}, Images: {len(data['images'])}, Words: {len(data['full_text'].split())}")

BATCH_PROMPT_TEMPLATE = """# دستور کار سنتز خلاصه پزشکی فیزیوپاتولوژی (بچ شماره {batch_idx:02d})

شما یک متخصص ارشد سنتز خلاصه جزوات پزشکی برای دانشجویان فیزیوپاتولوژی (Physiopathology) و آزمون‌های بالینی هستید.
وظیفه شما تبدیل متن خام این بخش به یک فایل JSON معتبر ساختاریافته به نام `summary_shard_{batch_idx:02d}.json` است.

## 📋 اصول ده‌گانه تلخیص، انسجام و پرهیز از حشو (High-Yield Editorial Directives)
1. **پرهیز مطلق از مقدمه‌چینی، توضیح درباره روند کار، نتیجه‌گیری، نظر شخصی، تکرار مطالب و حشو:** از هرگونه تعارف، فرا-کلام، نتیجه‌گیری کلیشه‌ای و نظر شخصی به شدت خودداری کنید.
2. **استخراج انحصاری عصاره علمی و نکات آزمونی:** فقط عصاره علمی و نکات قابل سؤال/قابل آزمون پزشکی را استخراج کنید.
3. **وفاداری ۱۰۰٪ به منبع و پوشش کامل رفرنس (Strict Grounding & Zero Omission):** کلیه فکت‌ها، کات‌آف‌ها، نکات دامی، و شرح‌های رفرنس یا هاریسون موجود در اسلایدها و جزوه استاد باید ۱۰۰٪ بدون حذف حفظ شوند. از اضافه کردن مطالب خارج از فایل‌های ورودی خودداری نمایید.
4. **معماری جدول‌محور فراگیر بر پایه الگوی رفرنس (Rule 22: Exemplar Full Table-Centric):**
   - تمامی مطالب سرفصل‌ها را در قالب **بخش‌های شماره‌دار جدول‌محور متوالی (`بخش ۱: ...`, `بخش ۲: ...`)** با ساختار ۳ الی ۴ ستونه استاندارد تنظیم کنید. هیچ پاراگراف انشایی بیرون از جداول رها نشود.
   - **کهن‌الگوهای استاندارد ستون‌ها:**
     * پاتوفیزیولوژی و علل (۳ ستونه): `["مکانیسم پاتوفیزیولوژیک", "علل و اتیولوژی‌های شایع", "نمونه‌های بالینی و نکات امتحانی"]`
     * کلاس‌های بیوشیمیایی هورمون‌ها (۳ ستونه): `["کلاس بیوشیمیایی", "هورمون‌های شاخص", "ویژگی‌های اختصاصی و نکات امتحانی"]`
     * گیرنده‌ها و مسیرهای پیام‌رسان (۴ ستونه): `["کلاس گیرنده", "مسیر و پیام‌رسان ثانویه", "هورمون‌های الگو", "اهمیت و ارتباط بالینی"]`
     * رویکرد بالینی و خطاهای رایج (۳ ستونه): `["عنوان مبحث بالینی", "شرح و مکانیسم فیزیولوژیک", "اهمیت در آزمون و خطاهای رایج"]`
     * پروتئین‌های حامل و کینتیک (۳ ستونه): `["پروتئین حامل / مفهوم", "لیگاند اختصاصی", "قوانین فارماکوکینتیک و تفسیر بالینی امتحانی"]`
     * آناتومی بالینی و جراحی (۳ ستونه): `["عنصر آناتومیک", "موقعیت و ساختار", "پیامد بالینی آسیب و نکات جراحی"]`
     * فیزیولوژی و اختلال هورمونی (۳ ستونه): `["عنوان مبحث", "ویژگی فیزیولوژیک و فارماکولوژیک", "تظاهرات بالینی و نکات کنکوری"]`
     * تست‌های دینامیک تشخیصی (۴ ستونه): `["نوع تست دینامیک", "اندیکاسیون بالینی", "پروتکل و داروی محرک/مهار", "تفسیر تشخیصی امتحانی"]`
5. **نگارش ساختاریافته تلگرافی بدون سقف عددی کلمات (First Aid Zero-Omission Style):**
   - هیچ سقف کلمه‌ای عددی ثابتی (نظیر ۲۵ کلمه) اعمال نمی‌شود تا هیچ معیار، کات‌آف یا شرح رفرنسی حذف نگردد.
   - **منع مطلق جملات انشایی داستان‌گونه:** آوردن جملات طویل با ساختار «فاعل + مفعول + فعل ربطی» درون سلول‌ها اکیداً ممنوع است و باید به ساختار تلگرافی با فلش‌های علی-معلولی (`->`) تبدیل شود.
   - **فرمول اجباری نگارش های‌یلد:** `• **کلیدواژه:** پیامد بالینی با نماد -> و کات‌آف عددی کامل`
   - *مثال مطلوب:* `• **تظاهر اصلی:** آمنوره اولیه + قد کوتاه (< صدک ۵) -> ناشی از دیس‌ژنز گنادی و هاپلوتایپ SHOX`
6. **تفکیک فرآیندهای پویا از جداول ایستا (Rule 19: Process Flowcharts vs Static Tables):**
   - برای فرآیندهای زمانی و چرخه‌های فیدبک هورمونی، به جای جداول ۳ ستونه متنی، از **فلوچارت عمودی Mermaid** (`flowchart`) استفاده کنید.
   - جداول صرفاً به «ماتریس‌های تشخیص افتراقی (Differential Matrices)»، «کات‌آف‌های آزمایشگاهی و استیجینگ»، و «داروشناسی» اختصاص دارند.
7. **قانون عدم تکرار بین لایه‌ها (Rule 18: Cross-Layer Deduplication Gate):**
   - اگر یک بیماری یا سندرم در قالب یک جدول تماتیک یا ماتریس افتراقی در `sections` گردآوری شده است، نباید مجدداً در `modules` با پاراگراف‌های طویل رونویسی شود! در `modules` فقط نکات فراتر از جدول (سمیولوژی، اقدامات اورژانسی، هشدارهای استاد و جمعیت‌های ویژه) درج گردند.
8. **مقیاس‌پذیری پویا بر پایه حجم جزوه (Adaptive Content Scaling):**
   - طول خلاصه و تعداد جداول، متناسب با گستردگی و سرفصل‌های جزوه منبع است و هیچ سقف صفحه یا محدودیت کلمه ثابتی تحمیل نمی‌شود تا تمام نکات علمی تدریس‌شده کاملاً و بدون حذف پوشش یابند.
9. **حفظ دقیق توالی زمانی و سرفصل‌های جزوه (Strict Chronological Sequence):** ترتیب بخش‌ها دقیقاً مطابق اسلایدهای تدریس باشد.
10. **جهت‌داری راست‌به‌چپ مطلق و ایزولاسیون BiDi:** تمام جداول و سلول‌ها راست‌به‌چپ بوده و عبارات لاتین و فرمول‌ها به صورت LTR ایزوله شوند.

---

## 🌟 ساختار یکپارچه فایل خروجی (Unified JSON Shard Schema)

فایل خروجی باید حاوی هر دو کلید `sections` (برای جداول و فلوچارت‌های تماتیک) و `modules` (برای شناسنامه‌های بالینی و مکمل) باشد:

```json
{{
  "batch_index": {batch_idx},
  "sections": [
    {{
      "section_id": "sec_01",
      "section_title": "ماتریس افتراقی یا جدول تماتیک مبحث (مثلاً ماتریس افتراقی آمنوره و اختلالات بلوغ)",
      "headers": ["بیماری / سندرم", "پاتوفیزیولوژی و ژنتیک", "تظاهرات بالینی و استیگما", "تست‌های آزمایشگاهی و کات‌آف", "اقدام درمانی و تله امتحانی"],
      "rows": [
        [
          "• **سندروم ترنر (45,X0)**",
          "• **ژنتیک:** مونوپلوئیدی X -> نقص ژن SHOX -> کوتاهی قد\\n• **پاتوفیزیولوژی:** آترزی تسریع‌شده فولیکول‌ها -> Streak Gonad",
          "• **تظاهر اصلی:** آمنوره اولیه + فقدان تکامل بلوغ\\n• **استیگمای فیزیکی:** Webbed Neck + قفسه سینه سپری + زاویه والگوس آرنج",
          "• **تست تشخیصی:** کاریوتایپ خون محیطی (استاندارد طلایی)\\n• **هورمون‌ها:** FSH/LH شدیداً بالا (> 40 mIU/mL) + استرادیول پایین (< 20 pg/mL)",
          "• **درمان:** هورمون رشد در کودکی -> استروژن در ۱۲ سالگی -> پروژستین پس از منارک\\n• 🔴 **تله امتحانی:** کوآرکتاسیون آئورت و BAV؛ قبل بارداری با تخمک اهدایی اکوکاردیوگرافی الزامی است."
        ]
      ],
      "flowchart": "graph TD\\n    A[شکایت آمنوره اولیه] --> B{{ارزیابی فنوتیپ و صفات ثانویه}}\\n    B -->|صفات غایب| C[سنجش FSH و LH سرم]"
    }}
  ],
  "modules": [
    {{
      "type": "illness_script",
      "name": "سندروم ترنر",
      "name_en": "Turner Syndrome",
      "definition_and_etiology": ["شایع‌ترین علت دیس‌ژنز گنادی با فنوتیپ زنانه"],
      "emergency_management": ["غربالگری اتساع و دایسکشن آئورت در شرایط بحرانی"],
      "special_populations_and_pregnancy": ["بارداری با تخمک اهدایی تنها با تایید سلامت قلب و عروق مجاز است."],
      "clinical_semiology": [{{"maneuver_or_sign": "معاینه استیگمای ترنر", "technique": "ارزیابی گردن پره‌دار و زاویه حمل آرنج (Cubitus Valgus)", "clinical_significance": "جهت‌دهی به کاریوتایپینگ فوری"}}],
      "contraindications_and_warnings": ["مصرف استروژن پیش از هورمون رشد منجر به بسته‌شدن زودرس اپی‌فیزها و کوتاهی ماندگار قد می‌شود."],
      "instructor_exam_alerts": ["تأکید اکید استاد: شایع‌ترین ناهنجاری قلبی همراه کوآرکتاسیون آئورت است."]
    }}
  ],
  "anki_flashcards": [
    {{
      "card_type": "cloze",
      "text": "در سندروم ترنر، استاندارد طلایی تشخیصی {{{{c1::کاریوتایپ خون محیطی}}}} است.",
      "extra": "بررسی مونوپلوئیدی 45,X0 یا موزائیسم.",
      "tag": "Endocrinology::Turner"
    }},
    {{
      "card_type": "basic",
      "question": "در بیمار مبتلا به DKA، نخستین و حیاتی‌ترین اقدام درمانی نجات‌بخش چیست؟",
      "answer": "سرم‌تراپی تهاجمی با نرمال سالین ۰.۹٪ پیش از شروع انسولین",
      "extra": "اصلاح هیپوولمی شدید و برقراری پرفیوژن کلیوی پیش‌نیاز تجویز ایمن انسولین است.",
      "tag": "Endocrinology::DKA"
    }}
  ]
}}
```

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

    merged_sections = []
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

        for sec in shard_data.get("sections", []):
            merged_sections.append(sec)
        for sec in shard_data.get("thematic_sections", []):
            merged_sections.append(sec)

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

    final_data = {
        "topic_title": topic_title or "خلاصه مباحث جزوه",
        "topic_title_en": topic_title_en,
        "organ_system": organ_system or "فیزیوپاتولوژی بالینی",
        "modules": merged_modules,
        "anki_flashcards": merged_flashcards
    }

    if merged_sections:
        final_data["sections"] = merged_sections

    if merged_semiology:
        final_data["clinical_semiology"] = merged_semiology

    if merged_matrices:
        final_data["comparative_matrices"] = merged_matrices

    if merged_checklist:
        final_data["morning_exam_checklist"] = merged_checklist

    # Global Synthesis: Ensure Master Pharmacology Matrix, Crash Sheet, and Morning Checklist
    ensure_master_matrices(final_data)

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
    inst = getattr(args, "institution", "دانشگاه بق")
    build_pdf_summary(data, pdf_path, html_path=html_path, course_name=data.get("organ_system", "فیزیوپاتولوژی بالینی"), institution=inst)
    build_markdown_summary(data, md_path)
    build_anki_tsv(data, tsv_path)
    build_anki_apkg(data, apkg_path)
    print(f"[OK] Build complete! Deliverables (DOCX, HTML, PDF, MD, TSV, APKG) ready in '{args.out_dir}/'.")

def cmd_bundle(args):
    """
    Bundles multiple session summary JSONs into a comprehensive master study guide booklet:
    1. Master PDF via MedicalPDFDesigner:
       - Full-page publication Cover page
       - Table of Contents
       - All chapters with chapter-specific dynamic headers and continuous 1-to-N page numbering
    2. Master Word DOCX
    3. Master Markdown (.md)
    4. Master Anki (.apkg and .tsv)
    """
    import glob
    from pdf_designer_engine import MedicalPDFDesigner

    input_files = []
    for item in args.inputs:
        matched = glob.glob(item)
        if matched:
            input_files.extend(matched)
        elif os.path.exists(item):
            input_files.append(item)

    input_files = sorted(list(set(input_files)))
    if not input_files:
        print(f"[ERROR] No valid JSON input files found from: {args.inputs}")
        sys.exit(1)

    print(f"=== Bundling {len(input_files)} Lecture Sessions into Comprehensive Master Course Booklet ===")
    os.makedirs(args.out_dir, exist_ok=True)

    chapters_data = []
    for fpath in input_files:
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                c_data = json.load(f)
                chapters_data.append(c_data)
        except Exception as e:
            print(f"[WARN] Failed to load {fpath}: {e}")

    if not chapters_data:
        print("[ERROR] No valid chapter data loaded!")
        sys.exit(1)

    course_title = args.course_title or "فیزیوپاتولوژی جامع"
    institution = args.institution or "دانشگاه بق"
    prefix = args.prefix or "comprehensive_master_booklet"

    # 1. Master PDF via MedicalPDFDesigner
    pdf_path = os.path.join(args.out_dir, f"{prefix}.pdf")
    designer = MedicalPDFDesigner(pdf_path, course_title=course_title, institution=institution)
    designer.add_cover_page(title=course_title, subtitle="مجموعه کامل و جامع درسنامه‌ها و مرور سریع آزمون", faculty="دانشکده پزشکی")
    designer.add_table_of_contents(chapters_data, title="فهرست فصول و مباحث درسنامه")

    for ch in chapters_data:
        designer.add_chapter_summary(ch)

    tmp_designer_dir = os.path.join(args.out_dir, "tmp_bundle_pdf")
    designer.build(tmp_dir=tmp_designer_dir)
    print(f"[OK] Master PDF Booklet generated: {pdf_path}")

    # 2. Master DOCX
    docx_path = os.path.join(args.out_dir, f"{prefix}.docx")
    master_docx_data = {
        "is_multi_chapter": True,
        "topic_title": course_title,
        "topic_title_en": "Comprehensive Medical Study Guide",
        "organ_system": institution,
        "chapters": chapters_data,
        "modules": [],
        "sections": [],
        "differential_tables": [],
        "high_yield_pharmacology": [],
        "morning_exam_checklist": []
    }
    for ch in chapters_data:
        master_docx_data["modules"].extend(ch.get("modules", []))
        master_docx_data["sections"].extend(ch.get("sections", []) or ch.get("thematic_sections", []))
        master_docx_data["differential_tables"].extend(ch.get("differential_tables", []))
        master_docx_data["high_yield_pharmacology"].extend(ch.get("high_yield_pharmacology", []))
        master_docx_data["morning_exam_checklist"].extend(ch.get("morning_exam_checklist", []))

    build_docx_summary(master_docx_data, docx_path)
    print(f"[OK] Master DOCX generated: {docx_path}")

    # 3. Master Markdown (.md)
    md_path = os.path.join(args.out_dir, f"{prefix}.md")
    build_markdown_summary(master_docx_data, md_path)
    print(f"[OK] Master Markdown generated: {md_path}")

    # 4. Master HTML (.html)
    html_path = os.path.join(args.out_dir, f"{prefix}.html")
    build_html_summary(master_docx_data, html_path)
    print(f"[OK] Master HTML generated: {html_path}")

    # 5. Master Anki Deck (.tsv & .apkg)
    tsv_path = os.path.join(args.out_dir, f"{prefix}_anki.tsv")
    apkg_path = os.path.join(args.out_dir, f"{prefix}_anki.apkg")
    master_anki_cards = []
    for ch in chapters_data:
        master_anki_cards.extend(ch.get("anki_flashcards", []))
    master_docx_data["anki_flashcards"] = master_anki_cards

    build_anki_tsv(master_docx_data, tsv_path)
    build_anki_apkg(master_docx_data, apkg_path)
    print(f"[OK] Master Anki Decks generated: {apkg_path} and {tsv_path}")

    print(f"\n=======================================================")
    print(f"🎉 COMPREHENSIVE COURSE MASTER BOOKLET READY in '{args.out_dir}/'!")
    print(f"PDF: {pdf_path}")
    print(f"DOCX: {docx_path}")
    print(f"Markdown: {md_path}")
    print(f"HTML: {html_path}")
    print(f"Anki: {apkg_path}")
    print(f"=======================================================\n")

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
    p_ingest.add_argument("--input", default=None, help="Input lecture file path")
    p_ingest.add_argument("--inputs", nargs="*", default=None, help="Multiple input lecture file paths or wildcards")
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
    p_build.add_argument("--institution", default="دانشگاه بق", help="Institution name for PDF header stamping")

    # Bundle
    p_bundle = subparsers.add_parser("bundle", help="Bundle multiple session JSONs into a comprehensive master course booklet")
    p_bundle.add_argument("--inputs", nargs="+", required=True, help="List of session JSON file paths or wildcards (e.g. sessions/*.json)")
    p_bundle.add_argument("--course-title", default="فیزیوپاتولوژی جامع", help="Course title for the master booklet")
    p_bundle.add_argument("--institution", default="دانشگاه بق", help="Institution name for the booklet and header")
    p_bundle.add_argument("--out-dir", default="dist/master_bundle", help="Output directory")
    p_bundle.add_argument("--prefix", default="comprehensive_master_course", help="File prefix")

    # Auto
    p_auto = subparsers.add_parser("auto", help="Run end-to-end pipeline")
    p_auto.add_argument("--input", default=None, help="Input lecture file path")
    p_auto.add_argument("--inputs", nargs="*", default=None, help="Multiple input lecture file paths or wildcards")
    p_auto.add_argument("--summary-json", default=None, help="Pre-synthesized summary JSON")
    p_auto.add_argument("--out-dir", default="dist", help="Output directory")
    p_auto.add_argument("--prefix", default="summary_pack", help="File prefix")
    p_auto.add_argument("--institution", default="دانشگاه بق", help="Institution name for PDF header stamping")

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
    elif args.subcommand == "bundle":
        cmd_bundle(args)
    elif args.subcommand == "auto":
        cmd_auto(args)

if __name__ == "__main__":
    main()
