# 🩺 Medical Summary Expert (v3.4)
## Publication-Grade Medical Lecture Notes & Study Guide Summarization Engine

[🇮🇷 راهنمای فارسی](#-راهنمای-جامع-فارسی) | [🇬🇧 English Documentation](#-english-overview)

[![Tests](https://img.shields.io/badge/Tests-19%20Passed%20(100%25)-success?style=for-the-badge&logo=pytest&logoColor=white)](tests/test_full_skill_diagnostics.py)
[![Version](https://img.shields.io/badge/Version-3.4.0-orange?style=for-the-badge)](SKILL.md)
[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Deliverable](https://img.shields.io/badge/Deliverable-Triple--Pack%20(Word%20|%20MD%20|%20Anki)-purple?style=for-the-badge)](dist/)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![Platform](https://img.shields.io/badge/Google%20Antigravity-Compatible-blueviolet?style=for-the-badge)](https://antigravity.google)

A publication-grade, automated medical knowledge summarization and high-yield study guide generator tailored for medical students in **Physiopathology (فیزیوپاتولوژی)** and clinical rotations. It ingests complex multi-format handouts (`.docx`, `.pdf`, `.md`, `.txt`), enforces a strict **Zero-Omission Verification Gate**, structures medical topics using canonical **Illness Scripts**, and outputs the synchronized **Triple-Pack**:
1. **Microsoft Word Study Guide (`.docx`)**: Custom RTL typography, Dubai font, color-coded callouts, comparative tables, and Master Pharmacology Matrix.
2. **Formatted Markdown (`.md`)**: GitHub/Obsidian-compatible Mermaid flowcharts, collapsible clinical vignettes (`<details>`), and escaped table formatting.
3. **Anki Spaced Repetition Decks (`.tsv` & native `.apkg`)**: Wozniak's 20 Rules compliant, atomic cloze cards (`model_type=CLOZE`), and deterministic prompt-hash GUIDs.

---

## ⛔ The Core Law: Strict Source Grounding (Zero Extrapolation)

> ### ⚠️ قانون حیاتی و خط قرمز: وفاداری ۱۰۰٪ به منبع تدریس‌شده
> **هیچ مطلبی از کتب مرجع خارجی یا حافظه پارامتریک مدل به خلاصه اضافه نمی‌شود.**
> خلاصه نهایی دقیقاً و منحصراً بر پایه فایل‌های داده‌شده (جزوه یا اسلاید استاد) تدوین می‌گردد. در صورت عدم ذکر یک تست تشخیصی، کات‌آف یا رژیم درمانی در جزوه، آن فیلد صریحاً `null` مانده و مدل حق ابداع یا تعمیم دادن آن را ندارد.

---

## 🌟 Key Innovations & Architectural Pillars

```
┌────────────────────────────────────────────────────────────────────────┐
│                   MEDICAL SUMMARY EXPERT PIPELINE                      │
├────────────────────────────────────────────────────────────────────────┤
│ 1. INGESTION ENGINE (DOCX & PDF)                                       │
│    - OOXML Deep Traversal: Word Textboxes, Table-Cell Shapes & SDT     │
│    - Natural Word Flow: <w:br>, <w:cr>, <w:tab>, <w:noBreakHyphen>     │
│    - PyMuPDF Vector Extraction + 150 DPI Scanned Page OCR Fallback     │
├────────────────────────────────────────────────────────────────────────┤
│ 2. SHARD-AND-MERGE MULTI-AGENT ORCHESTRATION                           │
│    - 2500-Word Semantic Partitioning by Section Headings               │
│    - 150-Word Context Halo preventing boundary amnesia                 │
│    - Worker Subagent Prompts generating atomic JSON shards             │
├────────────────────────────────────────────────────────────────────────┤
│ 3. ZERO-OMISSION CLINICAL NLP AUDIT GATE                               │
│    - Persian/Arabic Digit Canonicalization (۴۰٪ -> 40%, ۱،۵ -> 1.5)    │
│    - Thousands Separator Normalization (۵۰،۰۰۰ / 1,000,000 -> 50000)   │
│    - Substantive Contraindication & Warning Matching (WARNING_SIGNALS) │
│    - Medical Acronym Extraction filtering English prose sentence starts│
│    - Comprehensive USMLE & Katzung Pharmacological Suffix Detection    │
├────────────────────────────────────────────────────────────────────────┤
│ 4. TRIPLE-PACK COMPILATION ENGINE                                      │
│    - DOCX: Word RTL, Dubai Font, 5 Custom Callout Styles, RTL Tables   │
│    - MD: Sanitized Mermaid Flowcharts, Collapsible Vignettes, Pipe Esc │
│    - Anki: CLOZE Model, Deterministic Hash GUIDs, Auto-Synthesized Deck│
└────────────────────────────────────────────────────────────────────────┘
```

### 1. 📑 True Document-Order Ingestion & Deep OOXML Extraction
* **Nested Table-Cell Shapes & Textboxes (Bug R Fixed):** Standard python-docx `cell.text` drops drawing textboxes. Our `table_to_markdown` uses `extract_text_including_txbx` to capture every badge, shape, and callout inside tables.
* **Typographic Continuity (Bug M Fixed):** Word line breaks (`<w:br>`, `<w:cr>`), tab stops (`<w:tab>`), and non-breaking hyphens (`<w:noBreakHyphen>`) are preserved without fusing words together.
* **Structured Document Tags (`w:sdt` - Bug F Fixed):** Extracts content controls from modern university Word templates recursively.
* **PDF Vector & Raster Fallback:** Uses PyMuPDF table detection (`page.find_tables()`), deduplicating text blocks while rasterizing scanned pages at 150 DPI.

### 2. 🛡️ Intelligent Clinical NLP Coverage Audit (`audit_summary_coverage.py`)
* **English Prose Sentence-Starter Filter (Bug S Fixed):** Capitalized sentence-starters (`When`, `Because`, `During`, `First`, `Other`) are filtered via `is_medical_acronym` and comprehensive stopwords, ensuring true clinical entity recall and eliminating false-positive alarms in Persian summaries.
* **Expanded Pharmacology Suffixes:** Recognizes all major USMLE/Katzung drug families (`-lol`, `-pril`, `-sartan`, `-statin`, `-gliflozin`, `-semide`, `-actone`, `-grel`, `-xaban`, `-gliptin`, `-afil`, etc.).
* **Bilingual Number & Unit Standardization:** Normalizes Persian digits (`۰-۹`), Arabic decimal commas (`۱،۵` -> `1.5`), thousands separators (`۵۰،۰۰۰` -> `50000`), relational prefixes (`کمتر از`, `بیش از`, `زیر`, `بالای`), and clinical units (`میلی‌گرم`, `درصد`, `ساعت`, `میلی‌متر جیوه`).
* **Substantive Warning Audit:** Flags missing absolute contraindications (`کنترااندیکاسیون`, `اکیدا ممنوع`, `washout`) and professor alerts (`⭐`) using substantive lemma matching.

### 3. 🎯 Full Parity Across Deliverables (`ensure_master_matrices`)
* **Master Pharmacology Matrix:** Comparison table across all conditions: Disease | Drug Class | Cellular Mechanism | Indication | Contraindication | Exam Pearl. Auto-synthesized if missing with 100% Word and Markdown parity.
* **Last-Minute High-Yield Cheat Sheet:** Ultra-dense telegraphic review table for morning-of-exam review.
* **Markdown Table Sanitization (`format_md_table_row`):** Automatically escapes pipe characters (`|` -> `\|`) and collapses newlines inside cells, preventing table layout breakage across GitHub, Obsidian, and Typora.

### 4. 🗃️ Anki Spaced-Repetition Stability & Auto-Synthesis
* **Deterministic GUID Protection:** Card GUIDs are hashed exclusively from normalized prompt text (`genanki.guid_for(norm_prompt)`). Updating explanations or mechanisms in your summary will **never** duplicate cards or reset review intervals in Anki.
* **True Cloze Model (`model_type=1`):** Generates active cloze cards (`{{c1::...}}`) with Persian RTL CSS styling.
* **Auto-Synthesis (Bug Q Fixed):** In addition to diagnostic steps and pearls, cards are automatically synthesized from **Absolute Contraindications (⛔)** and **Instructor Exam Alerts (⭐)** with hierarchical tags (`PhysioPath::[Topic]::Contraindications`).

---

## 📁 Repository Structure

```text
├── run_summary_pipeline.py     # Main CLI orchestrator (ingest, partition, merge, audit, build, auto)
├── parse_input_lecture.py      # DOCX & PDF ingestion engine (tables, shapes, SDT, images, OCR)
├── build_summary_pack.py       # Triple-pack compiler (DOCX, Markdown, Anki TSV/APKG)
├── audit_summary_coverage.py   # Clinical NLP audit gate (numbers, cutoffs, warnings, entities)
├── SKILL.md                    # Antigravity skill specification and guidelines
├── sample_summary.json         # Standard structured JSON data (Heart Failure HFrEF vs HFpEF)
├── tests/                      # Comprehensive automated stress-test suite
│   ├── test_full_skill_diagnostics.py   # 19-test suite (Bug A to Bug S)
│   ├── test_omission_bugs.py            # Word table & textbox preservation test
│   ├── test_pdf_ingestion.py            # PyMuPDF table and scanned page test
│   └── test_digit_and_guid_bugs.py      # Persian digit normalization & Anki GUID test
└── dist/                       # Output artifacts directory
    ├── summary_pack.docx       # Publication-ready Persian Word study guide
    ├── summary_pack.md         # GitHub/Obsidian formatted Markdown document
    ├── summary_pack_anki.tsv   # Anki TSV import file
    └── summary_pack_anki.apkg  # One-click native Anki deck package
```

---

## 🚀 CLI Commands & Workflow

All steps are executed via the unified CLI orchestrator (`run_summary_pipeline.py`):

```bash
# 1. Exhaustive Ingestion of Handout (DOCX or PDF)
uv run python run_summary_pipeline.py ingest --input "lecture.pdf" --out "ingested_lecture.json"

# 2. Partition Long Handouts into Semantic Batches (for documents > 15 pages)
uv run python run_summary_pipeline.py partition --input "ingested_lecture.json" --batch-words 2500 --out-dir "batches"

# 3. Merge Subagent Generated Shards into Unified Summary JSON
uv run python run_summary_pipeline.py merge-batches --batch-dir "batches" --out "summary_data.json"

# 4. Run Clinical Zero-Omission Audit Gate
uv run python run_summary_pipeline.py audit --source "ingested_lecture.json" --summary "summary_data.json" --report "coverage_audit_report.json"

# 5. Build the Final Triple-Pack (Word + Markdown + Anki TSV/APKG)
uv run python run_summary_pipeline.py build --summary "summary_data.json" --out-dir "dist" --prefix "summary_pack"

# Auto-Pipeline (All-in-one for single handouts)
uv run python run_summary_pipeline.py auto --input "lecture.docx" --summary-json "summary_data.json"
```

---

## 🧪 Automated Test Suite (19/19 Passing)

Our hardened test suite verifies zero data loss, null resilience, and clinical accuracy:

| Test Identifier | Diagnosed Condition & Verification Gate | Result |
| :--- | :--- | :---: |
| **Bug A** | Auto-synthesizes Anki `.apkg` and `.tsv` when `anki_flashcards` array is omitted | **PASS** ✅ |
| **Bug B** | Extracts Persian clinical units (`میلی‌گرم`, `درصد`, `ساعت`) and relational prefixes | **PASS** ✅ |
| **Bug C** | Protects `add_comparison_table` against column overflow `IndexError` | **PASS** ✅ |
| **Bug D** | Interoperability across alternate pharmacology keys (`mechanism` vs `mechanism_of_action`) | **PASS** ✅ |
| **Bug E** | Suppresses orphan diagnostic labels when tests are `null` or unmentioned | **PASS** ✅ |
| **Bug F** | Unrolls Word Structured Document Tags (`w:sdt`) without content loss | **PASS** ✅ |
| **Bug G** | Complete null-resilience across DOCX, MD, TSV, APKG, and Audit builders | **PASS** ✅ |
| **Bug H** | Resilient shard merging with missing optional fields | **PASS** ✅ |
| **Bug I** | Generates valid cloze cards with `model_type=CLOZE` and RTL Persian styling | **PASS** ✅ |
| **Bug J** | Heading-based semantic batch partitioning with context halo continuity | **PASS** ✅ |
| **Bug K** | Decimals preserved across sentence splitting with exact boundary numeric matching | **PASS** ✅ |
| **Bug L** | Suppresses empty callout boxes and unpopulated comparison tables | **PASS** ✅ |
| **Bug M** | Preserves `<w:br>`, `<w:cr>`, `<w:tab>`, `<w:noBreakHyphen>` without text fusion | **PASS** ✅ |
| **Bug N** | 100% Word/Markdown parity for Master Pharmacology Matrix with Mermaid sanitization | **PASS** ✅ |
| **Bug O** | Persian stopwords filtering preventing false-negative contraindication bypass | **PASS** ✅ |
| **Bug P** | Canonical normalization of thousands separators (`۵۰،۰۰۰`) and Arabic decimal commas (`۱،۵`) | **PASS** ✅ |
| **Bug Q** | Auto-synthesizes Anki cards from absolute contraindications and instructor alerts | **PASS** ✅ |
| **Bug R** | Extracts drawing shapes and textboxes located inside Word table cells into Markdown | **PASS** ✅ |
| **Bug S** | Filters English prose sentence-starters from medical entity recall audit | **PASS** ✅ |

To run the full diagnostic suite:
```bash
uv run python tests/test_full_skill_diagnostics.py
```

---

## 🇮🇷 راهنمای جامع فارسی

### هدف اسکیل
این اسکیل تخصصی برای تبدیل جزوات اسلایدی و متنی اساتید، ترنسکریپت‌های کلاسی و هندآوت‌های درسی دانشگاه‌های علوم پزشکی به **بسته سه‌گانه مرور های‌یلد (Triple-Pack)** طراحی شده است:

1. **جزوه رسمی Word (`.docx`):** با استایل‌بندی مدرن، فونت خوانای دبی، دایرکشن راست‌به‌چپ (RTL)، کادرهای رنگی اختصاصی (زنجیره پاتوفیزیولوژی، تأکیدات مستقیم استاد، موارد مصرف ممنوع، سناریوهای خودآزمایی بالینی، نکات دامی آزمون)، ماتریس مقایسه جامع داروها و جدول فشرده شب امتحان.
2. **سند مارک‌داون (`.md`):** شامل فلوچارت‌های کدنویسی‌شده استاندارد Mermaid، خودآزمایی بالینی تاشو با پاسخ تشریحی مکانیسمی (`<details>`)، و پاک‌سازی کامل جدول‌ها با اسکپ کاراکترهای پایپ.
3. **فلش‌کارت‌های استاندارد آنکی (`.tsv` و `.apkg`):** کارت‌های جای‌خالی اتمیک بر پایه ۲۰ قانون Wozniak، بدون ریسک تکراری شدن یا ریست زمان‌بندی مرور دانشجو به دلیل بهره‌گیری از شناسه‌های پایدار قطعی (Deterministic GUIDs).

---

## 📄 لایسنس (License)

این پروژه تحت مجوز [MIT License](LICENSE) منتشر شده است. استفاده و توسعه آن برای دانشجویان و اساتید علوم پزشکی آزاد است.
