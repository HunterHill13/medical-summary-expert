# 🩺 Medical Summary Expert (v3.9.9)
## Publication-Grade Medical Lecture Notes & Study Guide Summarization Engine

[🇮🇷 راهنمای فارسی](#-راهنمای-جامع-فارسی) | [🇬🇧 English Documentation](#-english-overview)

[![Tests](https://img.shields.io/badge/Tests-43%20Passed%20(100%25)-success?style=for-the-badge&logo=pytest&logoColor=white)](tests/test_full_skill_diagnostics.py)
[![Version](https://img.shields.io/badge/Version-3.9.9-teal?style=for-the-badge)](SKILL.md)
[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Deliverable](https://img.shields.io/badge/Deliverable-Quad--Pack%20(Word%20|%20PDF%20|%20HTML%20|%20MD%20|%20Anki)-purple?style=for-the-badge)](dist/)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![Platform](https://img.shields.io/badge/Google%20Antigravity-Compatible-blueviolet?style=for-the-badge)](https://antigravity.google)

A publication-grade, automated medical knowledge summarization and high-yield study guide generator tailored for medical students in **Physiopathology (فیزیوپاتولوژی)** and clinical rotations. It ingests complex multi-format handouts (`.docx`, `.pdf`, `.md`, `.txt`), enforces a strict **Zero-Omission & Anti-Bloat Verification Gate**, enforces **Pure Dynamic Table-Centric & Infographic Chart Architecture (v3.9.9 - Rule 24)** with dynamic 2-to-5 context-driven columns, zero-bloat single-tier sections, high-DPI modern Playwright Mermaid cards, and high-yield badges, enforces **Chronological Topic Pairing & Multi-Handout Table of Contents Gate (v3.9.6 - Rule 21)**, enforces **Local Context Binding for Contraindications (v3.9.5 - Rule 20)**, enforces **Clinical Cutoff Isolation & Zero-Omission Gate (v3.9.5 - Rule 20)**, enforces **Mandatory Instructor Emphasis Gate (v3.9.5 - Rule 20)**, provides **Abolition of Numeric Word Caps & Compression Ratios in favor of Zero-Omission First Aid Style (v3.9.8 - Rule 22)**, supports **Multi-Tier Scanned Page Ingestion with PyMuPDF OCR & Multimodal Vision Tokens (v3.9.5 - Rule 20)**, preserves exact **Chronological Sequence of Topics & Teaching Flow (v3.9.1)**, enforces universal **Strict RTL Directionality & BiDi Typography (v3.9.1)**, mandates **Subtopic Atomicity & Dedicated Tables (v3.9.2 - Rule 15)**, enforces **Telegraphic First Aid Syntax (Rule 17)**, guarantees **Cross-Layer Deduplication (Rule 18)**, scales document length dynamically with source volume via **Adaptive Content Scaling (Rule 19)**, fixes Python string iteration bugs via `ensure_list`, eliminates GenAnki invalid HTML warnings (**Rule 16**), guarantees Word table XML safety (`w:cantSplit`, `w:tblHeader`, `w:bidiVisual`), provides graceful **Windows Exclusive PDF Lock Fallback (`_Updated.pdf`)**, highlights critical points via **In-Cell Badges** (`🔴 تله امتحانی`, `💎 استاندارد طلایی`, `⚠️ خطای رایج`), renders high-DPI clinical diagrams adhering to the **Golden Aspect Ratio Legibility Gate**, isolates BiDi math formulas, synthesizes atomic Anki cards via the **Table-Row Cloze Engine**, and outputs the synchronized **Quad-Pack**:
1. **Microsoft Word Study Guide (`.docx`)**: Custom RTL typography, Dubai font, Teal Medical palette (`#007A87`), repeating table headers across page breaks (`tblHeader`), anti-row split enforcement (`cantSplit`), table RTL visual orientation (`<w:bidiVisual w:val="1"/>`), subtle zebra striping (`#F0FDFA`), true Word bold runs from markdown (`run.bold = True`), native formula cards, High-DPI transparent diagram PNG rendering (`Inches(5.8)`), dedicated Multimodal Approach Atlas section (`Inches(4.8)` centered), color-coded callouts (emergency, physiology, contraindications, special populations), comparative tables, and Master Pharmacology Matrix.
2. **Direct Publication-Grade PDF (`.pdf`)**: Native HTML-to-PDF rendering via Playwright using system Edge/Chrome, custom teal medical headers (`#007A87`) featuring institution branding (**دانشگاه بق**), complete browser artifact suppression, Windows PDF exclusive-lock fallback (`_Updated.pdf`), and Persian BiDi page stamping (`صفحه X از Y`).
3. **Interactive Standalone HTML (`.html`)**: Mobile-ready responsive study dashboard with Mermaid.js integration, strict RTL tables (`<table dir="rtl">`), native `.formula-card` and `.figure-card` CSS, KaTeX auto-rendering, A4 print CSS, in-cell badge styling, collapsible clinical vignettes (`<details>`), and interactive matrices.
4. **Formatted Markdown (`.md`)**: GitHub/Obsidian-compatible Mermaid flowcharts, sequential section order matching lecture slides, continuous blockquote alerts (`> ` prefix on all sub-lines), Multimodal Visual Atlas callout blocks, Track 4 Clinical Semiology tables, and escaped table formatting with `<br>` multi-line cells.
5. **Anki Spaced Repetition Decks (`.tsv` & native `.apkg` with Smart Dual Note Types)**: Supports simultaneous **PhysioPath Basic Model** (Question isolated on Front, Answer on Back with emerald badge) and **PhysioPath Cloze Model** (atomic deletions), Teal Medical styling (`#007A87` border, rounded container, Vazirmatn typography, nightMode dark mode), native TSV import directives (`#separator:tab`, `#html:true`, `#notetype:Cloze`), and companion `_basic.tsv` (`#notetype:Basic`).

---

## ⛔ The Core Law: Strict Source Grounding (Zero Extrapolation)

> ### ⚠️ قانون حیاتی و خط قرمز: وفاداری ۱۰۰٪ به منبع تدریس‌شده
> **همه محتوای خلاصه باید فقط و فقط طبق محتوای فایل‌ها و جزواتی باشد که کاربر ارائه می‌دهد و نه از دانش پیش‌فرض یا حافظه پارامتریک ایجنت.**
> هیچ ایجنتی تحت هیچ شرایطی حق ندارد از حافظه پارامتریک خود یا کتب مرجع خارجی (هاریسون، سیسیل، آپ‌تودیت) بیماری‌ها، سندرم‌ها، داروها یا استیجینگ‌هایی را که استاد در جزوه تدریس نکرده، وارد خلاصه کند!
> خلاصه نهایی دقیقاً و منحصراً بر پایه فایل‌های داده‌شده (جزوه یا اسلاید استاد) تدوین می‌گردد. در صورت عدم ذکر یک تست تشخیصی، کات‌آف یا رژیم درمانی در جزوه، آن فیلد صریحاً `null` مانده و مدل حق ابداع یا تعمیم دادن آن را ندارد.

---

## 📋 اصول ده‌گانه تلخیص، انسجام و انضباط ساختاری (High-Yield Editorial Directives)

1. **پرهیز مطلق از مقدمه‌چینی، توضیح درباره روند کار، نتیجه‌گیری، نظر شخصی، تکرار مطالب و حشو:** از تعارفات کلامی و توضیحات فرامتنی خودداری شود.
2. **استخراج انحصاری عصاره علمی و نکات آزمونی:** فقط نکات تست‌خیز و قابل آزمون استخراج گردند.
3. **وفاداری ۱۰۰٪ به منبع و منع مطلق افزودن اطلاعات:** اگر مطلبی در فایل نیست، به هیچ وجه به خلاصه اضافه نشود و فیلد مربوطه `null` بماند.
4. **حفظ منظم و فوق‌فشرده ساختارهای بالینی:** علل، عوامل خطر، پاتوفیزیولوژی، علائم، آزمایش‌ها، تشخیص، درمان، عوارض، پیش‌آگهی، استیجینگ و افتراق‌ها در فشرده‌ترین قالب حفظ شوند.
5. **حفظ روابط علت و معلولی و زنجیره پاتوفیزیولوژی:** توجیه مکانیسمی و منطق رخداد علائم حفظ شود تا خلاصه فقط فهرستی خشک از کلمات نباشد.
6. **پرهیز از بولت‌سازی افراطی و تجمیع مطالب مرتبط:** مطالبی که به یکدیگر مرتبط هستند در یک بولت یا پاراگراف منسجم تجمیع شوند.
7. **خروجی یکدست، منسجم، فوق‌ساختاریافته و جامع (قانون تلخیص بدون حذف و پوشش ۱۰۰٪ فکت‌ها):** هدف از خلاصه، سازمان‌دهی سیستماتیک داده‌ها در قالب جداول استاندارد است نه حذف کورکورانه. هیچ ضریب عددی فشردگی یا سقف کلمه‌ای تحمیل نمی‌شود؛ کلیه معیارهای تشخیصی، کات‌آف‌ها، تله‌های آزمونی و شرح‌های تکمیلی رفرنس و اسلایدهای کلاسی موجود در جزوه باید ۱۰۰٪ و با جزئیات کامل درون سلول‌های جدول با بولت‌های ساختاریافته (First Aid Zero-Omission Style) پوشش یابند.
8. **حفظ دقیق توالی زمانی و سرفصل‌های جزوه (Strict Chronological Sequence):** ترتیب تمام بخش‌ها، جداول و بیماری‌ها در تمامی خروجی‌ها (Word, PDF, HTML, Markdown) باید **دقیقاً و مو‌به‌مو منطبق بر روند تدریس استاد و توالی اسلایدهای جزوه اصلی** باشد؛ هرگونه جابه‌جایی یا پس‌وپیش کردن مباحث اکیداً ممنوع است.
9. **جهت‌داری راست‌به‌چپ مطلق و انضباط BiDi (Strict RTL Typography):** تمامی صفحات، جداول (`<table dir="rtl">`، `<w:bidiVisual w:val="1"/>`)، پاراگراف‌ها (`<w:bidi w:val="1"/>`) و سلول‌ها صددرصد راست‌به‌چپ بوده و عبارات انگلیسی و فرمول‌ها با ایزولاسیون کامل LTR رندر می‌شوند تا از پرش متنی جلوگیری گردد.
10. **تفکیک ساب‌تاپیک‌ها و منع ادغام موضوعات بالینی (Subtopic Atomicity & Dedicated Tables):** هر عنوان ماژور بالینی و هر ساب‌تاپیک مستقل، باید جدول اختصاصی خود را داشته باشد و نباید در یک سطر ناخوانای جدول مجاور ادغام یا فشرده گردد.
11. **هم‌مکانی موضوعی و فهرست مطالب چندجزوه (Chronological Topic Pairing & Multi-Handout TOC):** هم‌مکانی و قرارگیری بلافاصله‌ی جداول، فلوچارت‌ها و مرواریدهای هر مبحث در کنار هم در توالی دقیق جزوه اصلی، فهرست جامع مطالب در ابتدای اسناد چندجزوه با جهت راست‌به‌چپ، بنرهای تفکیک سرفصل‌ها و برچسب‌های شفاف انتساب (`[جزوه ۱]`, `[جزوه ۲]`).
12. **معماری جدول‌محور فراگیر بر پایه الگوی رفرنس (Exemplar Full Table-Centric Architecture - Rule 22):** کل سند بر مبنای بخش‌های شماره‌دار جدول‌محور (`بخش ۱: ...`, `بخش ۲: ...`) با کهن‌الگوهای ۳ و ۴ ستونه استاندارد مشابه فایل مرجع ۱.pdf بدون پاراگراف‌های پراکنده بیرونی تدوین می‌گردد و در انتهای سند با ماتریس داروشناسی و چک‌لیست صبح آزمون تکمیل می‌شود.

---

## 🌟 Key Innovations & Architectural Pillars (v3.9.2)

```
┌────────────────────────────────────────────────────────────────────────┐
│             MEDICAL SUMMARY EXPERT PIPELINE (v3.9.2)                   │
├────────────────────────────────────────────────────────────────────────┤
│ 1. INGESTION & ZERO-OMISSION PARSING                                   │
│    - OOXML Deep Traversal: Word Textboxes, Table-Cell Shapes & SDT     │
│    - PyMuPDF Vector Extraction + 150 DPI Scanned Page OCR Fallback     │
│    - 100% Mandatory Slide & Clinical Photo Inspection                  │
├────────────────────────────────────────────────────────────────────────┤
│ 2. TABLE-FIRST HYBRID ARCHITECTURE & CONTEXT-AWARE SCHEMAS             │
│    - 6 Thematic Table Archetypes: Pathophysiology, Anatomy, Pharma,    │
│      Diagnostic Cutoffs, Differential Matrix & Dynamic Tests           │
│    - In-Cell Badges: 🔴 Trap, ⚠️ Error, 💎 Gold Standard, 📌 Key Note │
│    - 3-4 Page Compact Golden Review Budget per Lecture Session         │
├────────────────────────────────────────────────────────────────────────┤
│ 3. SEQUENTIAL ORDER & UNIVERSAL RTL BIDI TYPOGRAPHY                    │
│    - Strict Chronological Teaching Flow: Shard index order preserved   │
│    - Universal RTL: dir="rtl", w:bidiVisual="1", set_p_rtl everywhere  │
│    - Zero BiDi Flip: Latin acronyms, drug dosages & units LTR-isolated │
├────────────────────────────────────────────────────────────────────────┤
│ 4. RIGOROUS QUALITY GATES & VERIFICATION AUDIT                         │
│    - Flowchart Legibility Gate: Vertical TD, Aspect Ratio (0.7-1.6:1)  │
│    - Multimodal Policy: Strict Omission of Patient Photos from Canvas  │
│    - LaTeX/BiDi Gate: Zero raw TeX leaks, native math formula cards    │
│    - Word Bold Run Engine: True run.bold without asterisk leaks        │
├────────────────────────────────────────────────────────────────────────┤
│ 5. PUBLICATION-GRADE DELIVERABLES COMPILER                             │
│    - Teal Medical Palette (#007A87 header, #F0FDFA zebra, clean borders)│
│    - DOCX: Word RTL, Dubai Font, tblHeader, cantSplit, True Bold Runs  │
│    - Playwright PDF: Direct HTML-to-PDF + 'دانشگاه بق' header + BiDi   │
│    - Anki: Table-Row Cloze Engine, deterministic hash GUIDs, .apkg/.tsv│
└────────────────────────────────────────────────────────────────────────┘
```

### 1. 🧬 Polymorphic Module Architecture & Balanced Handout Coverage (v3.6.1)
* **Foundational Physiology (`foundational_concept`):** Normal homeostasis, organ axes (HPA, HPT, HPG), baseline functional tests, and regulatory feedback loops with Mermaid diagrams.
* **Clinical Diagnostic Approach (`clinical_approach`):** Step-by-step triage algorithms, decision bifurcations / pivot criteria, and definitive next steps.
* **Disease Illness Script (`illness_script`):** Full disease coverage with surgical/interventional indications, lifestyle/conservative management, complications/prognosis, and high-yield pearls.
* **Polymorphic Crash Sheets & Morning Checklist:** Auto-synthesizes non-empty crash sheet rows for non-disease modules and prioritizes emergency resuscitation, instructor alerts, contraindications, gold standards, and pivot points.

### 2. 📋 Staging Systems, Emergency Management & Special Populations
* **Staging & Diagnostic Criteria (`staging_and_classification`):** Native multi-format rendering of Duke, Wells, NYHA, Killip, KDIGO, TNM criteria as tables or structured lists.
* **Emergency Stabilization (`emergency_management`):** High-priority emergency resuscitation callouts for acute unstable patient care.
* **Special Populations (`special_populations_and_pregnancy`):** Dedicated callouts for trimester-specific drug selections (e.g. PTU vs Methimazole), pediatrics, and organ failure dose titrations.

### 3. 🔍 Universal Dynamic Clinical Syndrome & Disease Audit
* **Cross-Specialty NER Engine:** Dynamically extracts clinical syndromes, diseases, triads, and criteria across Cardiology, Nephrology, Pulmonology, Neurology, Rheumatology, Hematology, Pediatrics, and Basic Sciences.
* **Zero Omission Detection:** Eliminates dependence on a hardcoded endocrine list, ensuring syndromes like Eisenmenger syndrome, Kawasaki disease, Beck triad, or Duke criteria are never omitted.

### 4. 🖨️ Direct Playwright HTML-to-PDF & Persian BiDi Page Stamping
* **Playwright System Browser Rendering:** Renders publication-grade PDFs via Playwright using pre-installed Microsoft Edge or Google Chrome without downloading heavyweight browser binaries.
* **Header & Artifact Suppression:** Custom burgundy header (`border-bottom: 1.5px solid #78281F`) featuring chapter title, course name, and **«دانشگاه بق»** institution branding. Browser default URLs, dates, and timestamps are suppressed.
* **Persian BiDi Stamping (`صفحه X از Y`):** Utilizes `pypdf` + `reportlab` + `arabic-reshaper` + `python-bidi` with Windows Tahoma font to stamp bidirectional Persian page numbers at the footer without character reversal or font disconnection.

### 5. 📑 Subtopic Atomicity & Dedicated Tables (v3.9.2 - Rule 15)
* **Dedicated Thematic Tables (`SUBTOPIC_ATOMICITY_GATE`):** Prohibits compressing independent major clinical subtopics (e.g. CGMS, Diabetic Nephropathy & Kimmelstiel-Wilson, Retinopathy NPDR/PDR, Diabetic Foot & Charcot joint, Somogyi vs Dawn, Klinefelter, Kallmann, Gynecomastia & TRT) into single rows of general tables.
* **Rich Multi-Line Pathophysiology:** High-yield compression retains rich multi-line bullets (`•`) detailing causal mechanisms, numeric thresholds, and clinical pearls rather than dry single-word keywords.
* **Concept Depth Audit (`audit_concept_depth`):** Automatically flags summaries that aggressively truncate major topics (`< 0.08` ratio on >=500 words) or provide fewer than 3 analytical rows/bullets for core entities (`UNDER_EXPLAINED`).

### 6. 🛡️ Universal Table XML Safety, GenAnki Sanitization & PDF Lock Resilience (Rule 16)
* **Universal Word Table XML Tags:** Enforces `w:bidiVisual` on table properties, `w:cantSplit` across every row of every table, and `w:tblHeader` on header row 0 of all multi-row tables.
* **GenAnki HTML Sanitization:** Converts standalone `<` and `>` into `&lt;` and `&gt;` while preserving valid HTML formatting (`<b>`, `<br>`, `<span>`) and cloze deletions (`{{c1::...}}`), eliminating GenAnki `UserWarning: Field contained invalid HTML tags`.
* **Windows PDF Lock Fallback:** Gracefully handles `PermissionError [Errno 13]` when a target PDF is open in Adobe Acrobat or Edge browser by saving to `_Updated.pdf`.

### 7. 🎯 Telegraphic First Aid Syntax, Cross-Layer Deduplication & Adaptive Content Scaling (v3.9.4 - Rules 17-19)
* **Telegraphic First Aid Syntax (`TELEGRAPHIC_SYNTAX_AND_CELL_COMPRESSION_GATE` - Rule 17):** Enforces structured telegraphic entries (`• **کلیدواژه:** پیامد بالینی با نماد -> و کات‌آف عددی`) and caps table cell volume at **25-30 words**. Explicitly prohibits subject-verb-object narrative prose paragraphs (`فاعل + مفعول + فعل ربطی`) inside table cells.
* **Cross-Layer Deduplication Gate (`CROSS_LAYER_DEDUPLICATION_GATE` - Rule 18):** Resolves topic triplication between thematic tables (`sections`), module illness scripts (`modules`), and crash sheets (`master_matrices`). When an entity is presented in a thematic table, repetitive text paragraphs in `modules` are suppressed and replaced by a concise cross-reference banner, while strictly retaining high-value callouts (emergency management, instructor alerts, contraindications, special populations, semiology, pearls, and staging).
* **Adaptive Content Scaling (v3.9.4):** Dynamically scales summary length, tables, and page count with the volume and complexity of the lecture handout, without imposing an artificial fixed page cap or word limit so that extensive multi-slide handouts are preserved completely.
* **Process Flowcharts vs. Static Tables (`PROCESS_FLOWCHARTS_VS_STATIC_TABLES_GATE` - Rule 19):** Routes dynamic regulatory feedback loops, hormonal axes, and triage algorithms into vertical Mermaid flowcharts (`flowchart TD`), reserving 3-column tables exclusively for differential diagnosis matrices, numerical cutoffs, and pharmacology.
* **Universal String Iteration Immunity & Color Safety:** Implements `ensure_list(val)` across Word, Markdown, HTML, and Anki compiler routines to prevent string character-splitting bugs (`س`, `و`, `ن`, `و`...), and defines missing `COLOR_TEAL_BORDER` / `COLOR_TEAL_BG` palettes.

---

## 📁 Repository Structure

```text
├── run_summary_pipeline.py     # Main CLI orchestrator (ingest, partition, merge, audit, build, bundle, auto)
├── parse_input_lecture.py      # DOCX & PDF ingestion engine (tables, shapes, SDT, images, OCR)
├── build_summary_pack.py       # Quad-pack compiler (DOCX, Playwright PDF, HTML, MD, Anki TSV/APKG)
├── pdf_designer_engine.py      # Standalone dual-pass PDF engine with vector borders and A4 geometry
├── audit_summary_coverage.py   # Clinical NLP audit gate (cell bloat, deduplication, adaptive volume, concept depth)
├── SKILL.md                    # Antigravity skill specification and guidelines (v3.9.4)
├── sample_summary.json         # Standard structured JSON data
├── tests/                      # Comprehensive automated stress-test suite
│   ├── test_full_skill_diagnostics.py   # 38-test diagnostic suite (Bug A to Bug AL)
│   ├── test_omission_bugs.py            # Word table & textbox preservation test
│   ├── test_pdf_ingestion.py            # PyMuPDF table and scanned page test
│   └── test_digit_and_guid_bugs.py      # Persian digit normalization & Anki GUID test
└── dist/                       # Output artifacts directory
    ├── summary_pack.docx       # Publication-ready Persian Word study guide
    ├── summary_pack.pdf        # Playwright rendered PDF with Persian BiDi page stamping
    ├── summary_pack.html       # Interactive standalone HTML dashboard
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

# 4. Run Clinical Zero-Omission, Dynamic Syndrome & Anti-Bloat Compression Audit Gate
uv run python run_summary_pipeline.py audit --source "ingested_lecture.json" --summary "summary_data.json" --report "coverage_audit_report.json"

# 5. Build Deliverables (Word + Playwright PDF + HTML + Markdown + Anki TSV/APKG)
uv run python run_summary_pipeline.py build --summary "summary_data.json" --institution "دانشگاه بق" --out-dir "dist" --prefix "summary_pack"

# 6. Bundle Semester Course into Single Master Booklet (Cover + TOC + Chapters + Continuous BiDi Stamping)
uv run python run_summary_pipeline.py bundle --inputs "sessions/*.json" --course-title "فیزیوپاتولوژی جامع" --institution "دانشگاه بق" --out-dir "dist/master_bundle" --prefix "comprehensive_master_booklet"

# Auto-Pipeline (All-in-one for single handouts)
uv run python run_summary_pipeline.py auto --input "lecture.docx" --summary-json "summary_data.json" --institution "دانشگاه بق"
```

---

## 🧪 Automated Test Suite (38/38 Passing - 100% Green)

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
| **Bug T** | Playwright HTML-to-PDF rendering with Persian BiDi page stamping (`صفحه X از Y`) | **PASS** ✅ |
| **Bug U** | Anti-Pharmacology-Skew Law & Illness Script Integrity Gate enforcing subtle syndrome coverage | **PASS** ✅ |
| **Bug V** | Track 4 Clinical Semiology & Physical Examination extraction and Anki card synthesis | **PASS** ✅ |
| **Bug W** | 30-Point Morning-of-Exam Golden Capsule and Comparative Matrices across all deliverables | **PASS** ✅ |
| **Bug X** | Standalone `MedicalPDFDesigner` dual-pass hybrid rendering and A4 print geometry | **PASS** ✅ |
| **Bug Y** | Multi-session Comprehensive Course Master Booklet (Cover + TOC + Chapters + 'دانشگاه بق') | **PASS** ✅ |
| **Bug Z** | Polymorphic Module Architecture (Foundational, Algorithm, Illness Script) and Anki Parity | **PASS** ✅ |
| **Bug AA** | High-Yield Compression Gate & Verbatim Bloat Prevention Audit | **PASS** ✅ |
| **Bug AB** | Polymorphic Module Parity in Crash Sheet & Multi-Priority Morning Checklist Auto-Synthesis | **PASS** ✅ |
| **Bug AC** | Staging Systems, Scoring Criteria & Diagnostic Criteria Multi-Format Support | **PASS** ✅ |
| **Bug AD** | Immediate Emergency Stabilization & Special Populations (Pregnancy/Pediatrics) Support | **PASS** ✅ |
| **Bug AE** | Universal Dynamic Syndrome and Disease Detection Across All Medical Specialties | **PASS** ✅ |
| **Bug AF** | Capstone Compact, High-Yield, Comprehensive Quad-Pack Parity Verification | **PASS** ✅ |
| **Bug AG** | Visual Diagram Rendering (High-DPI Mermaid to PNG) & Multimodal Visual Atlas Pipeline | **PASS** ✅ |
| **Bug AH** | LaTeX/BiDi Sanitization, True Word Bold Runs, Flowchart Legibility & Alert Continuity | **PASS** ✅ |
| **Bug AI** | Table-First Hybrid Teal Architecture, In-Cell Badges & Anki Table-Row Cloze Synthesis | **PASS** ✅ |
| **Bug AJ** | Chronological Teaching Sequence Flow & Universal Strict RTL BiDi Typography | **PASS** ✅ |
| **Bug AK** | GenAnki HTML Escaping, Universal Word cantSplit/tblHeader & PDF Lock Resilience | **PASS** ✅ |
| **Bug AL** | Telegraphic First Aid Syntax, Cross-Layer Deduplication & Adaptive Content Scaling | **PASS** ✅ |

To run the full diagnostic suite:
```bash
uv run python tests/test_full_skill_diagnostics.py
```

---

## 🇮🇷 راهنمای جامع فارسی

### هدف اسکیل (نسخه ۳.۹.۴)
این اسکیل تخصصی برای تبدیل جزوات اسلایدی و متنی اساتید، ترنسکریپت‌های کلاسی و هندآوت‌های درسی دانشگاه‌های علوم پزشکی به **بسته چندگانه مرور های‌یلد (Quad-Pack)** طراحی شده است:

1. **جزوه رسمی Word (`.docx`):** با استایل‌بندی مدرن، فونت خوانای دبی، دایرکشن راست‌به‌چپ (RTL)، کادرهای رنگی اختصاصی (زنجیره پاتوفیزیولوژی، سمیولوژی، اقدامات اورژانس، ملاحظات بارداری، تأکیدات مستقیم استاد، موارد مصرف ممنوع، سناریوهای خودآزمایی بالینی، الگوریتم‌های تریاژ)، ماتریس‌های تطبیقی، جدول مرحله‌بندی و معیارهای بالینی، ماتریس مقایسه جامع داروها و جدول فشرده شب امتحان.
2. **سند پرینت مستقیم PDF (`.pdf` با Playwright):** تولید فایل PDF با کیفیت چاپی عالی از طریق Playwright با مرورگرهای بومی ویندوز، هدر عنابی رنگ اختصاصی (`#78281F`)، بدون آرتفکت‌های مرورگر، درج «دانشگاه بق» در هدر و مجهز به شماره‌گذاری فارسی `صفحه X از Y`.
3. **کتابچه جامع ترم (`comprehensive_master_booklet.pdf`):** صحافی و تجمیع تمام فصول و جلسات یک درس در قالب یک پی‌دی‌اف مرجع واحد با جلد نفیس رنگی، فهرست مطالب خودکار و شماره‌گذاری پیوسته.
4. **داشبورد تعاملی HTML (`.html`):** نمایش تمیز تحت وب با استایل‌های اختصاصی `@page`، آکاردئون‌های تاشوی سناریوهای بالینی، کادرهای اختصاصی اورژانس و بارداری و پشتیبانی کامل BiDi.
5. **سند مارک‌داون (`.md`):** شامل فلوچارت‌های کدنویسی‌شده استاندارد Mermaid، بخش سمیولوژی، معیارهای تشخیصی و مرحله‌بندی، خودآزمایی بالینی تاشو با پاسخ تشریحی مکانیسمی (`<details>`)، و پاک‌سازی کامل جدول‌ها با اسکپ کاراکترهای پایپ.
6. **فلش‌کارت‌های استاندارد آنکی (`.tsv` و `.apkg`):** کارت‌های جای‌خالی اتمیک بر پایه ۲۰ قانون Wozniak، بدون ریسک تکراری شدن یا ریست زمان‌بندی مرور دانشجو به دلیل بهره‌گیری از شناسه‌های پایدار قطعی (Deterministic GUIDs)، به همراه سنتز خودکار کارت‌های سمیولوژی، اورژانس، استیجینگ، بارداری، اندیکاسیون‌های جراحی، عوارض و الگوریتم‌ها.

---

## 📄 لایسنس (License)

این پروژه تحت مجوز [MIT License](LICENSE) منتشر شده است. استفاده و توسعه آن برای دانشجویان و اساتید علوم پزشکی آزاد است.
