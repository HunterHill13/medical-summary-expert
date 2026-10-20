# 🩺 Medical Summary Expert (v3.5)
## Publication-Grade Medical Lecture Notes & Study Guide Summarization Engine

[🇮🇷 راهنمای فارسی](#-راهنمای-جامع-فارسی) | [🇬🇧 English Documentation](#-english-overview)

[![Tests](https://img.shields.io/badge/Tests-24%20Passed%20(100%25)-success?style=for-the-badge&logo=pytest&logoColor=white)](tests/test_full_skill_diagnostics.py)
[![Version](https://img.shields.io/badge/Version-3.5.0-orange?style=for-the-badge)](SKILL.md)
[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Deliverable](https://img.shields.io/badge/Deliverable-Quad--Pack%20(Word%20|%20PDF%20|%20HTML%20|%20MD%20|%20Anki)-purple?style=for-the-badge)](dist/)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![Platform](https://img.shields.io/badge/Google%20Antigravity-Compatible-blueviolet?style=for-the-badge)](https://antigravity.google)

A publication-grade, automated medical knowledge summarization and high-yield study guide generator tailored for medical students in **Physiopathology (فیزیوپاتولوژی)** and clinical rotations. It ingests complex multi-format handouts (`.docx`, `.pdf`, `.md`, `.txt`), enforces a strict **Zero-Omission Verification Gate**, structures medical topics using canonical **Illness Scripts**, and outputs the synchronized **Quad-Pack**:
1. **Microsoft Word Study Guide (`.docx`)**: Custom RTL typography, Dubai font, color-coded callouts, comparative tables, and Master Pharmacology Matrix.
2. **Direct Publication-Grade PDF (`.pdf`)**: Native HTML-to-PDF rendering via Playwright using system Edge/Chrome, custom burgundy headers (`#78281F`), complete browser artifact suppression, and Persian BiDi page stamping (`صفحه X از Y`).
3. **Interactive Standalone HTML (`.html`)**: Mobile-ready responsive study dashboard with A4 print CSS, collapsible clinical vignettes (`<details>`), and interactive matrices.
4. **Formatted Markdown (`.md`)**: GitHub/Obsidian-compatible Mermaid flowcharts, Track 4 Clinical Semiology tables, and escaped table formatting.
5. **Anki Spaced Repetition Decks (`.tsv` & native `.apkg`)**: Wozniak's 20 Rules compliant, atomic cloze cards (`model_type=CLOZE`), deterministic prompt-hash GUIDs, and auto-synthesized semiology cards.

---

## ⛔ The Core Law: Strict Source Grounding (Zero Extrapolation)

> ### ⚠️ قانون حیاتی و خط قرمز: وفاداری ۱۰۰٪ به منبع تدریس‌شده
> **همه محتوای خلاصه باید فقط و فقط طبق محتوای فایل‌ها و جزواتی باشد که کاربر ارائه می‌دهد و نه از دانش پیش‌فرض یا حافظه پارامتریک ایجنت.**
> هیچ ایجنتی تحت هیچ شرایطی حق ندارد از حافظه پارامتریک خود یا کتب مرجع خارجی (هاریسون، سیسیل، آپ‌تودیت) بیماری‌ها، سندرم‌ها، داروها یا استیجینگ‌هایی را که استاد در جزوه تدریس نکرده، وارد خلاصه کند!
> خلاصه نهایی دقیقاً و منحصراً بر پایه فایل‌های داده‌شده (جزوه یا اسلاید استاد) تدوین می‌گردد. در صورت عدم ذکر یک تست تشخیصی، کات‌آف یا رژیم درمانی در جزوه، آن فیلد صریحاً `null` مانده و مدل حق ابداع یا تعمیم دادن آن را ندارد.

---

## 📋 اصول هفت‌گانه تلخیص، انسجام و پرهیز از حشو (High-Yield Editorial Directives)

1. **پرهیز مطلق از مقدمه‌چینی، توضیح درباره روند کار، نتیجه‌گیری، نظر شخصی، تکرار مطالب و حشو:** از تعارفات کلامی و توضیحات فرامتنی خودداری شود.
2. **استخراج انحصاری عصاره علمی و نکات آزمونی:** فقط نکات تست‌خیز و قابل آزمون استخراج گردند.
3. **وفاداری ۱۰۰٪ به منبع و منع مطلق افزودن اطلاعات:** اگر مطلبی در فایل نیست، به هیچ وجه به خلاصه اضافه نشود و فیلد مربوطه `null` بماند.
4. **حفظ منظم و فوق‌فشرده ساختارهای بالینی:** علل، عوامل خطر، پاتوفیزیولوژی، علائم، آزمایش‌ها، تشخیص، درمان، عوارض و افتراق‌ها در فشرده‌ترین قالب حفظ شوند.
5. **حفظ روابط علت و معلولی و زنجیره پاتوفیزیولوژی:** توجیه مکانیسمی و منطق رخداد علائم حفظ شود تا خلاصه فقط فهرستی خشک از کلمات نباشد.
6. **پرهیز از بولت‌سازی افراطی و تجمیع مطالب مرتبط:** مطالبی که به یکدیگر مرتبط هستند در یک بولت یا پاراگراف منسجم تجمیع شوند.
7. **خروجی یکدست، منسجم، کم‌حجم و پُرمحتوا:** خروجی نهایی متناسب با مرور سریع برای امتحان پزشکی کالیبره شود.

---

## 🌟 Key Innovations & Architectural Pillars (v3.5)

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
│ 3. ZERO-OMISSION & ANTI-PHARMACOLOGY-SKEW AUDIT GATE                   │
│    - Illness Script Integrity Gate (Etiology, Signs, Dx, Pearls)       │
│    - Subtle Syndrome Protection (Sheehan, Nelson, Carney, FHH, CAIS...)│
│    - Persian/Arabic Digit Canonicalization (۴۰٪ -> 40%, ۱،۵ -> 1.5)    │
│    - Thousands Separator Normalization (۵۰،۰۰۰ / 1,000,000 -> 50000)   │
│    - Substantive Warning Audit & USMLE Pharmacology Suffix Detection   │
├────────────────────────────────────────────────────────────────────────┤
│ 4. PUBLICATION-GRADE DELIVERABLES COMPILER                             │
│    - Direct Playwright PDF: Edge/Chrome headless + #78281F header      │
│    - Persian BiDi Stamping: 'صفحه X از Y' via ReportLab + pypdf        │
│    - DOCX: Word RTL, Dubai Font, 6 Custom Callout Styles, RTL Tables   │
│    - Interactive HTML: Print-ready CSS, Collapsible Vignettes          │
│    - MD: Sanitized Mermaid Flowcharts, Semiology & Matrix Tables       │
│    - Anki: CLOZE Model, Deterministic Hash GUIDs, Semiology Decks      │
└────────────────────────────────────────────────────────────────────────┘
```

### 1. 🖨️ Direct Playwright HTML-to-PDF & Persian BiDi Page Stamping (v3.5)
* **Playwright System Browser Rendering:** Renders publication-grade PDFs via Playwright using pre-installed Microsoft Edge or Google Chrome without downloading heavyweight browser binaries.
* **Header & Artifact Suppression:** Custom burgundy header (`border-bottom: 1.5px solid #78281F`) featuring chapter title and course name with RTL orientation. Browser default URLs, dates, and timestamps are suppressed.
* **Persian BiDi Stamping (`صفحه X از Y`):** Utilizes `pypdf` + `reportlab` + `arabic-reshaper` + `python-bidi` with Windows Tahoma font to stamp bidirectional Persian page numbers at the footer without character reversal or font disconnection.

### 2. ⚖️ Anti-Pharmacology-Skew Law (Illness Script Integrity Gate - v3.5)
* **Illness Script Integrity:** Enforces the 4 core pillars for every condition (Etiology & Cascade, Clinical Presentation & Triads, Diagnostics Bifurcation, High-Yield Pearls & Contraindications).
* **Subtle Syndrome Protection:** Prevents clinical syndromes (e.g. Sheehan syndrome, Nelson syndrome, Carney complex, FHH, Looser zones, Asherman, CAIS, 5-alpha reductase, pituitary apoplexy) from being eclipsed or dropped when extensive drug tables are generated.
* **Pharmacology Ratio Audit:** Audits text word distribution to ensure drug tables do not swallow clinical pathophysiology.

### 3. 🩺 Track 4: Clinical Semiology & Physical Examination (v3.5)
* **Dedicated Semiology Track:** Automatically extracts inspection and palpation maneuvers (water swallowing maneuver, Berry ligament, Crile technique, thyroid bruit), ophthalmic signs (Dalrymple, von Graefe, Stellwag, Hertel, NO SPECS), metabolic signs (Cushing violaceous striae, proximal myopathy chair test, Addison pigmentation, Chvostek and Trousseau signs), and neuropathy exams (128Hz tuning fork, 10g monofilament LOPS, Charcot foot, ABI).
* **Auto-Synthesized Semiology Decks:** Automatically generates spaced-repetition Anki cards tagged with `PhysioPath::[Topic]::Semiology`.

### 4. 🚀 30-Point Morning-of-Exam Golden Capsule & Comparative Matrices (v3.5)
* **Comparative Matrices (`comparative_matrices`):** Head-to-head comparison tables (e.g. comparative thyroiditis matrix, primary amenorrhea matrix, endocrine emergency triage order).
* **30-Point Morning Checklist (`morning_exam_checklist`):** 30 one-line high-yield pearls synthesized for rapid 10-minute review before exams.

---

## 📁 Repository Structure

```text
├── run_summary_pipeline.py     # Main CLI orchestrator (ingest, partition, merge, audit, build, auto)
├── parse_input_lecture.py      # DOCX & PDF ingestion engine (tables, shapes, SDT, images, OCR)
├── build_summary_pack.py       # Quad-pack compiler (DOCX, Playwright PDF, HTML, MD, Anki TSV/APKG)
├── pdf_designer_engine.py      # Standalone dual-pass PDF engine with vector borders and A4 geometry
├── audit_summary_coverage.py   # Clinical NLP audit gate (numbers, cutoffs, warnings, illness scripts)
├── SKILL.md                    # Antigravity skill specification and guidelines
├── sample_summary.json         # Standard structured JSON data (Heart Failure HFrEF vs HFpEF)
├── tests/                      # Comprehensive automated stress-test suite
│   ├── test_full_skill_diagnostics.py   # 23-test suite (Bug A to Bug W)
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

# 4. Run Clinical Zero-Omission & Illness Script Integrity Audit Gate
uv run python run_summary_pipeline.py audit --source "ingested_lecture.json" --summary "summary_data.json" --report "coverage_audit_report.json"

# 5. Build Deliverables (Word + Playwright PDF + HTML + Markdown + Anki TSV/APKG)
uv run python run_summary_pipeline.py build --summary "summary_data.json" --out-dir "dist" --prefix "summary_pack"

# Auto-Pipeline (All-in-one for single handouts)
uv run python run_summary_pipeline.py auto --input "lecture.docx" --summary-json "summary_data.json"
```

---

## 🧪 Automated Test Suite (23/23 Passing)

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

To run the full diagnostic suite:
```bash
uv run python tests/test_full_skill_diagnostics.py
```

---

## 🇮🇷 راهنمای جامع فارسی

### هدف اسکیل
این اسکیل تخصصی برای تبدیل جزوات اسلایدی و متنی اساتید، ترنسکریپت‌های کلاسی و هندآوت‌های درسی دانشگاه‌های علوم پزشکی به **بسته چندگانه مرور های‌یلد (Quad-Pack)** طراحی شده است:

1. **جزوه رسمی Word (`.docx`):** با استایل‌بندی مدرن، فونت خوانای دبی، دایرکشن راست‌به‌چپ (RTL)، کادرهای رنگی اختصاصی (زنجیره پاتوفیزیولوژی، سمیولوژی، تأکیدات مستقیم استاد، موارد مصرف ممنوع، سناریوهای خودآزمایی بالینی، نکات دامی آزمون)، ماتریس‌های تطبیقی، ماتریس مقایسه جامع داروها و جدول فشرده شب امتحان.
2. **سند پرینت مستقیم PDF (`.pdf` با Playwright):** تولید فایل PDF با کیفیت چاپی عالی از طریق Playwright با مرورگرهای بومی ویندوز، هدر عنابی رنگ اختصاصی (`#78281F`)، بدون آرتفکت‌های مرورگر و مجهز به شماره‌گذاری فارسی `صفحه X از Y`.
3. **داشبورد تعاملی HTML (`.html`):** نمایش تمیز تحت وب با استایل‌های اختصاصی `@page`، آکاردئون‌های تاشوی سناریوهای بالینی و پشتیبانی کامل BiDi.
4. **سند مارک‌داون (`.md`):** شامل فلوچارت‌های کدنویسی‌شده استاندارد Mermaid، بخش سمیولوژی، خودآزمایی بالینی تاشو با پاسخ تشریحی مکانیسمی (`<details>`)، و پاک‌سازی کامل جدول‌ها با اسکپ کاراکترهای پایپ.
5. **فلش‌کارت‌های استاندارد آنکی (`.tsv` و `.apkg`):** کارت‌های جای‌خالی اتمیک بر پایه ۲۰ قانون Wozniak، بدون ریسک تکراری شدن یا ریست زمان‌بندی مرور دانشجو به دلیل بهره‌گیری از شناسه‌های پایدار قطعی (Deterministic GUIDs)، به همراه سنتز خودکار کارت‌های سمیولوژی و معاینه بالینی.

---

## 📄 لایسنس (License)

این پروژه تحت مجوز [MIT License](LICENSE) منتشر شده است. استفاده و توسعه آن برای دانشجویان و اساتید علوم پزشکی آزاد است.

---

## 📄 لایسنس (License)

این پروژه تحت مجوز [MIT License](LICENSE) منتشر شده است. استفاده و توسعه آن برای دانشجویان و اساتید علوم پزشکی آزاد است.
