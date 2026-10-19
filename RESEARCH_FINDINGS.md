# Comprehensive Research Report: Medical Lecture Summarization, High-Yield Study Guide Generation, Active Recall & Anki Workflows

---

# Open-Source Repositories, Pedagogical Frameworks & Architecture for Medical Lecture Summarization and High-Yield Revision Note Generation

## Executive Summary
This document provides an exhaustive research foundation for designing a state-of-the-art **Medical Lecture Summarization & High-Yield Study Guide Generator** skill (`اسکیل خلاصه نویسی پزشکی`). It reviews existing open-source GitHub repositories, benchmarked medical prompt frameworks, evidence-based cognitive architectures (Illness Scripts, Cornell note-taking, Wozniak's 20 rules of knowledge formulation, cloze deletion engineering), and comparative differential diagnosis matrix designs.

---

## Part 1: GitHub Repositories & Open-Source Tools Review

### 1.1 Flashcard & Spaced Repetition (Anki) Generators
* **[AnKing-VIP/AnKing-Note-Types](https://github.com/AnKing-VIP/AnKing-Note-Types)** & **[AnKing-VIP/anking_notes_addon](https://github.com/AnKing-VIP/anking_notes_addon)**
  * **Core Architecture:** The definitive industry gold-standard schema for medical flashcards used by over 100,000+ USMLE/medical students globally.
  * **Standard Note Fields:** `Text` (the prompt/cloze), `Extra` (overarching mechanism, Pathoma/First Aid snippets), `Lecture Notes` (school-specific high-yield notes protected against remote overwrites), `Missed Questions`, `First Aid`, `Pathoma`, `Sketchy`, `Physeo`, `HighYield` tags.
  * **Key Takeaway for New Skill:** Any Anki-generation engine must align its export fields with this schema (or export clean Anki-compatible CSV/Markdown) so medical students can seamlessly import or integrate into AnKing.
* **[abusuraihsakhri/flashcard-generator](https://github.com/abusuraihsakhri/flashcard-generator)**
  * **Features:** Full-stack React + FastAPI tool built specifically for medical students. Ingests raw medical prose or **PubMed PMIDs**, leverages Claude AI to extract high-yield clinical entities, and formats them into `.apkg` and CSV files with cloze deletions.
* **[thiswillbeyourgithub/AnkiAIUtils](https://github.com/thiswillbeyourgithub/AnkiAIUtils)**
  * **Features:** AI utility tailored for medical curricula. Ingests challenging cards and uses LLMs to enrich explanations, create memory hooks, and embed mnemonics while strictly preserving existing cloze tags (`{{c1::...}}`) and image links.
* **[raine/anki-llm](https://github.com/raine/anki-llm)**
  * **Features:** CLI/TUI toolkit for bulk-generating and refining flashcards using multiple LLM providers (Gemini, Claude, GPT, Ollama). Uses `AnkiConnect` API for direct in-place card updates, avoiding tedious manual `.apkg` imports.
* **[terkelg/anki-markdown](https://github.com/terkelg/anki-markdown)**
  * **Features:** Modern Markdown-to-Anki card renderer with Shiki syntax highlighting and built-in AI agent skills that allow LLMs to directly interact with Anki decks via `AnkiConnect`.

### 1.2 Clinical Document & Lecture Summarizers
* **[Neilus03/recsum](https://github.com/Neilus03/recsum)**
  * **Features:** High-recall medical document summarizer utilizing the GoLLIE (Guideline-following Information Extraction) framework and Llama-3. Prevents omission of crucial clinical facts and quantitative lab values.
* **[the-momentum/notetaker](https://github.com/the-momentum/notetaker)**
  * **Features:** Audio-to-clinical note pipeline converting long-form medical dialogue into structured **SOAP notes** (Subjective, Objective, Assessment, Plan) with high clinical accuracy.
* **[microsoft/clinical_visit_note_summarization_corpus](https://github.com/microsoft/clinical_visit_note_summarization_corpus)**
  * **Features:** Standardized corpus and evaluation methodologies for clinical dialogue synthesis, providing benchmarks on factual consistency and omission rates.
* **[FreedomIntelligence/Awesome-Specialized-Medical-LLMs](https://github.com/FreedomIntelligence/Awesome-Specialized-Medical-LLMs)**
  * **Features:** Comprehensive curation of medical LLMs, prompt frameworks, and clinical reasoning benchmark suites.

### 1.3 Markdown-to-Anki Parsers (Python Automation)
* **[Mochitto/Markdown2Anki](https://github.com/Mochitto/Markdown2Anki)**: Highly customizable Python CLI tool parsing structured Markdown headings, code blocks, and clozes directly into Anki packages (`.apkg`).
* **[mlcivilengineer/obsankipy](https://github.com/mlcivilengineer/obsankipy)**: Python bridge synchronizing Markdown vaults with Anki via AnkiConnect, supporting cloze deletions and math notations.
* **[ObsidianToAnki/Obsidian_to_Anki](https://github.com/ObsidianToAnki/Obsidian_to_Anki)**: Regex-driven parser for extracting inline clozes and Q&A blocks from Markdown files.

---

## Part 2: Evidence-Based Pedagogical Structures for Medical Revision

### 2.1 The "Illness Script" Cognitive Architecture
Cognitive science in medical education shows that expert diagnosticians organize disease knowledge as **Illness Scripts** rather than exhaustive textbooks. An illness script comprises 3 core mental buckets:
1. **Precedent / Enabling Conditions:** Host epidemiology (age, sex, ethnicity), environmental exposures, comorbidities, and genetic predispositions.
2. **Pathophysiological Fault:** The fundamental cellular/biochemical insult, anatomy, and mechanism of disease.
3. **Clinical Consequences:** The temporal evolution (hyperacute, acute, subacute, chronic), cardinal symptoms, physical signs, and diagnostic test perturbations.

### 2.2 USMLE First Aid High-Yield 7-Anchor Structure
Every medical condition summarized for board exam review (USMLE Step 1/2 CK, Shelf, or Iranian Medical Residency Exam) should follow this strict 7-anchor chain:

```text
┌────────────────────────────────────────────────────────┐
│ 1. Core Pathophysiology & Underlying Mechanism         │
├────────────────────────────────────────────────────────┤
│ 2. Patient Demographics & Enabling Risk Factors        │
├────────────────────────────────────────────────────────┤
│ 3. Classic Presentation & Pathognomonic Triads/Buzzwords│
├────────────────────────────────────────────────────────┤
│ 4. Initial Diagnostic Test (Best Next Step / Screening)│
├────────────────────────────────────────────────────────┤
│ 5. Gold Standard Diagnostic Test (Definitive / Biopsy) │
├────────────────────────────────────────────────────────┤
│ 6. First-Line Treatment & Acute Management Protocol    │
├────────────────────────────────────────────────────────┤
│ 7. High-Yield Exam Traps, Pertinent Negatives & Pearls │
└────────────────────────────────────────────────────────┘
```

#### Diagnostic Crux: "Best Next Step" vs. "Gold Standard"
* A major source of exam failure and cognitive error is confusing the **Initial Screening / Best Next Step** (typically cheap, non-invasive, safe, e.g., Ultrasound, ECG, CXR, D-Dimer) with the **Gold Standard / Definitive Confirmation** (often invasive or resource-heavy, e.g., Biopsy, Cardiac Catheterization, CT Pulmonary Angiography, Endoscopy). Summaries must explicitly delineate both.

### 2.3 Differential Diagnosis & "Pivot Feature" Comparison Tables
Rather than static lists, differential diagnoses should be organized around **Pivot Features** (discriminating factors that force a fork in the clinical reasoning path):

| Condition / Pathology | Defining Features (What matches) | **Pivot / Discriminating Features** (The clinching discriminator) | Pertinent Negatives (What must NOT be present) | Best Next Step / Confirmatory Test |
| :--- | :--- | :--- | :--- | :--- |
| **Disease A** *(Target)* | Core presenting symptoms matching the illness script | Specific pathognomonic sign, biomarker, or unique radiation of pain | Absence of atypical systemic features | First-line targeted test |
| **Disease B** *(Look-alike)* | Overlapping symptoms (e.g., retrosternal chest pain) | Unique differentiator (e.g., pain worsens on recumbency, PR depression) | Absence of elevated troponin | Confirmatory test for B |
| **Disease C** *(Can't-Miss Emergency)* | Shared acute presentation | Red flag / catastrophe sign (e.g., blood pressure differential between arms) | Mediastinal widening absent (does not rule out completely) | Emergency CTA / Echo |

### 2.4 Pharmacological Class Comparison Matrix
When summarizing medical pharmacology lectures, drug classes must be structured to prevent interference:

| Drug Class & Suffix | Representative Agents | Molecular Mechanism of Action | High-Yield Clinical Indications | Cardinal Toxicities & Black Box Warnings | Absolute Contraindications & Dangerous Interactions |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **e.g., SGLT2 Inhibitors** (`-gliflozin`) | Empagliflozin, Dapagliflozin | Inhibits SGLT2 in PCT -> glucosuria & natriuresis | T2DM, HFrEF/HFpEF (reduces mortality), CKD | Euglycemic DKA, mycotic genital infections, Fournier gangrene | eGFR < 20-30 mL/min (glycemic efficacy drops), severe renal impairment |

### 2.5 Active Recall: Cornell Notes for Medical Education
Passive reading produces rapid forgetting (Ebbinghaus curve). Cornell Notes adapted for medical lectures enforce the **Two-Pass Method**:
1. **Right Column (Notes Column - 70%):** Shorthand mechanisms, pathways, and clinical facts captured during or from the lecture.
2. **Left Column (Cue Column - 30%):** Formulated strictly as **Active Recall Questions** (The "Fold Test").
   * *Example:* Instead of writing "S3 heart sound," write *"In what physiological vs. pathological states is an S3 gallop audible, and what physical event creates the sound?"*
3. **Bottom Summary Box (2-3 lines):** Concise synthesis connecting the molecular mechanism to the bedside presentation.

### 2.6 Spaced Repetition: Wozniak's 20 Rules & Cloze Deletion Best Practices
When turning lecture summaries into Anki flashcards, adherence to Piotr Wozniak's principles is paramount:
1. **Rule 1 & 2 (Comprehension Before Memorization):** Never generate flashcards for ungrounded raw facts. The student must understand the mechanism first.
2. **Rule 4 (Minimum Information Principle / Atomicity):** **One fact per card.** A card asking for "symptoms, diagnosis, and treatment of Wilson disease" will become a high-frustration "leech" card.
3. **Rules 9 & 10 (Avoid Enumerations & Lists):** Never ask *"What are the 5 diagnostic criteria for Kawasaki disease?"* Instead, use cloze deletions or targeted prompts for individual criteria (e.g., *"In Kawasaki disease, fever must persist for at least {{c1::5 days}}"*).
4. **Rule 11 (Combat Interference):** Explicitly contrast easily confused diseases on the card back (e.g., Crohn vs. UC, Niemann-Pick vs. Tay-Sachs).
5. **Cloze Deletion Guidelines:**
   * **Single-Target Clozes:** Avoid sentence-overload like `{{c1::Disease}} causes {{c2::symptom}} treated with {{c3::drug}}`. Split into distinct atomic cards.
   * **Context Cues:** Provide enough clinical context so the card tests medical knowledge, not sentence pattern recognition.
   * **Image Occlusion (Rule 8):** Essential for gross anatomy, dermatological rashes, radiology (CXR/CT), and histology.

---

## Part 3: Architecture for Adapting Existing Tools & Skills

### 3.1 Leveraging the Existing `medical-lecture-transcriber` Plugin
In the local environment (`C:\Users\Hill\.gemini\config\plugins\medical-lecture-transcriber`), there is already a production-grade pipeline (v5.8.7) that:
1. Chunks audio and runs Gemini multimodal transcription.
2. Extracts shapes, tables, and images from PowerPoint (`.pptx`) and PDF (`.pdf`).
3. Generates structured JSON batches (`batches/batch_XX_YY.json`) and merged documents.

**Integration Strategy:**
Instead of re-inventing slide parsing or audio chunking from scratch, the new **Medical Summarizer & Revision Skill (`اسکیل خلاصه نویسی پزشکی`)** can serve as:
1. **A Standalone Module:** Can ingest any raw medical text, Word docx, PDF lecture notes, or transcript directly.
2. **A Downstream Mode for Transcripts:** Can ingest the structured outputs produced by `medical-lecture-transcriber` and automatically synthesize:
   * **Module A: High-Yield Revision Master Sheet** (Physiopath 6-step flow + Pivot Feature Differential Tables).
   * **Module B: Flowcharts & Mechanism Diagrams** (Mermaid / Text Flowcharts of pathophysiological cascades).
   * **Module C: Anki-Ready Cloze Flashcard Deck** (formatted for AnKing note type or Markdown2Anki).
   * **Module D: Self-Assessment Questions & Exam Pearls** (Active Recall Q&A).

### 3.2 Anki Export Compatibility Format
The generated flashcards should be outputted in both:
1. **Markdown / Obsidian-to-Anki format:**
   ```markdown
   #flashcards/cardiology
   In patients with suspected acute pulmonary embolism, the initial screening test of choice in a low-risk patient is {{c1::D-Dimer}}, whereas the gold-standard confirmatory imaging is {{c2::CT Pulmonary Angiography (CTPA)}}.
   ?
   **Extra:** Westermark sign (focal oligemia) and Hampton hump (wedge infarct) on CXR are specific but rare. In pregnancy or renal failure, consider V/Q scan.
   ```
2. **TSV / CSV Format (Ready for Direct Anki Import into AnKing Overhaul schema):**
   * Fields: `Text` \t `Extra` \t `Lecture Notes` \t `Tags`

---

## Part 4: Recommended Specification for the Summarizer Skill

### Prompt Framework: Multi-Stage Cognitive Synthesizer
The skill should utilize a 4-pass prompt pipeline:
1. **Pass 1: Entity & Concept Extraction** (extract diseases, drugs, pathophysiology, diagnostics, lab cutoffs, epidemiological associations).
2. **Pass 2: Illness Script & High-Yield Structuring** (organize extracted concepts into the 6-anchor Physiopath layout).
3. **Pass 3: Comparative Analysis & Table Synthesis** (detect look-alikes and generate Pivot Feature Differential tables and Drug Comparison matrices).
4. **Pass 4: Active Recall & Flashcard Distillation** (generate atomic cloze cards obeying Wozniak's 20 rules, along with 2-3 clinical review questions).

---

## References & Citations
1. **Wozniak, P.** (1999). *Effective learning: Twenty rules of formulating knowledge.* SuperMemo.
2. **Bordage, G.** (1994). *Prototypes and semantic qualifiers: from past evidence to a current model of clinical reasoning.* Medical Education.
3. **Roediger, H. L., & Karpicke, J. D.** (2006). *The power of testing memory: Basic research and implications for educational practice.* Perspectives on Psychological Science.
4. **Lorin, S., et al.** (2008). *Clinical pearls: a tradition in medical education.* Academic Medicine.
5. **AnKing VIP Project.** *AnKing Note Types & AnKingMed Overhaul.* [github.com/AnKing-VIP/AnKing-Note-Types](https://github.com/AnKing-VIP/AnKing-Note-Types).
6. **Mochitto.** *Markdown2Anki: Markdown to Anki Flashcard Converter.* [github.com/Mochitto/Markdown2Anki](https://github.com/Mochitto/Markdown2Anki).
7. **Neilus03.** *recsum: High-recall clinical document summarization.* [github.com/Neilus03/recsum](https://github.com/Neilus03/recsum).
8. **329162516.** *MedPrompt: Healthcare Prompts & Chains for LLMs.* [github.com/329162516/medprompt](https://github.com/329162516/medprompt).
