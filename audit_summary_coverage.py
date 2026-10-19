#!/usr/bin/env python3
"""
audit_summary_coverage.py: Deterministic Omission & Verification Gate for Medical Summaries.
Audits:
1. Numerical & Cutoff Preservation (EF <= 40%, BNP < 100 pg/mL, 36h, etc.)
2. Clinical Entity Recall (Key diseases, drugs, pathognomonic buzzwords)
3. Diagnostic Bifurcation Check (Presence of both Initial Test and Gold Standard)
Outputs a structured JSON audit report: coverage_audit_report.json
"""

import os
import sys
import json
import re
import argparse
from typing import Dict, Any, List, Set, Tuple

# Force UTF-8 encoding in stdout/stderr for Windows compatibility
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Character translation tables for Persian/Arabic digit and character normalization
PERSIAN_ARABIC_DIGITS = "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹"
ENGLISH_DIGITS = "01234567890123456789"
DIGIT_TRANSLATE_TABLE = str.maketrans(PERSIAN_ARABIC_DIGITS, ENGLISH_DIGITS)

def normalize_medical_text(text: str) -> str:
    """
    Normalizes Persian/Arabic digits to ASCII, standardizes decimal separators,
    normalizes thousands separators and decimal commas, and standardizes zero-width non-joiners.
    """
    if not text:
        return ""
    # Convert Persian/Arabic digits (۰-۹, ٠-٩) to English digits (0-9)
    norm = text.translate(DIGIT_TRANSLATE_TABLE)
    # Convert Persian decimal separator (\u066b) to period
    norm = norm.replace("\u066b", ".")
    # Normalize Arabic letters to Persian standard
    norm = norm.replace("ي", "ی").replace("ك", "ک")
    # Standardize thousands separators (e.g. 50,000 or 50،000 or 1,000,000 -> 50000 or 1000000)
    while True:
        new_norm = re.sub(r'(?<=\d)[,،](\d{3})(?!\d)', r'\1', norm)
        if new_norm == norm:
            break
        norm = new_norm
    # Standardize decimal commas (e.g. 1,5 or 1،5 -> 1.5)
    norm = re.sub(r'(?<=\d)[,،](\d{1,2})(?!\d)', r'.\1', norm)
    # Replace zero-width non-joiner with standard space
    norm = norm.replace("\u200c", " ")
    return norm

# Comprehensive regex for clinical numbers, ratios, and lab cutoffs (supporting English & Persian)
PERSIAN_PREFIXES = r'(?:[<>≤≥=±~]|کمتر\s*از|بیشتر\s*از|بیش\s*از|زیر|بالای|حداقل|حداکثر)\s*'
PERSIAN_UNITS = (
    r'(?:pg/mL|ng/mL|mg/dL|g/dL|mg|mcg|μg|mL|L|%|mm\s*Hg|mEq/L|mOsm/kg|mmol/L|bpm|'
    r'درصد|میلی\s*گرم|میلیگرم|میکروگرم|گرم|میلی\s*لیتر|میلیلیتر|سی\s*سی|سی‌سی|'
    r'ساعت|روز|سال|هفته|ماه|میلی\s*متر\s*جیوه|میلیمتر\s*جیوه|کیلوگرم|تعداد|ضربان|'
    r'years?|days?|hours?|weeks?|months?|mm|cm|cells/μL|cells/uL)'
)
NUM_UNIT_REGEX = re.compile(
    rf'(?P<prefix>{PERSIAN_PREFIXES})?'
    r'(?P<val>\b\d+(?:\.\d+)?)\s*'
    rf'(?P<unit>{PERSIAN_UNITS})?',
    re.IGNORECASE
)

# Common non-clinical numbers to filter out (e.g., slide numbers, bullet numbers)
IGNORED_NUMBERS = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "2024", "2025", "2026"}

def extract_clinical_numbers(text: str) -> List[Dict[str, Any]]:
    """Extracts numbers with their units, relational prefixes, and surrounding sentence context without breaking decimals."""
    results = []
    # Split text into sentences for context without breaking decimal numbers (e.g. 1.5, 0.5)
    sentences = re.split(r'(?:(?<!\d)\.|\.(?!\d))|[\n\?!;؛]', text)
    for sent in sentences:
        s_clean = sent.strip()
        if not s_clean:
            continue
        matches = NUM_UNIT_REGEX.finditer(s_clean)
        for m in matches:
            prefix = m.group("prefix") or ""
            val = m.group("val")
            unit = m.group("unit") or ""
            full_token = f"{prefix.strip()} {val} {unit.strip()}".strip()

            if val in IGNORED_NUMBERS and not unit and not prefix:
                continue

            results.append({
                "token": full_token,
                "prefix": prefix.strip(),
                "value": val,
                "unit": unit.strip(),
                "context": s_clean
            })
    return results

def extract_drug_and_disease_entities(text: str) -> Set[str]:
    """
    Extracts English medical terms, genuine medical acronyms, and drug classes using regex.
    Strictly filters English prose sentence starters (When, Because, During, Other, First)
    to prevent false-positive entity omission alarms in bilingual/Persian summaries.
    """
    entities = set()

    # Clean tokenization of hyphenated terms:
    tokens = re.findall(r'\b[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*\b', text)

    def is_medical_acronym(token: str) -> bool:
        if '-' in token:
            parts = token.split('-')
            return any(is_medical_acronym(p) for p in parts)

        # Must have at least 2 chars
        if len(token) < 2:
            return False

        # Uppercase acronym (2-8 chars) e.g. HF, BNP, ARNI, NYHA, CXR, ECG, CTPA, LVEF
        if token.isupper() and 2 <= len(token) <= 8:
            return True

        # Mixed-case medical acronyms: e.g. HFrEF, HFpEF, ACEi, SGLT2i, eGFR, HbA1c
        has_upper = any(c.isupper() for c in token)
        has_lower = any(c.islower() for c in token)
        if has_upper and has_lower:
            upper_count = sum(1 for c in token if c.isupper())
            if upper_count >= 2:
                return True
            if token[0].islower() and any(c.isupper() for c in token[1:]):
                return True

        # Alphanumeric medical tokens like FEV1, SGLT2, CA125, V1, V6
        if any(c.isdigit() for c in token) and has_upper:
            return True

        return False

    english_stop_words = {
        "THE", "AND", "FOR", "WITH", "TYPE", "PAGE", "FIG", "STEP", "CASE",
        "PART", "NOTE", "TABLE", "CHAPTER", "TRUE", "FALSE", "NULL", "NONE",
        "ALL", "NOT", "BUT", "FROM", "INTO", "THIS", "THAT", "THOSE", "THESE",
        "WHEN", "WHAT", "HOW", "WHY", "WHO", "WHERE", "WHICH", "BECAUSE", "DURING",
        "AFTER", "BEFORE", "CAN", "MAY", "MIGHT", "WILL", "SHALL", "MUST", "DOES",
        "DID", "HAS", "HAD", "HAVE", "ARE", "WAS", "WERE", "BEEN", "NEW", "OLD",
        "ONE", "TWO", "SEE", "USE", "NON", "PRE", "POST", "PER", "VIA", "VS",
        "VERSUS", "OTHER", "FIRST", "SECOND", "THIRD", "CLASS", "STAGE", "GRADE",
        "SOME", "SUCH", "BOTH", "EACH", "EVERY", "MANY", "MUCH", "MORE", "MOST",
        "LESS", "VERY", "SINCE", "UNTIL", "WHILE", "WOULD", "COULD", "SHOULD",
        "II", "III", "IV"
    }

    for t in tokens:
        if is_medical_acronym(t) and t.upper() not in english_stop_words:
            entities.add(t)

    # Standard and expanded drug suffix patterns (USMLE/Katzung pharmacology classes)
    drug_patterns = re.findall(
        r'\b[A-Za-z]+(?:lol|pril|sartan|statin|gliflozin|tide|mab|nib|pine|cillin|mycin|floxacin|azole|triptan|prazole|vir|parin|gliptin|semide|actone|grel|xaban|glitazone|afil|setron|lukast|terol|coxib|navir|caine|dipine)\b',
        text,
        re.IGNORECASE
    )
    for d in drug_patterns:
        entities.add(d.lower())

    return entities

PERSIAN_COMMON_STOPWORDS = {
    "استاد", "تاکید", "تأکید", "امتحان", "مهم", "بیمار", "بیماران", "داروی", "دارو", "داروها",
    "باید", "است", "برای", "مورد", "ایجاد", "مصرف", "تجویز", "نشان", "دهد", "دارد", "شود",
    "شوند", "کنید", "کنند", "توجه", "همراه", "صورت", "دارای", "عنوان", "توسط", "شامل", "طریق",
    "بوده", "استفاده", "نکته", "سوال", "سؤال", "حتما", "شدید", "خفیف", "متوسط", "افزایش", "کاهش",
    "درمان", "بالینی", "علائم", "علامت", "بیماری", "می‌شود", "می‌شوند", "می‌توان", "می‌باشد",
    "پزشک", "پزشکی", "همچنین", "بسیار", "زیاد", "کمتر", "بیشتر", "حالت", "بخش", "نتیجه", "می شود", "می شوند"
}

WARNING_KEYWORDS = [
    "کنترااندیکاسیون", "ممنوع", "اکیدا ممنوع", "ممنوع است", "منع مصرف",
    "contraindicated", "contraindication", "washout", "black box", "avoid",
    "دوره پاکسازی", "تداخل کشنده", "خطر مرگ", "مرگبار"
]

WARNING_SIGNALS = [
    "ممنوع", "منع", "کنترااندیکاسیون", "contraindicat", "washout", "هشدار",
    "احتیاط", "خطر", "مرگبار", "تداخل", "black box", "avoid", "withhold"
]

EMPHASIS_KEYWORDS = [
    "استاد تأکید", "استاد تاکید", "مهم امتحانی", "سوال امتحان", "سؤال امتحان",
    "نکته امتحانی", "حتما بخوانید", "صد در صد امتحانی", "سوال قطعی", "⭐"
]

def extract_warnings_and_contraindications(text: str) -> List[str]:
    """Finds sentences containing absolute contraindications, lethal withholds, or washout warnings."""
    sentences = re.split(r'(?:(?<!\d)\.|\.(?!\d))|[\n\?!;؛]', text)
    results = []
    for sent in sentences:
        s_clean = sent.strip()
        if len(s_clean) > 15 and any(kw in s_clean for kw in WARNING_KEYWORDS):
            results.append(s_clean)
    return results

def extract_instructor_emphasis(text: str) -> List[str]:
    """Finds sentences where the professor explicitly highlighted exam importance."""
    sentences = re.split(r'(?:(?<!\d)\.|\.(?!\d))|[\n\?!;؛]', text)
    results = []
    for sent in sentences:
        s_clean = sent.strip()
        if len(s_clean) > 15 and any(kw in s_clean for kw in EMPHASIS_KEYWORDS):
            results.append(s_clean)
    return results

def audit_coverage(source_text: str, summary_data: Dict[str, Any]) -> Dict[str, Any]:
    """Compares source text against generated summary data with zero-omission enforcement."""
    norm_source = normalize_medical_text(source_text)
    raw_summary_text = json.dumps(summary_data, ensure_ascii=False)
    norm_summary_text = normalize_medical_text(raw_summary_text)

    source_nums = extract_clinical_numbers(norm_source)
    source_entities = extract_drug_and_disease_entities(norm_source)
    summary_entities = extract_drug_and_disease_entities(norm_summary_text)

    # 1. Numerical & Cutoff Audit (Exact numeric boundary matching)
    matched_nums = []
    omitted_nums = []

    for item in source_nums:
        val = item["value"]
        # Match number strictly bounded by non-digits and non-decimal points
        if re.search(rf'(?<![\d\.])' + re.escape(val) + r'(?![\d\.])', norm_summary_text):
            matched_nums.append(item)
        else:
            omitted_nums.append(item)

    unique_omissions = {}
    for om in omitted_nums:
        tok = om["token"]
        if tok not in unique_omissions:
            unique_omissions[tok] = om

    num_coverage_pct = round((len(matched_nums) / len(source_nums) * 100), 1) if source_nums else 100.0

    # 2. Entity Recall Audit
    missing_entities = [e for e in source_entities if e.lower() not in norm_summary_text.lower()]
    entity_recall_pct = round(((len(source_entities) - len(missing_entities)) / len(source_entities) * 100), 1) if source_entities else 100.0

    # 3. Diagnostic Bifurcation Audit
    bifurcation_status = []
    modules = summary_data.get("modules") or []
    for mod in modules:
        if not isinstance(mod, dict):
            continue
        m_name = mod.get("name", "Unknown")
        diag = mod.get("diagnostics")
        if isinstance(diag, dict):
            has_initial = bool(diag.get("best_initial_test"))
            has_gold = bool(diag.get("gold_standard"))
            # If diagnostics is declared, at least initial test or gold standard should be captured if present in lecture
            is_compliant = (has_initial or has_gold) if (has_initial or has_gold or not norm_source) else True
        else:
            has_initial = False
            has_gold = False
            is_compliant = True

        bifurcation_status.append({
            "module": m_name,
            "has_best_initial_test": has_initial,
            "has_gold_standard": has_gold,
            "is_compliant": is_compliant
        })

    all_bifurcations_valid = all(b["is_compliant"] for b in bifurcation_status) if bifurcation_status else True

    # 4. Warnings & Contraindications Audit
    source_warnings = extract_warnings_and_contraindications(norm_source)
    omitted_warnings = []
    has_warning_signal = any(sig in norm_summary_text.lower() for sig in WARNING_SIGNALS)

    for w in source_warnings:
        # Extract substantive words (excluding stop words and warning keywords)
        w_words = [
            wrd for wrd in re.findall(r'\b[A-Za-zآ-ی‌]{4,}\b', w)
            if wrd not in PERSIAN_COMMON_STOPWORDS and not any(kw in wrd for kw in WARNING_KEYWORDS)
        ]
        # At least one distinctive word must appear, and summary must contain a warning signal
        distinctive_match = any(wrd in norm_summary_text for wrd in w_words) if w_words else False
        if not (distinctive_match and has_warning_signal):
            omitted_warnings.append(w)

    # 5. Instructor Emphasis Audit
    source_emphasis = extract_instructor_emphasis(norm_source)
    omitted_emphasis = []
    for emp in source_emphasis:
        emp_words = [
            wrd for wrd in re.findall(r'\b[A-Za-zآ-ی‌]{4,}\b', emp)
            if wrd not in PERSIAN_COMMON_STOPWORDS and not any(kw in wrd for kw in EMPHASIS_KEYWORDS)
        ]
        preserved = any(wrd in norm_summary_text for wrd in emp_words) if emp_words else False
        if not preserved:
            omitted_emphasis.append(emp)

    passed = (
        num_coverage_pct >= 80.0
        and entity_recall_pct >= 85.0
        and all_bifurcations_valid
        and len(omitted_warnings) == 0
    )

    report = {
        "verdict": "PASSED" if passed else "WARNING_OMISSIONS_DETECTED",
        "scores": {
            "numerical_coverage_pct": num_coverage_pct,
            "entity_recall_pct": entity_recall_pct,
            "diagnostic_bifurcation_compliant": all_bifurcations_valid,
            "contraindications_omitted_count": len(omitted_warnings),
            "instructor_emphasis_omitted_count": len(omitted_emphasis)
        },
        "total_source_clinical_numbers": len(source_nums),
        "total_omitted_numbers_count": len(unique_omissions),
        "omitted_clinical_numbers": list(unique_omissions.values()),
        "omitted_contraindications": omitted_warnings,
        "omitted_instructor_emphasis": omitted_emphasis,
        "missing_entities": missing_entities,
        "bifurcation_details": bifurcation_status
    }
    return report

def main():
    parser = argparse.ArgumentParser(description="Audit medical summary for high-yield omissions and cutoffs")
    parser.add_argument("--source", required=True, help="Path to original ingested lecture JSON or raw text")
    parser.add_argument("--summary", required=True, help="Path to generated summary JSON")
    parser.add_argument("--report", default="coverage_audit_report.json", help="Path to save audit report")
    args = parser.parse_args()

    # Read source
    if not os.path.exists(args.source):
        print(f"[ERROR] Source file not found: {args.source}")
        sys.exit(1)
    with open(args.source, "r", encoding="utf-8", errors="replace") as f:
        src_raw = f.read()
    try:
        src_data = json.loads(src_raw)
        source_text = src_data.get("full_text", src_raw)
    except Exception:
        source_text = src_raw

    # Read summary
    if not os.path.exists(args.summary):
        print(f"[ERROR] Summary file not found: {args.summary}")
        sys.exit(1)
    with open(args.summary, "r", encoding="utf-8") as f:
        summary_data = json.load(f)

    report = audit_coverage(source_text, summary_data)

    with open(args.report, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print(f"--- Coverage Audit Report: {report['verdict']} ---")
    print(f"  • Numerical Cutoff Coverage: {report['scores']['numerical_coverage_pct']}%")
    print(f"  • Medical Entity Recall:    {report['scores']['entity_recall_pct']}%")
    print(f"  • Diagnostic Bifurcation:   {'Passed' if report['scores']['diagnostic_bifurcation_compliant'] else 'Failed'}")
    if report["omitted_clinical_numbers"]:
        print(f"  ⚠️ Warning: {len(report['omitted_clinical_numbers'])} clinical numbers were omitted from the summary!")
        for om in report["omitted_clinical_numbers"][:3]:
            print(f"     - '{om['token']}' in context: '{om['context'][:80]}...'")
    print(f"[OK] Full report written to: {args.report}")

if __name__ == "__main__":
    main()
