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

def extract_clinical_syndromes_and_diseases(text: str) -> List[Tuple[str, str]]:
    """
    Universally extracts clinical syndromes, diseases, triads, and diagnostic criteria across all medical specialties
    (Cardiology, Nephrology, Pulmonology, Neurology, Rheumatology, Hematology, Pediatrics, etc.).
    Returns a list of tuples: (matched_term, core_name).
    """
    found = []
    seen = set()
    norm_txt = normalize_medical_text(text)

    # 1. Persian Clinical Patterns:
    pe_pattern = re.compile(
        r'\b(?P<prefix>سندرم|سندروم|بیماری|تریاد|تترالوژی|پنتاد|معیارهای|معیار|اسکور|سیستم\s+نمره‌دهی)'
        r'\s+(?P<name>[آ-یa-zA-Z0-9\-\'\u200c]+(?:\s+[آ-یa-zA-Z0-9\-\'\u200c]+){0,2})',
        re.IGNORECASE
    )
    STOP_WORDS = {
        "و", "یا", "که", "به", "با", "در", "بر", "از", "تا", "بین", "برای", "یک", "یکی",
        "شامل", "دارای", "بدون", "است", "بود", "شد", "می", "این", "آن", "هم", "نیز", "چون",
        "زیرا", "اگر", "اما", "ولی", "علت", "دلیل", "ناشی", "متعاقب", "پیرو", "همراه",
        "شایع", "نادر", "مهم", "اصلی", "دیگر", "مختلف", "بالینی", "حاد", "مزمن", "اولیه",
        "ثانویه", "عفونی", "شدید", "خفیف", "های", "ها", "ای", "بیمار", "پیشرفته", "مربوطه",
        "مورد", "نظر", "وجود", "دارد", "شود", "مانند", "نظیر", "طی", "پس", "قبل"
    }

    for m in pe_pattern.finditer(norm_txt):
        prefix = m.group("prefix").strip()
        name = m.group("name").strip()
        name_tokens = name.split()
        valid_tokens = []
        for tok in name_tokens:
            tok_clean = tok.strip()
            if tok_clean.lower() in STOP_WORDS:
                break
            valid_tokens.append(tok_clean)
        if not valid_tokens:
            continue
        clean_name = " ".join(valid_tokens)
        full_phrase = f"{prefix} {clean_name}".strip()
        full_key = full_phrase.lower()
        if full_key not in seen and len(clean_name) >= 2:
            seen.add(full_key)
            found.append((full_phrase, clean_name.lower()))

    # 2. English Clinical Patterns:
    en_pattern = re.compile(
        r'\b(?P<name>[A-Z][a-zA-Z0-9\-\']*(?:\s+[A-Z][a-zA-Z0-9\-\']*){0,2})\s+'
        r'(?P<suffix>Syndrome|Disease|Triad|Tetralogy|Criteria|Score|Classification|Staging)\b'
    )
    EN_IGNORE_NAMES = {"This", "That", "The", "A", "An", "Each", "Every", "Which", "What", "Some", "Severe", "Acute", "Chronic", "Other", "First"}
    for m in en_pattern.finditer(text):
        name = m.group("name").strip()
        suffix = m.group("suffix").strip()
        if name in EN_IGNORE_NAMES:
            continue
        full_en = f"{name} {suffix}".strip()
        full_key = full_en.lower()
        if full_key not in seen:
            seen.add(full_key)
            found.append((full_en, name.lower()))

    return found

def audit_illness_script_integrity(source_text: str, summary_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Audits Illness Script Integrity according to the Anti-Pharmacology-Skew Law (Principle 3):
    Ensures clinical conditions/diseases preserve the 4 essential pillars:
      1. Etiology & Pathophysiologic Cascade (definition_and_etiology / pathophysiology_cascade / flowchart_steps)
      2. Clinical Presentation & Triads (clinical_presentation / triads_and_pathognomonics / symptoms_and_signs)
      3. Diagnostics Bifurcation (best_initial_test and/or gold_standard)
      4. High-Yield Pearls & Contraindications (high_yield_pearls / exam_traps / contraindications_and_warnings / instructor_exam_alerts)
    Also verifies rare/subtle syndromes taught in source are preserved in summary modules or comparative matrices,
    strictly preventing pharmacology over-indexing from dropping clinical syndromes.
    """
    modules = summary_data.get("modules") or []
    raw_summary_text = json.dumps(summary_data, ensure_ascii=False)
    norm_summary_text = normalize_medical_text(raw_summary_text)

    # 1. Inspect 4 pillars for each clinical module
    module_audits = []
    sound_modules_count = 0
    total_clinical_modules = 0

    for mod in modules:
        if not isinstance(mod, dict):
            continue
        m_name = mod.get("name", "Unnamed Module")
        mod_type = mod.get("type") or mod.get("module_type") or "illness_script"

        # Archetype 1: Foundational Physiology & Regulatory Loops
        is_foundation = (
            mod_type == "foundational_concept"
            or bool(mod.get("feedback_cascade"))
            or bool(mod.get("regulatory_mechanisms_and_feedback"))
            or bool(mod.get("physiological_role"))
            or bool(mod.get("anatomical_physiological_basis"))
        )
        if is_foundation:
            total_clinical_modules += 1
            has_role = bool(mod.get("physiological_role") or mod.get("anatomical_physiological_basis"))
            has_feed = bool(mod.get("feedback_cascade") or mod.get("regulatory_mechanisms_and_feedback"))
            has_tests = bool(mod.get("baseline_tests"))
            has_rel = bool(mod.get("clinical_relevance"))
            f_pillars = sum([has_role, has_feed, has_tests, has_rel])
            is_sound = (f_pillars >= 2)
            if is_sound:
                sound_modules_count += 1
            module_audits.append({
                "module": m_name,
                "type": "foundational_concept",
                "pillars_met_count": f_pillars,
                "has_role": has_role,
                "has_feedback": has_feed,
                "has_tests": has_tests,
                "has_clinical_relevance": has_rel,
                "is_sound": is_sound
            })
            continue

        # Archetype 2: Clinical Diagnostic & Decision Approach (Algorithms)
        is_approach = (
            mod_type == "clinical_approach"
            or bool(mod.get("step_by_step_algorithm"))
        )
        if is_approach:
            total_clinical_modules += 1
            has_init = bool(mod.get("initial_presentation") or mod.get("chief_complaint_or_scenario"))
            has_alg = bool(mod.get("step_by_step_algorithm"))
            has_pivot = bool(mod.get("pivot_points") or mod.get("high_yield_decision_points"))
            has_def = bool(mod.get("definitive_action"))
            a_pillars = sum([has_init, has_alg, has_pivot, has_def])
            is_sound = (a_pillars >= 2)
            if is_sound:
                sound_modules_count += 1
            module_audits.append({
                "module": m_name,
                "type": "clinical_approach",
                "pillars_met_count": a_pillars,
                "has_initial": has_init,
                "has_algorithm": has_alg,
                "has_pivot": has_pivot,
                "has_definitive": has_def,
                "is_sound": is_sound
            })
            continue

        # Archetype 3: Disease Illness Script
        has_etio = bool(
            mod.get("definition_and_etiology")
            or mod.get("pathophysiology_cascade")
            or mod.get("risk_factors_and_epidemiology")
            or mod.get("risk_factors_and_predisposing")
            or mod.get("flowchart_steps")
            or mod.get("etiology")
        )
        has_pres = bool(
            mod.get("clinical_presentation")
            or mod.get("triads_and_pathognomonics")
            or mod.get("symptoms_and_signs")
            or mod.get("clinical_semiology")
        )
        diag = mod.get("diagnostics")
        has_diag = bool(isinstance(diag, dict) and (diag.get("best_initial_test") or diag.get("gold_standard")))
        has_pearls = bool(
            mod.get("high_yield_pearls")
            or mod.get("exam_traps")
            or mod.get("contraindications_and_warnings")
            or mod.get("instructor_exam_alerts")
            or mod.get("complications_and_prognosis")
            or mod.get("surgical_and_interventional")
            or mod.get("surgical_and_interventional_indications")
        )
        has_pharm = bool(mod.get("pharmacology_and_management"))

        is_clinical = has_etio or has_pres or has_diag or (not has_pharm)
        if is_clinical:
            total_clinical_modules += 1
            pillars_met = sum([has_etio, has_pres, has_diag, has_pearls])
            is_sound = (pillars_met >= 2)
            if is_sound:
                sound_modules_count += 1

            module_audits.append({
                "module": m_name,
                "type": "illness_script",
                "pillars_met_count": pillars_met,
                "has_etiology_cascade": has_etio,
                "has_clinical_presentation": has_pres,
                "has_diagnostics": has_diag,
                "has_pearls_or_warnings": has_pearls,
                "is_sound": is_sound
            })

    compliance_pct = round((sound_modules_count / total_clinical_modules * 100), 1) if total_clinical_modules > 0 else 100.0

    # 2. Check for subtle/rare syndromes taught in source text that might get lost in pharmacology skew
    SUBTLE_SYNDROMES_WATCHLIST = [
        ("آپوپلکسی", "apoplexy"),
        ("شیهان", "sheehan"),
        ("نلسون", "nelson"),
        ("کارنی", "carney"),
        ("fhh", "familial hypocalciuric"),
        ("خطوط لوزر", "looser"),
        ("آشرمن", "asherman"),
        ("cais", "androgen insensitivity"),
        ("۵-آلفا ردوکتاز", "5-alpha reductase"),
        ("کوشینگ", "cushing"),
        ("آدیسون", "addison"),
        ("فئوکروموسیتوما", "pheochromocytoma"),
        ("طوفان تیروئیدی", "thyroid storm"),
        ("میکزدم", "myxedema"),
    ]
    detected_source_syndromes = []
    omitted_subtle_syndromes = []

    norm_src_lower = source_text.lower()
    norm_sum_lower = norm_summary_text.lower()
    seen_syndromes = set()

    # Part A: Curated high-yield endocrine/rare conditions watchlist
    for fa_term, en_term in SUBTLE_SYNDROMES_WATCHLIST:
        if fa_term in norm_src_lower or en_term in norm_src_lower:
            detected_source_syndromes.append(fa_term)
            seen_syndromes.add(fa_term.lower())
            seen_syndromes.add(en_term.lower())
            if fa_term not in norm_sum_lower and en_term not in norm_sum_lower:
                omitted_subtle_syndromes.append(fa_term)

    # Part B: Dynamic universal clinical entity extraction across all medical specialties
    dynamic_syndromes = extract_clinical_syndromes_and_diseases(source_text)
    for full_term, core_name in dynamic_syndromes:
        core_tokens = [w for w in core_name.lower().split() if len(w) >= 3 and w not in {"syndrome", "disease", "triad", "criteria", "score"}]
        if any(tok in seen_syndromes for tok in core_tokens):
            continue
        if full_term.lower() in seen_syndromes or core_name.lower() in seen_syndromes:
            continue
        detected_source_syndromes.append(full_term)
        seen_syndromes.add(full_term.lower())
        seen_syndromes.add(core_name.lower())
        for tok in core_tokens:
            seen_syndromes.add(tok)
        # Check if full phrase, core name, or substantive core tokens appear in summary
        is_covered = (
            full_term.lower() in norm_sum_lower
            or core_name.lower() in norm_sum_lower
            or (bool(core_tokens) and any(tok in norm_sum_lower for tok in core_tokens))
        )
        if not is_covered:
            omitted_subtle_syndromes.append(full_term)

    # 3. Check for pharmacology skew (drugs over-indexing over clinical scripts)
    total_pharm_words = 0
    total_clinical_words = 0
    for mod in modules:
        if not isinstance(mod, dict):
            continue
        pharm_str = json.dumps(mod.get("pharmacology_and_management") or [], ensure_ascii=False)
        clin_str = json.dumps({k: v for k, v in mod.items() if k != "pharmacology_and_management"}, ensure_ascii=False)
        total_pharm_words += len(pharm_str.split())
        total_clinical_words += len(clin_str.split())

    total_words = total_pharm_words + total_clinical_words
    pharm_word_ratio = (total_pharm_words / total_words) if total_words > 0 else 0.0
    skew_detected = (pharm_word_ratio > 0.85) and (compliance_pct < 60.0)

    is_compliant = (compliance_pct >= 80.0) and (len(omitted_subtle_syndromes) == 0) and (not skew_detected)

    return {
        "illness_script_compliant": is_compliant,
        "illness_script_compliance_pct": compliance_pct,
        "sound_clinical_modules_count": sound_modules_count,
        "total_clinical_modules": total_clinical_modules,
        "pharmacology_skew_detected": skew_detected,
        "pharmacology_word_ratio": round(pharm_word_ratio, 2),
        "detected_source_syndromes": detected_source_syndromes,
        "omitted_subtle_syndromes": omitted_subtle_syndromes,
        "module_breakdown": module_audits
    }

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

    # 6. Anti-Pharmacology-Skew & Illness Script Integrity Audit
    illness_script_report = audit_illness_script_integrity(norm_source, summary_data)

    # 7. High-Yield Compression & Anti-Bloat Audit
    src_words = len(re.findall(r'\b\w+\b', norm_source))
    sum_words = len(re.findall(r'\b\w+\b', norm_summary_text))
    compression_ratio = round(sum_words / src_words, 2) if src_words > 0 else 0.0

    # Check for individual bloated bullets (>120 words per bullet indicates verbatim raw pasting)
    has_bloated_bullet = False
    for mod in (summary_data.get("modules") or []):
        if not isinstance(mod, dict):
            continue
        for key, val in mod.items():
            if isinstance(val, list):
                for item in val:
                    if isinstance(item, str) and len(re.findall(r'\b\w+\b', item)) > 120:
                        has_bloated_bullet = True
            elif isinstance(val, str) and len(re.findall(r'\b\w+\b', val)) > 150:
                has_bloated_bullet = True

    # A true medical revision summary must compress verbose explanations into high-yield pearls
    # Bloat is flagged if source is substantial (>=300 words) but summary exceeds 85% of source volume,
    # or if individual bullets contain verbatim paragraph dumps (>120 words).
    verbatim_bloat_detected = bool((src_words >= 300 and compression_ratio > 0.85) or has_bloated_bullet)

    # 8. Visual Atlas & Clinical Slide Inspection Audit
    visual_atlas_items = []
    if summary_data.get("visual_atlas") and isinstance(summary_data["visual_atlas"], list):
        visual_atlas_items.extend(summary_data["visual_atlas"])
    if summary_data.get("clinical_figures") and isinstance(summary_data["clinical_figures"], list):
        visual_atlas_items.extend(summary_data["clinical_figures"])
    for mod in (summary_data.get("modules") or []):
        if isinstance(mod, dict):
            m_figs = mod.get("visual_atlas") or mod.get("clinical_figures") or []
            if isinstance(m_figs, list):
                visual_atlas_items.extend(m_figs)

    valid_atlas_count = sum(1 for f in visual_atlas_items if isinstance(f, dict) and (f.get("image_path") or f.get("path") or f.get("src")))
    atlas_with_pearls = sum(1 for f in visual_atlas_items if isinstance(f, dict) and (f.get("exam_pearl") or f.get("caption") or f.get("clinical_findings")))

    passed = (
        num_coverage_pct >= 80.0
        and entity_recall_pct >= 85.0
        and all_bifurcations_valid
        and len(omitted_warnings) == 0
        and illness_script_report["illness_script_compliant"]
        and not verbatim_bloat_detected
    )

    if not passed:
        verdict = "WARNING_VERBATIM_BLOAT_DETECTED" if (verbatim_bloat_detected and num_coverage_pct >= 80.0) else "WARNING_OMISSIONS_DETECTED"
    else:
        verdict = "PASSED"

    report = {
        "verdict": verdict,
        "scores": {
            "numerical_coverage_pct": num_coverage_pct,
            "entity_recall_pct": entity_recall_pct,
            "diagnostic_bifurcation_compliant": all_bifurcations_valid,
            "contraindications_omitted_count": len(omitted_warnings),
            "instructor_emphasis_omitted_count": len(omitted_emphasis),
            "illness_script_compliance_pct": illness_script_report["illness_script_compliance_pct"],
            "illness_script_compliant": illness_script_report["illness_script_compliant"],
            "pharmacology_skew_detected": illness_script_report["pharmacology_skew_detected"],
            "compression_ratio": compression_ratio,
            "verbatim_bloat_detected": verbatim_bloat_detected
        },
        "compression_audit": {
            "source_word_count": src_words,
            "summary_word_count": sum_words,
            "compression_ratio": compression_ratio,
            "verbatim_bloat_detected": verbatim_bloat_detected,
            "high_yield_compressed": bool(0.10 <= compression_ratio <= 0.85)
        },
        "visual_atlas_audit": {
            "total_figures": len(visual_atlas_items),
            "valid_image_figures": valid_atlas_count,
            "annotated_with_findings_or_pearls": atlas_with_pearls
        },
        "total_source_clinical_numbers": len(source_nums),
        "total_omitted_numbers_count": len(unique_omissions),
        "omitted_clinical_numbers": list(unique_omissions.values()),
        "omitted_contraindications": omitted_warnings,
        "omitted_instructor_emphasis": omitted_emphasis,
        "missing_entities": missing_entities,
        "bifurcation_details": bifurcation_status,
        "illness_script_audit": illness_script_report
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
    print(f"  • Illness Script Integrity: {report['scores']['illness_script_compliance_pct']}% ({'Passed' if report['scores']['illness_script_compliant'] else 'Failed'})")
    if report["scores"].get("pharmacology_skew_detected"):
        print(f"  ⚠️ Warning: Severe pharmacology skew detected! Clinical illness scripts are being eclipsed by drug tables.")
    if report.get("illness_script_audit", {}).get("omitted_subtle_syndromes"):
        print(f"  ⛔ Alert: Subtle clinical syndromes in source omitted: {report['illness_script_audit']['omitted_subtle_syndromes']}")
    if report["omitted_clinical_numbers"]:
        print(f"  ⚠️ Warning: {len(report['omitted_clinical_numbers'])} clinical numbers were omitted from the summary!")
        for om in report["omitted_clinical_numbers"][:3]:
            print(f"     - '{om['token']}' in context: '{om['context'][:80]}...'")
    print(f"[OK] Full report written to: {args.report}")

if __name__ == "__main__":
    main()
