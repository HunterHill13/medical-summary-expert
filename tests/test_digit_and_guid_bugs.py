#!/usr/bin/env python3
"""
tests/test_digit_and_guid_bugs.py: Tight feedback loop testing:
1. Bug A: Persian/Arabic digits (۴۰, ۱۰۰) in source text must match English digits (40, 100) in summary.
2. Bug B: Anki flashcards must have deterministic GUIDs to prevent duplicates on re-import.
"""

import os
import sys
import json
import hashlib

SCRIPTS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from audit_summary_coverage import audit_coverage

def test_persian_digit_matching():
    # Source text with Persian digits
    source_text = "در ارزیابی نارسایی قلبی، کسر تخلیه‌ای بطن چپ کمتر از ۴۰٪ و بی ان پی کمتر از ۱۰۰ pg/mL است."
    # Summary with English digits
    summary_data = {
        "modules": [
            {
                "name": "نارسایی قلبی",
                "diagnostics": {
                    "best_initial_test": "اندازه‌گیری BNP کمتر از 100 pg/mL",
                    "gold_standard": "اکوکاردیوگرافی با کسر تخلیه‌ای 40%",
                    "key_labs": ["LVEF <= 40%", "BNP < 100 pg/mL"]
                }
            }
        ]
    }

    report = audit_coverage(source_text, summary_data)
    print(f"[CHECK] Verdict: {report['verdict']}")
    print(f"[CHECK] Numerical coverage: {report['scores']['numerical_coverage_pct']}%")
    print(f"[CHECK] Omitted numbers: {report['omitted_clinical_numbers']}")

    assert report["scores"]["numerical_coverage_pct"] >= 99.0, (
        f"BUG DETECTED: Persian digits failed to match English digits! Coverage: {report['scores']['numerical_coverage_pct']}%"
    )
    print("[SUCCESS] Persian digit normalization test passed!")

def test_anki_deterministic_guid_and_deck_id():
    try:
        import genanki
    except ImportError:
        print("[SKIP] genanki not installed in this environment; skipping .apkg test.")
        return

    from build_summary_pack import build_anki_apkg

    # 1. Test GUID stability across revisions of 'extra' notes
    prompt_text = "در نارسایی قلبی HFrEF، کسر تخلیه‌ای بطن چپ {{c1::کمتر یا مساوی ۴۰٪}} است."
    norm_prompt_1 = " ".join(prompt_text.strip().translate(str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")).replace('\u200c', ' ').split())
    guid_v1 = genanki.guid_for(norm_prompt_1)

    # Re-exporting when extra field is updated with additional clinical detail
    extra_v2 = "به‌روزرسانی مکانیسم: کاهش اینوتروپی منجر به افزایش LVEDP می‌شود."
    guid_v2 = genanki.guid_for(norm_prompt_1)

    assert guid_v1 == guid_v2, f"BUG: Note GUID changed upon updating extra field! {guid_v1} != {guid_v2}"
    print(f"[SUCCESS] Deterministic GUID verified: {guid_v1}")

    # 2. Test Deterministic Deck ID
    topic_1 = "قلب و عروق - نارسایی قلبی"
    deck_name_1 = f"PhysioPath::{topic_1}"
    deck_id_1 = int(hashlib.md5(deck_name_1.encode('utf-8')).hexdigest()[:8], 16)
    deck_id_1_repeat = int(hashlib.md5(deck_name_1.encode('utf-8')).hexdigest()[:8], 16)
    assert deck_id_1 == deck_id_1_repeat, "Deck ID is not deterministic!"

    topic_2 = "نفرولوژی - نارسایی حاد کلیه"
    deck_name_2 = f"PhysioPath::{topic_2}"
    deck_id_2 = int(hashlib.md5(deck_name_2.encode('utf-8')).hexdigest()[:8], 16)
    assert deck_id_1 != deck_id_2, "Different topics must have distinct deck IDs!"
    print(f"[SUCCESS] Deterministic Deck ID verified: {deck_id_1} vs {deck_id_2}")

    # 3. Test building actual .apkg
    test_data = {
        "topic_title": "تست آنکی",
        "anki_flashcards": [
            {
                "text": "در بیمار با سپسیس، خط اول احیای مایعات {{c1::کریستالوئید به میزان 30 mL/kg}} است.",
                "extra": "جهت بازگرداندن پرفیوژن بافتی ظرف ۳ ساعت اول.",
                "tag": "PhysioPath::Infection"
            }
        ]
    }
    test_out = "dist/test_deterministic.apkg"
    os.makedirs("dist", exist_ok=True)
    build_anki_apkg(test_data, test_out)
    assert os.path.exists(test_out), "Failed to produce test .apkg file!"
    print(f"[SUCCESS] Test .apkg generated at: {test_out}")

if __name__ == "__main__":
    test_persian_digit_matching()
    test_anki_deterministic_guid_and_deck_id()
