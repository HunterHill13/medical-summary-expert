# -*- coding: utf-8 -*-
"""
pdf_designer_engine.py
موتور رندرینگ و کادربندی نشریاتی اسناد پزشکی بر پایه معماری دولایه (Dual-Pass Hybrid Architecture)
کالیبره‌شده برای ابعاد استاندارد A4، حذف کامل آرتفکت‌های مرورگر و حل تداخل هدر و تیتر.
"""

import os
import sys
import io
from typing import Optional, List, Dict, Any
import pypdf
from playwright.sync_api import sync_playwright
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

# Force UTF-8 encoding in stdout/stderr
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass


def shape_fa(text: str) -> str:
    """اصلاح جهت‌داری و اتصال حروف فارسی برای ReportLab"""
    if not text:
        return ""
    try:
        return get_display(arabic_reshaper.reshape(str(text)))
    except Exception:
        return str(text)


class MedicalPDFDesigner:
    """
    موتور رندرینگ نشریاتی دولایه:
    پاس ۱: رندر با Playwright و مارجین‌های امن میلی‌متری (سرکوب کامل هدر/فوتر مرورگر)
    پاس ۲: استامپ وکتوری کادر خارجی دور صفحه، هدر موضوعی اختصاصی، خط حائل عنابی، خط فوتر و شماره صفحه دوسویه فارسی
    """

    def __init__(self, output_pdf_path: str, course_title: str = "فیزیوپاتولوژی جامع", institution: str = "دانشگاه بق"):
        self.output_pdf_path = output_pdf_path
        self.course_title = course_title
        self.institution = institution
        self.sections: List[Dict[str, Any]] = []
        self._setup_fonts()

    def _setup_fonts(self):
        self.font_name = 'Helvetica'
        candidates = [
            r"C:\Windows\Fonts\tahoma.ttf",
            r"C:\Windows\Fonts\arial.ttf",
            r"C:\Windows\Fonts\seguiemj.ttf"
        ]
        for c in candidates:
            if os.path.exists(c):
                try:
                    font_alias = f"DesignerFont_{os.path.basename(c).split('.')[0]}"
                    pdfmetrics.registerFont(TTFont(font_alias, c))
                    self.font_name = font_alias
                    break
                except Exception:
                    continue

    def add_section(self, section_id: str, html_content: str, topic_title: Optional[str] = None, is_cover: bool = False):
        self.sections.append({
            "id": section_id,
            "html": html_content,
            "topic": topic_title or "خلاصه مبحث فیزیوپاتولوژی",
            "is_cover": is_cover
        })

    def add_cover_page(self, title: Optional[str] = None, subtitle: str = "مجموعه کامل و جامع درسنامه‌ها و مرور سریع آزمون", faculty: str = "دانشکده پزشکی", academic_year: str = "سال تحصیلی ۱۴۰۴-۱۴۰۵"):
        """تولید و ثبت جلد تمام‌صفحه نشریاتی با گرادینت عنابی-سرمه‌ای و نشان دانشگاه"""
        main_title = title or self.course_title
        cover_html = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8">
<style>
    @page {{
        size: A4 portrait;
        margin: 0;
    }}
    * {{
        box-sizing: border-box;
        margin: 0;
        padding: 0;
    }}
    body {{
        font-family: 'Dubai', 'Vazirmatn', Tahoma, sans-serif;
        background: linear-gradient(135deg, #042F2E 0%, #007A87 45%, #0891B2 100%);
        color: #FFFFFF;
        width: 210mm;
        height: 297mm;
        padding: 20mm;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        direction: rtl;
        text-align: center;
        overflow: hidden;
    }}
    .border-frame {{
        border: 2.5px solid rgba(212, 172, 13, 0.7);
        outline: 1px solid rgba(255, 255, 255, 0.2);
        outline-offset: 4px;
        border-radius: 8px;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        padding: 16mm 12mm;
    }}
    .header-banner {{
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 6px;
    }}
    .uni-emblem {{
        font-size: 20pt;
        letter-spacing: 1px;
        color: #F7DC6F;
        font-weight: bold;
    }}
    .faculty-badge {{
        font-size: 11pt;
        color: #E5E7E9;
        background: rgba(255, 255, 255, 0.12);
        padding: 4px 16px;
        border-radius: 20px;
        display: inline-block;
    }}
    .main-title-box {{
        margin: auto 0;
        padding: 20px 0;
    }}
    .badge-pill {{
        display: inline-block;
        background: #D4AC0D;
        color: #1A252F;
        font-size: 9.5pt;
        font-weight: bold;
        padding: 5px 14px;
        border-radius: 12px;
        margin-bottom: 18px;
    }}
    h1 {{
        font-size: 26pt;
        font-weight: 900;
        line-height: 1.4;
        color: #FFFFFF;
        text-shadow: 0 2px 4px rgba(0,0,0,0.5);
        margin-bottom: 14px;
    }}
    .subtitle {{
        font-size: 12.5pt;
        color: #FADBD8;
        line-height: 1.6;
        max-width: 85%;
        margin: 0 auto;
    }}
    .gold-divider {{
        width: 140px;
        height: 3px;
        background: linear-gradient(90deg, transparent, #F7DC6F, transparent);
        margin: 22px auto;
    }}
    .features-grid {{
        display: flex;
        justify-content: center;
        gap: 15px;
        margin-top: 15px;
    }}
    .feature-tag {{
        background: rgba(255,255,255,0.08);
        border: 1px solid rgba(255,255,255,0.15);
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 8.5pt;
        color: #FADBD8;
    }}
    .footer-meta {{
        border-top: 1px solid rgba(255, 255, 255, 0.2);
        padding-top: 12px;
        font-size: 9pt;
        color: #BDC3C7;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
</style>
</head>
<body>
    <div class="border-frame">
        <div class="header-banner">
            <div class="uni-emblem">{self.institution}</div>
            <div class="faculty-badge">{faculty} • درسنامه جامع و مرور سریع آزمون</div>
        </div>

        <div class="main-title-box">
            <div class="badge-pill">سرفصل جامع فیزیوپاتولوژی بالینی</div>
            <h1>{main_title}</h1>
            <div class="gold-divider"></div>
            <div class="subtitle">{subtitle}</div>
            <div class="features-grid">
                <span class="feature-tag">✔ زنجیره پاتوفیزیولوژی فشرده</span>
                <span class="feature-tag">✔ جداول افتراقی و الگوهای طلایی</span>
                <span class="feature-tag">✔ فلش‌کارت‌های Anki آزمون</span>
            </div>
        </div>

        <div class="footer-meta">
            <span>تهیه شده در: {self.institution}</span>
            <span>{academic_year}</span>
        </div>
    </div>
</body>
</html>
"""
        self.add_section("cover_page", cover_html, topic_title=main_title, is_cover=True)

    def add_table_of_contents(self, chapters: List[Dict[str, Any]], title: str = "فهرست مباحث درسنامه"):
        """تولید و ثبت فهرست فصول و سرفصل‌های درسنامه جامع"""
        items_html = []
        for idx, ch in enumerate(chapters, 1):
            ch_title = ch.get("topic_title") or ch.get("title", f"مبحث {idx}")
            organ = ch.get("organ_system", "")
            pearls_count = len(ch.get("clinical_pearls", [])) or len(ch.get("modules", []))
            badges = []
            if organ:
                badges.append(f"<span class='toc-badge'>{organ}</span>")
            if pearls_count:
                badges.append(f"<span class='toc-subbadge'>{pearls_count} مبحث / نکته کلیدی</span>")
            badge_str = " ".join(badges)

            items_html.append(f"""
            <div class="toc-item">
                <div class="toc-number">فصل {idx}</div>
                <div class="toc-title-box">
                    <div class="toc-title">{ch_title}</div>
                    <div class="toc-meta">{badge_str}</div>
                </div>
            </div>
            """)

        toc_body = "\n".join(items_html)

        toc_html = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8">
<style>
    @page {{
        size: A4 portrait;
        margin: 24mm 17mm 20mm 17mm;
    }}
    * {{
        box-sizing: border-box;
    }}
    body {{
        font-family: 'Dubai', 'Vazirmatn', Tahoma, sans-serif;
        color: #2C3E50;
        direction: rtl;
        line-height: 1.6;
    }}
    .toc-header {{
        border-bottom: 2px solid #007A87;
        padding-bottom: 8px;
        margin-bottom: 22px;
        display: flex;
        justify-content: space-between;
        align-items: baseline;
    }}
    .toc-header h1 {{
        color: #007A87;
        font-size: 17pt;
        margin: 0;
    }}
    .toc-header span {{
        color: #0891B2;
        font-size: 10pt;
    }}
    .toc-container {{
        display: flex;
        flex-direction: column;
        gap: 10px;
    }}
    .toc-item {{
        display: flex;
        align-items: center;
        background: #FDFEFE;
        border: 1px solid #E5E7E9;
        border-right: 4px solid #007A87;
        border-radius: 6px;
        padding: 9px 14px;
        gap: 15px;
    }}
    .toc-number {{
        font-weight: bold;
        color: #007A87;
        font-size: 10.5pt;
        min-width: 65px;
        background: #E0F2FE;
        padding: 3px 8px;
        border-radius: 4px;
        text-align: center;
    }}
    .toc-title-box {{
        flex-grow: 1;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .toc-title {{
        font-size: 10.5pt;
        font-weight: bold;
        color: #1B2631;
    }}
    .toc-badge {{
        background: #E8F8F5;
        color: #117864;
        font-size: 8pt;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
    }}
    .toc-subbadge {{
        background: #FEF9E7;
        color: #B7950B;
        font-size: 8pt;
        padding: 2px 8px;
        border-radius: 4px;
    }}
</style>
</head>
<body>
    <div class="toc-header">
        <h1>{title}</h1>
        <span>{self.institution} • مرور ساختاریافته آزمون</span>
    </div>
    <div class="toc-container">
        {toc_body}
    </div>
</body>
</html>
"""
        self.add_section("table_of_contents", toc_html, topic_title=title, is_cover=False)

    def add_chapter_summary(self, summary_data: Dict[str, Any], chapter_title: Optional[str] = None):
        """تولید و ثبت HTML خلاصه جلسه و ثبت آن به عنوان بخش اختصاصی با هدر مستقل"""
        from build_summary_pack import build_html_summary
        title = chapter_title or summary_data.get("topic_title", "خلاصه مبحث")
        tmp_html_str = build_html_summary(summary_data)
        sec_id = f"chap_{len(self.sections) + 1}"
        self.add_section(sec_id, tmp_html_str, topic_title=title, is_cover=False)

    def build(self, browser_executable: Optional[str] = None, tmp_dir: str = "tmp_pdf_build") -> str:
        os.makedirs(tmp_dir, exist_ok=True)
        raw_pdfs = []
        section_map = []

        # پاس اول: رندر تمیز با مارجین‌های استاندارد (DisplayHeaderFooter = False)
        with sync_playwright() as p:
            browser = None
            # Launch system Edge or Chrome
            if browser_executable and os.path.exists(browser_executable):
                try:
                    browser = p.chromium.launch(executable_path=browser_executable, headless=True)
                except Exception:
                    browser = None

            if not browser:
                for channel in ["msedge", "chrome", None]:
                    try:
                        if channel:
                            browser = p.chromium.launch(channel=channel, headless=True)
                        else:
                            browser = p.chromium.launch(headless=True)
                        break
                    except Exception:
                        continue

            if not browser:
                raise RuntimeError("Could not launch Chromium browser (Edge or Chrome) via Playwright!")

            page = browser.new_page()

            for sec in self.sections:
                sec_html_file = os.path.join(tmp_dir, f"{sec['id']}.html")
                sec_pdf_file = os.path.join(tmp_dir, f"{sec['id']}_raw.pdf")

                with open(sec_html_file, "w", encoding="utf-8") as f:
                    f.write(sec['html'])

                page.goto('file:///' + os.path.abspath(sec_html_file).replace('\\', '/'), wait_until='networkidle')

                if sec['is_cover']:
                    page.pdf(
                        path=sec_pdf_file,
                        format='A4',
                        print_background=True,
                        display_header_footer=False,
                        margin={'top': '0mm', 'bottom': '0mm', 'left': '0mm', 'right': '0mm'}
                    )
                else:
                    page.pdf(
                        path=sec_pdf_file,
                        format='A4',
                        print_background=True,
                        display_header_footer=False,
                        margin={'top': '24mm', 'bottom': '20mm', 'left': '17mm', 'right': '17mm'}
                    )

                r = pypdf.PdfReader(sec_pdf_file)
                raw_pdfs.append(sec_pdf_file)
                section_map.append((len(r.pages), sec['topic'], sec['is_cover']))

            browser.close()

        # ادغام موقت بخش‌ها
        merger = pypdf.PdfWriter()
        for p_file in raw_pdfs:
            r = pypdf.PdfReader(p_file)
            for page_obj in r.pages:
                merger.add_page(page_obj)

        merged_tmp = os.path.join(tmp_dir, "merged_raw.pdf")
        with open(merged_tmp, "wb") as f:
            merger.write(f)

        merged_reader = pypdf.PdfReader(merged_tmp)
        total_pages = len(merged_reader.pages)

        # نگاشت شماره صفحه به موضوع اختصاصی
        page_topics = {}
        curr = 0
        for count, topic, is_cov in section_map:
            for _ in range(count):
                page_topics[curr] = (topic, is_cov)
                curr += 1

        # پاس دوم: استامپ وکتوری کادر، هدر متغیر و شماره صفحه فارسی
        final_writer = pypdf.PdfWriter()
        for idx in range(total_pages):
            base_p = merged_reader.pages[idx]
            topic, is_cov = page_topics[idx]

            # معافیت جلد: اگر صفحه جلد باشد، کادر و هدر اعمال نمی‌شود
            if is_cov:
                final_writer.add_page(base_p)
                continue

            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)

            # ۱. کادر بیرونی دور تا دور صفحه با گوشه‌های ملایم (A4: 595.28 x 841.89 pt)
            # مختصات: x=34pt, y=26pt, w=527.28pt, h=789.89pt, radius=4pt
            can.setStrokeColorRGB(0.7, 0.73, 0.76)  # #B2BABB
            can.setLineWidth(0.85)
            bx, by, bw, bh = 34, 26, A4[0] - 68, A4[1] - 52
            can.roundRect(bx, by, bw, bh, radius=4, stroke=1, fill=0)

            # ۲. خط حائل و محتوای هدر (y = 789.89 pt)
            can.setStrokeColorRGB(0.0, 0.48, 0.53)  # #007A87 فیروزه‌ای رسمی
            can.setLineWidth(1.2)
            hy = A4[1] - 52
            can.line(bx, hy, bx + bw, hy)

            can.setFont(self.font_name, 8.5)
            can.setFillColorRGB(0.0, 0.48, 0.53)
            if topic:
                can.drawRightString(bx + bw - 12, hy + 7, shape_fa(topic))

            can.setFont(self.font_name, 7.5)
            can.setFillColorRGB(0.03, 0.57, 0.70)  # #0891B2
            can.drawString(bx + 12, hy + 7, shape_fa(self.institution))

            # ۳. خط حائل و محتوای فوتر (y = 44.0 pt)
            can.setStrokeColorRGB(0.8, 0.83, 0.86)  # #D5D8DC طوسی ملایم
            can.setLineWidth(0.75)
            fy = 44.0
            can.line(bx, fy, bx + bw, fy)

            can.setFont(self.font_name, 8.5)
            can.setFillColorRGB(0.25, 0.3, 0.35)  # #2E4053
            can.drawCentredString(A4[0] / 2.0, 31, shape_fa(f"صفحه {idx + 1} از {total_pages}"))

            can.setFont(self.font_name, 7.5)
            can.setFillColorRGB(0.5, 0.55, 0.6)
            can.drawRightString(bx + bw - 12, 31, shape_fa("مرور سریع آزمون"))
            can.drawString(bx + 12, 31, shape_fa(self.course_title))

            can.save()
            packet.seek(0)
            stamp_pdf = pypdf.PdfReader(packet)
            base_p.merge_page(stamp_pdf.pages[0])
            final_writer.add_page(base_p)

        os.makedirs(os.path.dirname(os.path.abspath(self.output_pdf_path)), exist_ok=True)
        from build_summary_pack import safe_write_pdf
        written_path = safe_write_pdf(self.output_pdf_path, lambda f: final_writer.write(f))

        print(f"[OK] Publication-Grade PDF Generated: {written_path} ({total_pages} pages)")
        return written_path
