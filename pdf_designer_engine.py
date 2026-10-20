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

    def __init__(self, output_pdf_path: str, course_title: str = "فیزیوپاتولوژی • دانشکده پزشکی"):
        self.output_pdf_path = output_pdf_path
        self.course_title = course_title
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
            can.setStrokeColorRGB(0.47, 0.16, 0.12)  # #78281F عنابی رسمی
            can.setLineWidth(1.2)
            hy = A4[1] - 52
            can.line(bx, hy, bx + bw, hy)

            can.setFont(self.font_name, 8.5)
            can.setFillColorRGB(0.47, 0.16, 0.12)
            if topic:
                can.drawRightString(bx + bw - 12, hy + 7, shape_fa(topic))

            can.setFont(self.font_name, 7.5)
            can.setFillColorRGB(0.1, 0.32, 0.46)  # #1A5276 سرمه‌ای متای دوره
            can.drawString(bx + 12, hy + 7, shape_fa(self.course_title))

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
            can.drawString(bx + 12, 31, shape_fa("دانشکده پزشکی"))

            can.save()
            packet.seek(0)
            stamp_pdf = pypdf.PdfReader(packet)
            base_p.merge_page(stamp_pdf.pages[0])
            final_writer.add_page(base_p)

        os.makedirs(os.path.dirname(os.path.abspath(self.output_pdf_path)), exist_ok=True)
        with open(self.output_pdf_path, "wb") as f:
            final_writer.write(f)

        print(f"[OK] Publication-Grade PDF Generated: {self.output_pdf_path} ({total_pages} pages)")
        return self.output_pdf_path
