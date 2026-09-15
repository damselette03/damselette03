#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate detailed Traditional Chinese notes for Chapter 2 PPT slides."""

from __future__ import annotations

import json
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    PageBreak,
    HRFlowable,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

ROOT = Path(__file__).resolve().parent
FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
pdfmetrics.registerFont(TTFont("WQYmh", FONT_PATH, subfontIndex=0))

PRIMARY = HexColor("#1a365d")
ACCENT = HexColor("#2b6cb0")
LIGHT_BG = HexColor("#edf2f7")
MUTED = HexColor("#4a5568")
BORDER = HexColor("#cbd5e0")


def esc(text: str) -> str:
    """Escape XML-sensitive characters for ReportLab Paragraph."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def make_styles():
    return {
        "cover_title": ParagraphStyle(
            "cover_title",
            fontName="WQYmh",
            fontSize=17,
            leading=24,
            textColor=PRIMARY,
            alignment=TA_LEFT,
            spaceAfter=4,
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub",
            fontName="WQYmh",
            fontSize=10,
            leading=14,
            textColor=ACCENT,
            spaceAfter=8,
        ),
        "intro": ParagraphStyle(
            "intro",
            fontName="WQYmh",
            fontSize=10,
            leading=15,
            textColor=HexColor("#2d3748"),
            alignment=TA_JUSTIFY,
            spaceAfter=8,
        ),
        "h2": ParagraphStyle(
            "h2",
            fontName="WQYmh",
            fontSize=11.5,
            leading=16,
            textColor=PRIMARY,
            spaceBefore=6,
            spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "body",
            fontName="WQYmh",
            fontSize=9.5,
            leading=14,
            textColor=black,
            alignment=TA_JUSTIFY,
            spaceAfter=5,
        ),
        "term_name": ParagraphStyle(
            "term_name",
            fontName="WQYmh",
            fontSize=9.5,
            leading=13,
            textColor=ACCENT,
            spaceBefore=3,
            spaceAfter=1,
        ),
        "term_def": ParagraphStyle(
            "term_def",
            fontName="WQYmh",
            fontSize=9,
            leading=13,
            textColor=HexColor("#1a202c"),
            alignment=TA_JUSTIFY,
            spaceAfter=3,
            leftIndent=8,
        ),
    }


def add_page_decoration(canvas, doc):
    canvas.saveState()
    width, height = A4
    canvas.setFillColor(PRIMARY)
    canvas.rect(0, height - 18, width, 18, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont("WQYmh", 8)
    canvas.drawString(
        1.5 * cm,
        height - 12,
        "總體經濟學 Chapter 2｜國民所得的衡量｜詳細講解講義",
    )
    canvas.setFillColor(LIGHT_BG)
    canvas.rect(0, 0, width, 16, fill=1, stroke=0)
    canvas.setFillColor(MUTED)
    canvas.setFont("WQYmh", 8)
    total = getattr(doc, "total_pages_hint", 26)
    canvas.drawCentredString(
        width / 2,
        5,
        f"第 {doc.page} / {total} 頁（含封面）｜請與原 PPT 同步對照閱讀",
    )
    canvas.restoreState()


def build_pdf(output_path: Path, pages: list):
    styles = make_styles()
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=1.6 * cm,
        rightMargin=1.6 * cm,
        topMargin=1.4 * cm,
        bottomMargin=1.3 * cm,
        title="總體經濟學 Chapter2 國民所得的衡量｜詳細講解講義",
        author="依謝德宗《總體經濟學》Chapter2 投影片整理",
    )
    doc.total_pages_hint = len(pages) + 1

    story = []
    story.append(Spacer(1, 2.0 * cm))
    story.append(Paragraph("總體經濟學 Macroeconomics", styles["cover_sub"]))
    story.append(Paragraph("Chapter 2　國民所得的衡量", styles["cover_title"]))
    story.append(Paragraph("詳細講解講義（逐頁對應原投影片）", styles["cover_title"]))
    story.append(Spacer(1, 0.25 * cm))
    story.append(
        HRFlowable(width="100%", thickness=2, color=ACCENT, spaceBefore=2, spaceAfter=10)
    )
    story.append(
        Paragraph(
            esc(
                "原投影片（謝德宗著《總體經濟學》Chapter2）條列較簡、部分圖頁缺少文字說明。"
                "本講義依原檔 25 頁順序編排：每一頁對應原 PPT 同一頁，補充完整脈絡、公式意涵與專業名詞定義，"
                "方便自修、複習與對照上課投影片使用。"
            ),
            styles["intro"],
        )
    )
    story.append(Paragraph("使用方式", styles["h2"]))
    story.append(
        Paragraph(
            esc(
                "請將本 PDF 與原投影片同步翻頁：看到原 PPT 第 N 頁時，閱讀本講義「對應 PPT 第 N 頁」。"
                "每頁包含：頁面主題說明、內容補完、專業名詞詳解。最後一頁補上原檔缺失的全章收束與自我檢核。"
            ),
            styles["body"],
        )
    )
    story.append(Paragraph("內容範圍", styles["h2"]))
    story.append(
        Paragraph(
            esc(
                "2.1 國民會計帳（生產／支出／所得法）→ 2.2 物價指數 → "
                "2.3 社會福祉與綠色帳 → 2.4 國民財富帳。"
            ),
            styles["body"],
        )
    )
    story.append(PageBreak())

    for page in pages:
        story.append(
            Paragraph(
                esc(f"原投影片對照｜第 {page['slide']} / 25 頁"),
                styles["cover_sub"],
            )
        )
        story.append(Paragraph(esc(page["title"]), styles["cover_title"]))
        story.append(Paragraph(esc(page["subtitle"]), styles["cover_sub"]))
        story.append(
            HRFlowable(width="100%", thickness=1, color=BORDER, spaceBefore=0, spaceAfter=6)
        )
        story.append(Paragraph(esc(page["intro"]), styles["intro"]))

        for sec in page["sections"]:
            story.append(Paragraph(esc(sec["heading"]), styles["h2"]))
            for para in sec.get("body", []):
                story.append(Paragraph("• " + esc(para), styles["body"]))
            for name, definition in sec.get("terms", []):
                story.append(Paragraph(esc(f"【{name}】"), styles["term_name"]))
                story.append(Paragraph(esc(definition), styles["term_def"]))

        if page["slide"] < len(pages):
            story.append(PageBreak())

    doc.build(story, onFirstPage=add_page_decoration, onLaterPages=add_page_decoration)
    print(f"Wrote: {output_path}")


def main():
    pages = json.loads((ROOT / "chapter2_pages.json").read_text(encoding="utf-8"))
    out = ROOT / "Chapter2_國民所得的衡量_詳細講解講義.pdf"
    build_pdf(out, pages)


if __name__ == "__main__":
    main()
