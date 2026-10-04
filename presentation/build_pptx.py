#!/usr/bin/env python3
"""Build an editable 16:9 PowerPoint for the Ohtani media-bias talk."""

from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
OUT = ROOT / "Biased-Sports-News-Ohtani.pptx"

# Shared palette — keep every slide on these five colors.
PAPER = RGBColor(0xF6, 0xF3, 0xEC)
INK = RGBColor(0x1A, 0x23, 0x32)
NAVY = RGBColor(0x16, 0x32, 0x4F)
RED = RGBColor(0x9E, 0x2A, 0x2B)
MUTED = RGBColor(0x6B, 0x6A, 0x64)
LINE = RGBColor(0xD8, 0xD2, 0xC4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xEE, 0xE8, 0xDC)

FONT = "Calibri"
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def _set_solid_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _no_line(shape):
    shape.line.fill.background()


def rect(slide, l, t, w, h, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    _no_line(shape)
    return shape


def textbox(
    slide,
    l,
    t,
    w,
    h,
    lines,
    *,
    size=20,
    bold=False,
    italic=False,
    color=INK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font=FONT,
):
    """lines: str or list of (text, kwargs) / str."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.anchor = anchor

    if isinstance(lines, str):
        lines = [lines]

    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(0)
        p.space_before = Pt(0)
        if isinstance(item, str):
            runs = [(item, {})]
        elif isinstance(item, tuple) and item and isinstance(item[0], str) and isinstance(item[1], dict):
            runs = [item]
        else:
            runs = item
        for text, kw in runs:
            run = p.add_run()
            run.text = text
            run.font.name = kw.get("font", font)
            run.font.size = Pt(kw.get("size", size))
            run.font.bold = kw.get("bold", bold)
            run.font.italic = kw.get("italic", italic)
            run.font.color.rgb = kw.get("color", color)
    return box


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def picture(slide, path, l, t, w, h):
    return slide.shapes.add_picture(str(path), l, t, w, h)


def make_icons():
    ASSETS.mkdir(parents=True, exist_ok=True)
    specs = {
        "icon-word.png": _draw_word,
        "icon-paraphrase.png": _draw_paraphrase,
        "icon-clock.png": _draw_clock,
    }
    for name, fn in specs.items():
        img = Image.new("RGBA", (320, 320), (0, 0, 0, 0))
        fn(ImageDraw.Draw(img))
        img.save(ASSETS / name)


def _draw_word(d):
    d.rounded_rectangle((40, 56, 280, 264), radius=28, outline=(22, 50, 79, 255), width=10)
    d.line((88, 118, 232, 118), fill=(158, 42, 43, 255), width=10)
    d.line((88, 160, 200, 160), fill=(22, 50, 79, 255), width=10)
    d.line((88, 202, 176, 202), fill=(22, 50, 79, 255), width=10)


def _draw_paraphrase(d):
    d.rounded_rectangle((52, 72, 210, 210), radius=20, outline=(22, 50, 79, 255), width=10)
    d.rounded_rectangle((110, 118, 268, 256), radius=20, outline=(158, 42, 43, 255), width=10)


def _draw_clock(d):
    d.ellipse((48, 48, 272, 272), outline=(22, 50, 79, 255), width=10)
    d.line((160, 160, 160, 92), fill=(22, 50, 79, 255), width=12)
    d.line((160, 160, 214, 186), fill=(158, 42, 43, 255), width=12)
    d.ellipse((148, 148, 172, 172), fill=(22, 50, 79, 255))


def footer(slide, page, total=7):
    textbox(slide, Inches(0.7), Inches(7.12), Inches(8), Inches(0.28),
            "Biased Sports News  ·  Shohei Ohtani", size=11, color=MUTED)
    textbox(slide, Inches(11.6), Inches(7.12), Inches(1.1), Inches(0.28),
            f"{page} / {total}", size=11, color=MUTED, align=PP_ALIGN.RIGHT)


def kicker(slide, text, l=Inches(0.7), t=Inches(0.38)):
    textbox(slide, l, t, Inches(10), Inches(0.32), text.upper(),
            size=12, bold=True, color=RED)


def title(slide, text, l=Inches(0.7), t=Inches(0.68), w=Inches(12)):
    textbox(slide, l, t, w, Inches(0.7), text, size=32, bold=True, color=NAVY)


def build():
    make_icons()
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    # --- 1 Title ---
    s = prs.slides.add_slide(blank)
    _set_solid_bg(s, PAPER)
    rect(s, Inches(0), Inches(0), Inches(0.16), SLIDE_H, NAVY)
    kicker(s, "English talk  ·  5 minutes")
    textbox(s, Inches(0.7), Inches(1.7), Inches(7.4), Inches(2.2),
            "Biased Sports News", size=48, bold=True, color=NAVY)
    textbox(s, Inches(0.7), Inches(3.85), Inches(7.4), Inches(1.3),
            "One New York Post column about Shohei Ohtani —\nand how a writer hides his opinion inside a “story.”",
            size=20, color=INK)
    textbox(s, Inches(0.7), Inches(5.4), Inches(7.4), Inches(1.0),
            "Jon Heyman  ·  New York Post  ·  21 March 2024",
            size=16, color=MUTED)
    photo = picture(s, ASSETS / "ohtani-portrait.jpg", Inches(8.85), Inches(1.35), Inches(3.7), Inches(5.55))
    photo.name = "Photo_Ohtani_portrait"
    textbox(s, Inches(8.85), Inches(6.95), Inches(3.7), Inches(0.28),
            "Photo: All-Pro Reels / Wikimedia  CC BY-SA 2.0",
            size=9, color=MUTED, align=PP_ALIGN.RIGHT)
    notes(s, (
        "SAY:\n"
        "Today I will talk about media bias in sports news. Bias is not always a lie. "
        "Sometimes the writer uses other people’s “story” to push his own opinion. "
        "My example is a New York Post column about Shohei Ohtani.\n\n"
        "中文提示：40 秒。點題即可，不要開始講案情。"
    ))

    # --- 2 Facts ---
    s = prs.slides.add_slide(blank)
    _set_solid_bg(s, PAPER)
    rect(s, Inches(0), Inches(0), Inches(0.16), SLIDE_H, NAVY)
    kicker(s, "Before the opinion")
    title(s, "What we know")
    cards = [
        ("01", "The money", "Millions of dollars left\nOhtani’s bank account."),
        ("02", "The firing", "The Dodgers fired his\ninterpreter, Ippei Mizuhara."),
        ("03", "The claim", "Ohtani’s lawyers said he was\nthe “victim of massive theft.”"),
    ]
    for i, (num, head, body) in enumerate(cards):
        x = Inches(0.7 + i * 4.1)
        card = rect(s, x, Inches(2.15), Inches(3.85), Inches(3.7), WHITE)
        card.name = f"Card_{i+1}_background"
        bar = rect(s, x, Inches(2.15), Inches(3.85), Inches(0.08), NAVY)
        bar.name = f"Card_{i+1}_bar"
        textbox(s, x + Inches(0.28), Inches(2.45), Inches(3.3), Inches(0.45),
                num, size=18, bold=True, color=RED)
        textbox(s, x + Inches(0.28), Inches(3.0), Inches(3.3), Inches(0.55),
                head, size=24, bold=True, color=NAVY)
        textbox(s, x + Inches(0.28), Inches(3.7), Inches(3.3), Inches(1.7),
                body, size=18, color=INK)
    footer(s, 2)
    notes(s, (
        "SAY:\n"
        "First, the facts. In March 2024, millions of dollars left Ohtani’s account. "
        "The Dodgers fired his interpreter, Ippei Mizuhara. Ohtani’s lawyers said he was "
        "the victim of theft. At that time, the investigation was not finished.\n\n"
        "中文提示：40 秒。只講事實，先不要罵報紙。"
    ))

    # --- 3 Headline ---
    s = prs.slides.add_slide(blank)
    _set_solid_bg(s, PAPER)
    rect(s, Inches(0), Inches(0), Inches(0.16), SLIDE_H, NAVY)
    kicker(s, "Bias 1  ·  word choice")
    title(s, "The headline already chooses a side")
    quote_bg = rect(s, Inches(0.7), Inches(1.85), Inches(11.95), Inches(2.7), WHITE)
    quote_bg.name = "Headline_card"
    rect(s, Inches(0.7), Inches(1.85), Inches(0.1), Inches(2.7), RED).name = "Headline_accent"
    textbox(s, Inches(1.15), Inches(2.35), Inches(11.1), Inches(1.7),
            [
                [("“…a baseball legend and a ", {"size": 28, "color": INK}),
                 ("financial dimwit", {"size": 28, "bold": True, "color": RED}),
                 ("”", {"size": 28, "color": INK})],
            ],
            size=28, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, Inches(0.7), Inches(4.85), Inches(11.95), Inches(1.5),
            [
                [("dimwit", {"size": 22, "bold": True, "color": RED}),
                 ("  =  a stupid person", {"size": 22, "color": INK})],
                [("", {})],
                [("The writer does not say “questions about the money.” He judges Ohtani first.",
                  {"size": 20, "color": MUTED})],
            ],
            size=20)
    footer(s, 3)
    notes(s, (
        "SAY:\n"
        "The headline already chooses a side. It does not say “questions about the money.” "
        "It calls Ohtani a “financial dimwit.” That word means a stupid person about money. "
        "This is word-choice bias. The writer judges Ohtani before the facts are clear.\n\n"
        "中文提示：50 秒。指紅字 dimwit。"
    ))

    # --- 4 Close reading ---
    s = prs.slides.add_slide(blank)
    _set_solid_bg(s, PAPER)
    rect(s, Inches(0), Inches(0), Inches(0.16), SLIDE_H, NAVY)
    kicker(s, "Bias 2  ·  loaded paraphrase")
    title(s, "It looks like reporting. It is not.")

    box_a = rect(s, Inches(0.7), Inches(1.75), Inches(11.95), Inches(1.55), WHITE)
    box_a.name = "Quote_reported"
    rect(s, Inches(0.7), Inches(1.75), Inches(0.1), Inches(1.55), NAVY).name = "Quote_reported_bar"
    textbox(s, Inches(1.1), Inches(1.88), Inches(11.2), Inches(0.3),
            "CLOSER TO A REPORT", size=11, bold=True, color=NAVY)
    textbox(s, Inches(1.1), Inches(2.2), Inches(11.2), Inches(0.9),
            "“The story of the Ohtani camp is that the interpreter fooled Ohtani for years…”",
            size=20, italic=True, color=INK)

    box_b = rect(s, Inches(0.7), Inches(3.5), Inches(11.95), Inches(1.85), WHITE)
    box_b.name = "Quote_added"
    rect(s, Inches(0.7), Inches(3.5), Inches(0.1), Inches(1.85), RED).name = "Quote_added_bar"
    textbox(s, Inches(1.1), Inches(3.63), Inches(11.2), Inches(0.3),
            "ADDED BY THE WRITER", size=11, bold=True, color=RED)
    textbox(s, Inches(1.1), Inches(3.95), Inches(11.2), Inches(1.2),
            [
                [("“…and ", {"size": 20, "italic": True, "color": INK}),
                 ("presumably", {"size": 20, "bold": True, "italic": True, "color": RED}),
                 (" that Ohtani is ", {"size": 20, "italic": True, "color": INK}),
                 ("gullible and naïve", {"size": 20, "bold": True, "italic": True, "color": RED}),
                 (", plus has ", {"size": 20, "italic": True, "color": INK}),
                 ("terrible taste in best friends", {"size": 20, "bold": True, "italic": True, "color": RED}),
                 (".”", {"size": 20, "italic": True, "color": INK})],
            ])

    textbox(s, Inches(0.7), Inches(5.55), Inches(11.95), Inches(1.1),
            "Who is “presumably”? The writer — not the lawyers.\n"
            "The lawyers never used these words. He puts his guess inside their “story.”",
            size=18, color=INK)
    footer(s, 4)
    notes(s, (
        "SAY:\n"
        "This paragraph looks like reporting. He says “the story of the Ohtani camp.” "
        "But look at the next words. The lawyers never said Ohtani is gullible or naïve. "
        "They never said he has terrible taste in friends. The writer adds these words. "
        "The key word is “presumably.” It means “I guess.” So the person who is guessing "
        "is the writer, not Ohtani’s team. He uses their story as a shell, and puts his opinion inside.\n\n"
        "中文提示：80 秒。這張最重要。講慢，指紅字。"
    ))

    # --- 5 Later facts ---
    s = prs.slides.add_slide(blank)
    _set_solid_bg(s, PAPER)
    rect(s, Inches(0), Inches(0), Inches(0.16), SLIDE_H, NAVY)
    kicker(s, "After the column")
    title(s, "The later story was different", w=Inches(8.2))
    later = [
        ("April 2024", "Mizuhara was charged. Prosecutors said about $16 million was taken."),
        ("The method", "They said he pretended to be Ohtani when he talked to the bank."),
        ("MLB", "Ohtani was not involved in sports betting."),
    ]
    for i, (head, body) in enumerate(later):
        y = Inches(1.9 + i * 1.45)
        row = rect(s, Inches(0.7), y, Inches(8.15), Inches(1.28), WHITE)
        row.name = f"Later_row_{i+1}"
        textbox(s, Inches(0.98), y + Inches(0.18), Inches(7.6), Inches(0.4),
                head, size=16, bold=True, color=RED)
        textbox(s, Inches(0.98), y + Inches(0.58), Inches(7.6), Inches(0.55),
                body, size=18, color=INK)
    photo = picture(s, ASSETS / "ohtani-embassy.jpg", Inches(9.15), Inches(1.9), Inches(3.4), Inches(4.53))
    photo.name = "Photo_Ohtani_embassy"
    textbox(s, Inches(9.15), Inches(6.5), Inches(3.4), Inches(0.4),
            "Photo: U.S. Embassy Tokyo  ·  public domain",
            size=9, color=MUTED, align=PP_ALIGN.RIGHT)
    footer(s, 5)
    notes(s, (
        "SAY:\n"
        "Later, federal prosecutors charged Mizuhara. They said he stole about 16 million dollars "
        "and even pretended to be Ohtani when he talked to the bank. MLB said Ohtani was not "
        "involved in betting. So the early “dimwit” story was not careful. It was a judgment, not news.\n\n"
        "中文提示：50 秒。三點即可。"
    ))

    # --- 6 Three types ---
    s = prs.slides.add_slide(blank)
    _set_solid_bg(s, PAPER)
    rect(s, Inches(0), Inches(0), Inches(0.16), SLIDE_H, NAVY)
    kicker(s, "In one slide")
    title(s, "Three kinds of bias")
    cols = [
        ("icon-word.png", "Word choice", "Emotional labels\nlike dimwit and\ngullible."),
        ("icon-paraphrase.png", "Loaded paraphrase", "A fake summary.\nHe retells their story\nwith his words."),
        ("icon-clock.png", "Too early", "He judges Ohtani’s\ncharacter before the\ninvestigation ends."),
    ]
    for i, (icon, head, body) in enumerate(cols):
        x = Inches(0.7 + i * 4.1)
        card = rect(s, x, Inches(1.95), Inches(3.85), Inches(4.15), WHITE)
        card.name = f"Bias_card_{i+1}"
        pic = picture(s, ASSETS / icon, x + Inches(0.35), Inches(2.2), Inches(0.8), Inches(0.8))
        pic.name = f"Bias_icon_{i+1}"
        textbox(s, x + Inches(0.3), Inches(3.2), Inches(3.25), Inches(0.8),
                head, size=22, bold=True, color=NAVY)
        textbox(s, x + Inches(0.3), Inches(4.1), Inches(3.25), Inches(1.5),
                body, size=18, color=INK)
    footer(s, 6)
    notes(s, (
        "SAY:\n"
        "So I see three kinds of bias. First, emotional words. Second, a fake summary — "
        "he pretends to retell their story. Third, he decides the character of Ohtani too early.\n\n"
        "中文提示：30 秒。三欄掃過即可。"
    ))

    # --- 7 Close ---
    s = prs.slides.add_slide(blank)
    _set_solid_bg(s, PAPER)
    rect(s, Inches(0), Inches(0), Inches(0.16), SLIDE_H, NAVY)
    kicker(s, "Takeaway")
    textbox(s, Inches(0.7), Inches(2.15), Inches(12), Inches(2.0),
            "Same facts.\nDifferent job: inform, or persuade.",
            size=40, bold=True, color=NAVY)
    textbox(s, Inches(0.7), Inches(4.5), Inches(12), Inches(1.2),
            "A good reader asks: Is this a quote — or the writer’s guess?\nThank you.",
            size=22, color=INK)
    footer(s, 7)
    notes(s, (
        "SAY:\n"
        "Good readers ask: Is this a quote, or is this the writer’s guess? Thank you.\n\n"
        "中文提示：10 秒。念完停住，不要加新例子。\n\n"
        "若被問「專欄本來就可以有意見」：\n"
        "Yes. But many people read it like news. And he writes “the story of the Ohtani camp,” "
        "so it looks like reporting.\n\n"
        "若被問「這是假新聞嗎」：\n"
        "He does not invent the bank transfers. Bias here is the extra judgment."
    ))

    prs.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build()
