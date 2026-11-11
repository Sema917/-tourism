"""Общие приёмы вёрстки презентаций комплекта."""

from __future__ import annotations

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

FONT = "Calibri"

NAVY = "16324A"
NAVY_DEEP = "102433"
INK = "1C2830"
CREAM = "F7F4EE"
WHITE = "FFFFFF"
MUTED = "5E6C76"
LINE = "E4DCD0"
CORAL = "E15A3A"
TEAL = "0F7A6C"
AMBER = "A86B12"
AMBER_BG = "F3C15B"
INDIGO = "3A46A0"
SOFT = {
    "E15A3A": "FDE8E2",
    "0F7A6C": "E5F4F1",
    "A86B12": "FFF4D8",
    "3A46A0": "E8EAF8",
}

TYPE_COLORS = {
    "П": "1F7A45",
    "Т": "3E5870",
    "Ч": "C24B35",
    "З": "2A6FBA",
    "Х": "7A4C96",
}

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def font_run(run, size, color, bold=False, italic=False, name=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", name)


def add_shape(slide, x, y, w, h, fill=None, line=None, rounded=False, radius=0.08):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.shadow.inherit = False
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(1.25)
    if rounded:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    spPr = shape._element.spPr
    effect = spPr.find(qn("a:effectLst"))
    if effect is not None:
        spPr.remove(effect)
    return shape


def textbox(slide, x, y, w, h, anchor="top"):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.anchor = ANCHOR[anchor]
    return tf


def write(slide, x, y, w, h, items, anchor="top"):
    """items: list of dicts with t/runs, size, color, bold, italic, align, before, after, spacing."""
    tf = textbox(slide, x, y, w, h, anchor)
    for i, item in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ALIGN[item.get("align", "left")]
        para.space_before = Pt(item.get("before", 0))
        para.space_after = Pt(item.get("after", 0))
        para.line_spacing = item.get("spacing", 1.0)
        runs = item.get("runs")
        if runs is None:
            runs = [item]
        for part in runs:
            run = para.add_run()
            run.text = part.get("t", "")
            font_run(
                run,
                part.get("size", item.get("size", 18)),
                part.get("color", item.get("color", INK)),
                part.get("bold", item.get("bold", False)),
                part.get("italic", item.get("italic", False)),
            )
    return tf


def new_slide(prs, dark=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, 13.333, 7.5, NAVY_DEEP if dark else CREAM)
    if not dark:
        add_shape(slide, 0, 0, 13.333, 0.08, NAVY)
    return slide


def footer(slide, num, total, dark=False):
    color = "C5D0D6" if dark else "7C8A93"
    write(
        slide,
        0.5,
        7.12,
        9.4,
        0.28,
        [{"t": f"Наука 0+    ·    Путь в профессию    ·    комплект 2026.1", "size": 11, "color": color}],
    )
    write(
        slide,
        10.4,
        7.12,
        2.45,
        0.28,
        [{"t": f"{num}   /   {total}", "size": 11, "color": color, "align": "right"}],
    )


def notes(slide, text):
    frame = slide.notes_slide.notes_text_frame
    frame.text = text


def chip(slide, x, y, w, h, text, fill, color, size=14, bold=True):
    add_shape(slide, x, y, w, h, fill, rounded=True, radius=0.3)
    write(
        slide,
        x + 0.12,
        y,
        w - 0.24,
        h,
        [{"t": text, "size": size, "color": color, "bold": bold, "align": "center"}],
        anchor="middle",
    )


def eyebrow(slide, text, color=TEAL):
    write(slide, 0.5, 0.24, 12.3, 0.28, [{"t": text.upper(), "size": 13, "color": color, "bold": True}])


def title(slide, text, y=0.5, h=0.62, size=32):
    write(slide, 0.5, y, 12.3, h, [{"t": text, "size": size, "color": INK, "bold": True, "spacing": 0.95}])
