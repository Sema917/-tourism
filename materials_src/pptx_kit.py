"""Вёрстка презентаций в стиле квиза: бордо / крем / золото, рамка, блоки."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from lxml import etree
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SANS = "Carlito"
SERIF = "Liberation Serif"
ASSETS = Path(__file__).resolve().parent / "assets"

# Палитра по мотивам квиза «Народы и традиции России»
BURGUNDY = "6B1E2A"
BURGUNDY_DEEP = "4E1420"
CREAM = "F6F0E6"
CREAM_SOFT = "EFE6D8"
INK = "1C1A18"
MUTED = "6A5E55"
GOLD = "C4A35A"
GOLD_SOFT = "E8D7A8"
WHITE = "FFFFFF"
CORAL = "B33A3A"
TEAL = "0F6B5C"
AMBER_D = "A86B12"
INDIGO = "3A46A0"

# совместимость со старыми именами
NAVY = BURGUNDY
NAVY2 = BURGUNDY_DEEP
CORAL_OLD = CORAL
AMBER = GOLD
SOFT_TEAL = "E5F4F1"
SOFT_CORAL = "FDE8E2"
SOFT_AMBER = "FFF4D8"
SOFT_INDIGO = "E8EAF8"

TYPE_COLORS = {
    "П": "1F7A45",
    "Т": "3E5870",
    "Ч": "A33A2C",
    "З": "2A6FBA",
    "Х": "7A4C96",
}

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}

_CACHE: dict[tuple, bytes] = {}
W, H = 13.333, 7.5
FRAME = 0.28


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def font_run(run, size, color, bold=False, italic=False, name=SANS):
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


def add_shape(slide, x, y, w, h, fill=None, line=None, rounded=False, radius=0.08, line_pt=1.25):
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
        shape.line.width = Pt(line_pt)
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
                part.get("font", item.get("font", SANS)),
            )
    return tf


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def prepare_image(name: str, *, width=1920, height=1080, darken=0.0, blur=0, crop="cover") -> BytesIO:
    key = (name, width, height, darken, blur, crop)
    if key not in _CACHE:
        path = ASSETS / name
        img = Image.open(path).convert("RGB")
        if crop == "cover":
            img = ImageOps.fit(img, (width, height), method=Image.Resampling.LANCZOS)
        else:
            img = img.resize((width, height), Image.Resampling.LANCZOS)
        if blur:
            img = img.filter(ImageFilter.GaussianBlur(blur))
        if darken:
            img = ImageEnhance.Brightness(img).enhance(1.0 - darken)
            img = ImageEnhance.Contrast(img).enhance(1.05)
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=88, optimize=True)
        _CACHE[key] = buf.getvalue()
    out = BytesIO(_CACHE[key])
    out.seek(0)
    return out


def picture(slide, name, x, y, w, h, **kwargs):
    stream = prepare_image(name, width=max(int(w * 180), 40), height=max(int(h * 180), 40), **kwargs)
    return slide.shapes.add_picture(stream, Inches(x), Inches(y), Inches(w), Inches(h))


def new_blank(prs, fill=CREAM):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, W, H, fill)
    return slide


def frame(slide, color=BURGUNDY, inset=FRAME, pt=1.35):
    """Тонкая рамка по периметру, как в квизе."""
    add_shape(slide, inset, inset, W - 2 * inset, H - 2 * inset, fill=None, line=color, line_pt=pt)


def hairline(slide, y, color=BURGUNDY, inset=FRAME + 0.08):
    add_shape(slide, inset, y, W - 2 * inset, 0.015, color)


def footer(slide, num, total, light=True, label="Путь в профессию"):
    color = GOLD_SOFT if not light else MUTED
    write(slide, FRAME + 0.12, 7.08, 9.2, 0.28, [{"t": f"Наука 0+    ·    {label}", "size": 11, "color": color}])
    write(
        slide,
        10.4,
        7.08,
        2.5,
        0.28,
        [{"t": str(num), "size": 12, "color": color, "align": "right", "font": SERIF}],
    )


def chip(slide, x, y, w, h, text, fill, color=WHITE, size=13):
    add_shape(slide, x, y, w, h, fill, rounded=True, radius=0.28)
    write(
        slide,
        x + 0.08,
        y,
        w - 0.16,
        h,
        [{"t": text, "size": size, "color": color, "bold": True, "align": "center"}],
        anchor="middle",
    )


def option_row(slide, x, y, w, h, letter, text, accent=BURGUNDY):
    """Кнопка-вариант А/Б в стиле квиза."""
    add_shape(slide, x, y, w, h, WHITE, line=GOLD_SOFT, rounded=True, radius=0.18, line_pt=1.4)
    add_shape(slide, x + 0.18, y + (h - 0.42) / 2, 0.42, 0.42, accent, rounded=True, radius=0.5)
    write(
        slide,
        x + 0.18,
        y + (h - 0.42) / 2,
        0.42,
        0.42,
        [{"t": letter, "size": 14, "color": WHITE, "bold": True, "align": "center"}],
        anchor="middle",
    )
    write(slide, x + 0.78, y, w - 0.95, h, [{"t": text, "size": 15, "color": INK, "font": SERIF}], anchor="middle")


def eyebrow(slide, text, color=BURGUNDY, x=0.55, y=0.42):
    write(slide, x, y, 12, 0.28, [{"t": text.upper(), "size": 12, "color": color, "bold": True}])


def title(slide, text, y=0.7, size=34, color=INK, w=12.0, x=0.55, font=SERIF):
    write(slide, x, y, w, 0.85, [{"t": text, "size": size, "color": color, "bold": True, "spacing": 0.95, "font": font}])


def finish(prs, dark_pages):
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, start=1):
        footer(slide, i, total, light=(i not in dark_pages))
    return prs
