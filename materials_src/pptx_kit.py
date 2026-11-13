"""Единая вёрстка презентации: светлая гамма, воздух, читаемый текст."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from lxml import etree
from PIL import Image, ImageEnhance, ImageOps
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SANS = "Carlito"
DISPLAY = "Noto Sans Display"
ASSETS = Path(__file__).resolve().parent / "assets"

INK = "1A2330"
MUTED = "667482"
SOFT = "F3F6F8"
WHITE = "FFFFFF"
TEAL = "0C8B7A"
CORAL = "E15A45"
SKY = "2F5D9A"
AMBER = "C47A1A"
LINE = "D7DEE5"
PANEL = "FFFFFF"

TYPE_COLORS = {
    "П": "1F8A4C",
    "Т": "3E5A70",
    "Ч": "D24B38",
    "З": "2A6FBA",
    "Х": "6B4C9A",
}

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
_CACHE: dict[tuple, bytes] = {}
W, H = 13.333, 7.5


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


def add_shape(slide, x, y, w, h, fill=None, line=None, rounded=False, radius=0.1, line_pt=1.0):
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
        para.line_spacing = item.get("spacing", 1.08)
        runs = item.get("runs") or [item]
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


def prepare_image(name: str, *, width=1920, height=1080, darken=0.0, crop="cover") -> BytesIO:
    key = (name, width, height, darken, crop)
    if key not in _CACHE:
        img = Image.open(ASSETS / name).convert("RGB")
        if crop == "cover":
            img = ImageOps.fit(img, (width, height), method=Image.Resampling.LANCZOS)
        else:
            img = img.resize((width, height), Image.Resampling.LANCZOS)
        if darken:
            img = ImageEnhance.Brightness(img).enhance(1.0 - darken)
            img = ImageEnhance.Contrast(img).enhance(1.04)
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=90, optimize=True)
        _CACHE[key] = buf.getvalue()
    out = BytesIO(_CACHE[key])
    out.seek(0)
    return out


def picture(slide, name, x, y, w, h, **kwargs):
    stream = prepare_image(name, width=max(int(w * 220), 40), height=max(int(h * 220), 40), **kwargs)
    return slide.shapes.add_picture(stream, Inches(x), Inches(y), Inches(w), Inches(h))


def sticker(slide, name: str, x, y, size=1.15, rotate=0):
    key = ("sticker", name, size, rotate)
    if key not in _CACHE:
        img = Image.open(ASSETS / name).convert("RGBA")
        if rotate:
            img = img.rotate(rotate, expand=True, resample=Image.Resampling.BICUBIC)
        px = max(int(size * 220), 64)
        img = ImageOps.contain(img, (px, px), method=Image.Resampling.LANCZOS)
        buf = BytesIO()
        img.save(buf, format="PNG", optimize=True)
        _CACHE[key] = buf.getvalue()
    out = BytesIO(_CACHE[key])
    out.seek(0)
    return slide.shapes.add_picture(out, Inches(x), Inches(y), Inches(size), Inches(size))


def new_blank(prs, fill=SOFT):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, 0, 0, W, H, fill)
    return slide


def rail(slide, color=TEAL):
    add_shape(slide, 0, 0, 0.12, H, color)


def chip(slide, x, y, w, h, text, fill, color=WHITE, size=13):
    add_shape(slide, x, y, w, h, fill, rounded=True, radius=0.22)
    write(slide, x + 0.08, y, w - 0.16, h, [{"t": text, "size": size, "color": color, "bold": True, "align": "center"}], anchor="middle")


def option_row(slide, x, y, w, h, letter, text, accent=TEAL):
    add_shape(slide, x, y, w, h, WHITE, line=LINE, rounded=True, radius=0.12, line_pt=1.1)
    add_shape(slide, x + 0.28, y + (h - 0.58) / 2, 0.58, 0.58, accent, rounded=True, radius=0.5)
    write(
        slide,
        x + 0.28,
        y + (h - 0.58) / 2,
        0.58,
        0.58,
        [{"t": letter, "size": 18, "color": WHITE, "bold": True, "align": "center"}],
        anchor="middle",
    )
    write(slide, x + 1.1, y, w - 1.4, h, [{"t": text, "size": 20, "color": INK}], anchor="middle")


def eyebrow(slide, text, color=TEAL, x=0.55, y=0.35):
    write(slide, x, y, 12, 0.28, [{"t": text.upper(), "size": 12, "color": color, "bold": True}])


def title(slide, text, y=0.65, size=34, color=INK, w=12.0, x=0.55):
    write(slide, x, y, w, 0.7, [{"t": text, "size": size, "color": color, "bold": True, "spacing": 0.96, "font": DISPLAY}])


def footer(slide, num, total, light=True):
    color = "8A97A3" if light else "9AABBA"
    write(slide, 0.55, 7.1, 9.2, 0.28, [{"t": "Наука 0+    ·    Путь в профессию", "size": 11, "color": color}])
    write(slide, 10.4, 7.1, 2.4, 0.28, [{"t": f"{num}  /  {total}", "size": 11, "color": color, "align": "right"}])


def finish(prs, dark_pages):
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, start=1):
        footer(slide, i, total, light=(i not in dark_pages))
    return prs
