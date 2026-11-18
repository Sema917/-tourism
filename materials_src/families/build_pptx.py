"""Редактируемая презентация «Многодетные семьи России. Меры их защиты».

Каждый слайд повторяет соответствующий кадр исходной вёрстки: текст — настоящие
текстовые поля, значки — фигуры PowerPoint, фотографии и герб — отдельные
изображения, которые можно двигать и заменять.
"""

from __future__ import annotations

import uuid
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
OUT = ROOT.parents[1] / "materials" / "Многодетные_семьи_России.pptx"
FONT = "Noto Sans"

SW, SH = 10.0, 7.5
WHITE = "FFFFFF"
INK = "6A1023"
FOOT = "F6D0DC"
BG_TOP = "6E1228"
BG_BOTTOM = "C21850"

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
ANCHOR = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def no_shadow(shape) -> None:
    shape.shadow.inherit = False
    sp_pr = shape._element.spPr
    effect = sp_pr.find(qn("a:effectLst"))
    if effect is not None:
        sp_pr.remove(effect)


def _solid(shape, fill: str | None, line: str | None = None, line_pt: float = 1.5) -> None:
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
    no_shadow(shape)


def add_shape(slide, kind, x, y, w, h, fill=None, line=None, line_pt=1.5, radius=None):
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    _solid(shape, fill, line, line_pt)
    if radius is not None:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    return shape


def oval(slide, x, y, w, h, fill=None, line=None, line_pt=1.5):
    return add_shape(slide, MSO_SHAPE.OVAL, x, y, w, h, fill, line, line_pt)


def rect(slide, x, y, w, h, fill=None, line=None, line_pt=1.5):
    return add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill, line, line_pt)


def rrect(slide, x, y, w, h, fill=None, line=None, line_pt=1.5, radius=0.2):
    return add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, line_pt, radius)


def freeform(slide, points, fill=None, line=None, line_pt=1.5):
    x0, y0 = points[0]
    builder = slide.shapes.build_freeform(Inches(x0), Inches(y0))
    builder.add_line_segments([(Inches(x), Inches(y)) for x, y in points[1:]], close=True)
    shape = builder.convert_to_shape()
    _solid(shape, fill, line, line_pt)
    return shape


def font_run(run, size, color, bold=False, spacing=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = rgb(color)
    run.font.name = FONT
    r_pr = run._r.get_or_add_rPr()
    r_pr.set("lang", "ru-RU")
    if spacing:
        r_pr.set("spc", str(int(spacing)))
    for tag in ("latin", "ea", "cs"):
        el = r_pr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(r_pr, qn(f"a:{tag}"))
        el.set("typeface", FONT)


def textbox(slide, x, y, w, h, anchor="top", vert=None):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    no_shadow(shape)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.anchor = ANCHOR[anchor]
    if vert:
        body = tf._txBody.find(qn("a:bodyPr"))
        body.set("vert", vert)
    return shape


def write(slide, x, y, w, h, lines, anchor="top", vert=None):
    shape = textbox(slide, x, y, w, h, anchor, vert)
    tf = shape.text_frame
    for i, item in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ALIGN[item.get("align", "left")]
        para.space_before = Pt(item.get("before", 0))
        para.space_after = Pt(item.get("after", 0))
        para.line_spacing = item.get("spacing", 1.0)
        if item.get("bullet"):
            _bullet(para)
        runs = item.get("runs") or [item]
        for part in runs:
            run = para.add_run()
            run.text = part.get("t", "")
            font_run(
                run,
                part.get("size", item.get("size", 16)),
                part.get("color", item.get("color", WHITE)),
                part.get("bold", item.get("bold", False)),
                part.get("spc"),
            )
    return shape


def _bullet(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    for child in list(p_pr):
        if child.tag in {
            qn("a:buNone"),
            qn("a:buFont"),
            qn("a:buChar"),
            qn("a:buAutoNum"),
            qn("a:buClr"),
        }:
            p_pr.remove(child)
    bu_font = etree.SubElement(p_pr, qn("a:buFont"))
    bu_font.set("typeface", FONT)
    bu_clr = etree.SubElement(p_pr, qn("a:buClr"))
    srgb = etree.SubElement(bu_clr, qn("a:srgbClr"))
    srgb.set("val", WHITE)
    bu_char = etree.SubElement(p_pr, qn("a:buChar"))
    bu_char.set("char", "•")
    p_pr.set("marL", str(int(Inches(0.28))))
    p_pr.set("indent", str(int(-Inches(0.28))))


def picture(slide, name, x, y, w, h):
    shape = slide.shapes.add_picture(str(ASSETS / name), Inches(x), Inches(y), Inches(w), Inches(h))
    no_shadow(shape)
    return shape


def send_back(slide, shape):
    sp_tree = slide.shapes._spTree
    el = shape._element
    sp_tree.remove(el)
    sp_tree.insert(2, el)


def group(slide, shapes, name="Группа"):
    shapes = [s for s in shapes if s is not None]
    if len(shapes) < 2:
        return None
    left = min(int(s.left) for s in shapes)
    top = min(int(s.top) for s in shapes)
    right = max(int(s.left) + int(s.width) for s in shapes)
    bottom = max(int(s.top) + int(s.height) for s in shapes)
    ids = [int(el.get("id")) for el in slide.shapes._spTree.iter() if el.get("id") and el.get("id").isdigit()]
    nxt = (max(ids) if ids else 1) + 1
    grp = etree.Element(qn("p:grpSp"))
    nv = etree.SubElement(grp, qn("p:nvGrpSpPr"))
    c_nv = etree.SubElement(nv, qn("p:cNvPr"))
    c_nv.set("id", str(nxt))
    c_nv.set("name", name)
    etree.SubElement(nv, qn("p:cNvGrpSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    grp_pr = etree.SubElement(grp, qn("p:grpSpPr"))
    xfrm = etree.SubElement(grp_pr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", str(left))
    off.set("y", str(top))
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(max(right - left, 1)))
    ext.set("cy", str(max(bottom - top, 1)))
    ch_off = etree.SubElement(xfrm, qn("a:chOff"))
    ch_off.set("x", str(left))
    ch_off.set("y", str(top))
    ch_ext = etree.SubElement(xfrm, qn("a:chExt"))
    ch_ext.set("cx", str(max(right - left, 1)))
    ch_ext.set("cy", str(max(bottom - top, 1)))
    for shape in shapes:
        el = shape._element
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
        grp.append(el)
    slide.shapes._spTree.append(grp)
    return grp


def background(slide):
    shape = rect(slide, 0, 0, SW, SH, fill=BG_TOP)
    fill = shape.fill
    fill.gradient()
    fill.gradient_angle = 225
    stops = fill.gradient_stops
    stops[0].position = 0.0
    stops[0].color.rgb = rgb(BG_TOP)
    stops[1].position = 1.0
    stops[1].color.rgb = rgb(BG_BOTTOM)
    no_shadow(shape)
    send_back(slide, shape)
    return shape


def chrome(slide):
    picture(slide, "seal.png", 0.42, 0.22, 0.72, 0.72)
    write(
        slide,
        1.24,
        0.28,
        4.6,
        0.64,
        [
            {"t": "ПРЕЗИДЕНТСКАЯ", "size": 10, "bold": True, "spc": 140, "spacing": 0.9},
            {"t": "АКАДЕМИЯ", "size": 16, "bold": True, "spc": 220, "spacing": 0.9},
        ],
        anchor="middle",
    )
    rule = rect(slide, 9.28, 0.28, 0.015, 6.55, fill=WHITE)
    _alpha(rule, 0.45)
    write(slide, 9.36, 0.26, 0.52, 1.15, [{"t": "2025", "size": 11, "align": "center"}], anchor="middle", vert="vert")
    write(
        slide,
        9.34,
        5.55,
        0.56,
        1.45,
        [{"t": "РАНХиГС", "size": 11, "align": "center"}],
        anchor="middle",
        vert="vert",
    )
    write(slide, 0.48, 7.02, 3.2, 0.32, [{"t": "Группа 4-1Б", "size": 12, "color": FOOT}])


def _alpha(shape, opacity: float):
    srgb = shape.fill._fill._solidFill.find(qn("a:srgbClr"))
    alpha = srgb.find(qn("a:alpha"))
    if alpha is None:
        alpha = etree.SubElement(srgb, qn("a:alpha"))
    alpha.set("val", str(int(opacity * 100000)))


def new_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background(slide)
    return slide


def title(slide, lines, y=1.08, h=0.95, size=26):
    return write(
        slide,
        0.48,
        y,
        8.6,
        h,
        [{"t": line, "size": size, "bold": True, "spacing": 0.92} for line in lines],
    )


def person(slide, cx, foot, h, color):
    head = h * 0.36
    shapes = [oval(slide, cx - head / 2, foot - h, head, head, fill=color)]
    body_w = h * 0.50
    body_h = h * 0.56
    top = foot - h + head * 0.78
    shapes.append(rrect(slide, cx - body_w / 2, top, body_w, body_h, fill=color, radius=0.5))
    return shapes


def family(slide, x, y, s, color=WHITE, count=5):
    foot = y + s * 0.96
    if count == 5:
        spec = [
            (0.34, 0.78),
            (0.66, 0.78),
            (0.16, 0.52),
            (0.84, 0.52),
            (0.50, 0.40),
        ]
    else:
        spec = [(0.28, 0.86), (0.72, 0.86), (0.50, 0.58)]
    shapes = []
    for cx, hh in spec:
        shapes += person(slide, x + s * cx, foot, s * hh, color)
    return shapes


def house_poly(x, y, s):
    return [
        (x + s * 0.50, y + s * 0.02),
        (x + s * 0.98, y + s * 0.42),
        (x + s * 0.80, y + s * 0.42),
        (x + s * 0.80, y + s * 0.96),
        (x + s * 0.20, y + s * 0.96),
        (x + s * 0.20, y + s * 0.42),
        (x + s * 0.02, y + s * 0.42),
    ]


def badge(slide, x, y, s, glyph):
    shapes = [oval(slide, x, y, s, s, fill=WHITE)]
    pad = s * 0.24
    shapes += glyph(slide, x + pad, y + pad, s * 0.52, INK)
    group(slide, shapes, "Значок")


def glyph_house(slide, x, y, s, color):
    roof = add_shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, x, y, s, s * 0.50, fill=color)
    body = rect(slide, x + s * 0.16, y + s * 0.38, s * 0.68, s * 0.58, fill=color)
    door = rect(slide, x + s * 0.40, y + s * 0.64, s * 0.20, s * 0.32, fill=WHITE)
    return [roof, body, door]


def glyph_heart(slide, x, y, s, color):
    return [add_shape(slide, MSO_SHAPE.HEART, x + s * 0.04, y + s * 0.06, s * 0.92, s * 0.86, fill=color)]


def glyph_percent(slide, x, y, s, color):
    return [write(slide, x, y - s * 0.08, s, s * 1.15, [{"t": "%", "size": max(12, s * 46), "bold": True, "color": color, "align": "center"}], anchor="middle")]


def glyph_plus(slide, x, y, s, color):
    t = s * 0.22
    return [
        rect(slide, x + (s - t) / 2, y + s * 0.08, t, s * 0.84, fill=color),
        rect(slide, x + s * 0.08, y + (s - t) / 2, s * 0.84, t, fill=color),
    ]


def glyph_bars(slide, x, y, s, color):
    heights = (0.46, 0.68, 0.92)
    gap = s * 0.08
    bw = (s - gap * 2) / 3
    shapes = []
    for i, hh in enumerate(heights):
        bh = s * hh
        shapes.append(rect(slide, x + i * (bw + gap), y + s - bh, bw, bh, fill=color))
    return shapes


def glyph_doc(slide, x, y, s, color):
    page = add_shape(slide, MSO_SHAPE.FOLDED_CORNER, x + s * 0.12, y, s * 0.76, s, fill=color)
    lines = [
        rect(slide, x + s * 0.26, y + s * 0.28 + i * s * 0.16, s * 0.36, s * 0.07, fill=WHITE)
        for i in range(3)
    ]
    return [page, *lines]


def glyph_child(slide, x, y, s, color):
    return person(slide, x + s * 0.50, y + s * 0.98, s * 0.96, color)


def glyph_family(slide, x, y, s, color):
    return family(slide, x - s * 0.04, y, s * 1.08, color, count=3)


def glyph_book(slide, x, y, s, color):
    pages = [
        rrect(slide, x, y + s * 0.14, s * 0.47, s * 0.74, fill=color, radius=0.1),
        rrect(slide, x + s * 0.53, y + s * 0.14, s * 0.47, s * 0.74, fill=color, radius=0.1),
    ]
    lines = []
    for i in range(2):
        lines.append(rect(slide, x + s * 0.08, y + s * 0.34 + i * s * 0.18, s * 0.30, s * 0.06, fill=WHITE))
        lines.append(rect(slide, x + s * 0.62, y + s * 0.34 + i * s * 0.18, s * 0.30, s * 0.06, fill=WHITE))
    return [*pages, *lines]


def glyph_building(slide, x, y, s, color):
    shapes = [
        add_shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, x + s * 0.06, y, s * 0.88, s * 0.34, fill=color),
        rect(slide, x + s * 0.14, y + s * 0.30, s * 0.72, s * 0.66, fill=color),
    ]
    col_w = s * 0.07
    for i in range(3):
        cx = x + s * 0.26 + i * s * 0.18
        shapes.append(rect(slide, cx, y + s * 0.46, col_w, s * 0.40, fill=WHITE))
    return shapes


def glyph_wallet(slide, x, y, s, color):
    return [
        rrect(slide, x, y + s * 0.30, s, s * 0.58, fill=color, radius=0.16),
        rrect(slide, x + s * 0.06, y + s * 0.08, s * 0.52, s * 0.32, fill=color, radius=0.25),
        oval(slide, x + s * 0.60, y + s * 0.44, s * 0.28, s * 0.28, fill=WHITE),
    ]


def ruble_sign(slide, x, y, s, color):
    """Знак рубля из фигур: не зависит от того, есть ли символ в шрифте."""
    stem_w = s * 0.16
    shapes = [rect(slide, x + s * 0.18, y + s * 0.06, stem_w, s * 0.88, fill=color)]
    bowl = freeform(
        slide,
        [
            (x + s * 0.28, y + s * 0.08),
            (x + s * 0.62, y + s * 0.08),
            (x + s * 0.86, y + s * 0.28),
            (x + s * 0.62, y + s * 0.48),
            (x + s * 0.28, y + s * 0.48),
            (x + s * 0.34, y + s * 0.36),
            (x + s * 0.58, y + s * 0.36),
            (x + s * 0.70, y + s * 0.28),
            (x + s * 0.58, y + s * 0.20),
            (x + s * 0.34, y + s * 0.20),
        ],
        fill=color,
    )
    shapes.append(bowl)
    shapes.append(rect(slide, x + s * 0.10, y + s * 0.50, s * 0.62, s * 0.11, fill=color))
    shapes.append(rect(slide, x + s * 0.10, y + s * 0.68, s * 0.52, s * 0.11, fill=color))
    return shapes


def icon_row(slide, x, y, kind, text, text_w=6.4, size=15, box=0.48, row_h=0.72):
    glyphs = {
        "house": glyph_house,
        "heart": glyph_heart,
        "percent": glyph_percent,
        "plus": glyph_plus,
        "bars": glyph_bars,
        "doc": glyph_doc,
        "child": glyph_child,
        "family": glyph_family,
        "book": glyph_book,
        "building": glyph_building,
        "wallet": glyph_wallet,
        "ruble": lambda sl, gx, gy, gs, color: ruble_sign(sl, gx, gy, gs, color),
    }
    badge(slide, x, y + max(0, (row_h - box) / 2) * 0.15, box, glyphs[kind])
    return write(
        slide,
        x + box + 0.16,
        y,
        text_w,
        row_h,
        [{"t": text, "size": size, "spacing": 0.98}],
        anchor="middle",
    )


def bullets(slide, x, y, lines, w, size=16, pitch=0.58):
    h = pitch * len(lines) + 0.1
    return write(
        slide,
        x,
        y,
        w,
        h,
        [{"t": line, "size": size, "bullet": True, "spacing": 1.0, "before": 6} for line in lines],
    )


def big_family(slide, x, y, s):
    shapes = family(slide, x, y, s, WHITE, count=5)
    group(slide, shapes, "Семья")


def big_hand(slide, x, y, s):
    """Ладонь из овалов и скруглений: её можно разобрать и перекрасить."""
    shapes = []
    shapes.append(oval(slide, x + s * 0.52, y, s * 0.40, s * 0.40, fill=None, line=WHITE, line_pt=2.25))
    shapes += ruble_sign(slide, x + s * 0.59, y + s * 0.08, s * 0.26, WHITE)
    wrist = rrect(slide, x, y + s * 0.66, s * 0.42, s * 0.22, fill=WHITE, radius=0.45)
    shapes.append(wrist)
    palm = oval(slide, x + s * 0.16, y + s * 0.56, s * 0.66, s * 0.32, fill=WHITE)
    shapes.append(palm)
    thumb = oval(slide, x + s * 0.10, y + s * 0.44, s * 0.26, s * 0.36, fill=WHITE)
    thumb.rotation = -28
    shapes.append(thumb)
    fingers = (
        (0.50, 0.40, 0.16, 0.34, -18),
        (0.64, 0.34, 0.15, 0.36, -4),
        (0.77, 0.40, 0.14, 0.30, 12),
    )
    for fx, fy, fw, fh, rot in fingers:
        finger = rrect(slide, x + s * fx, y + s * fy, s * fw, s * fh, fill=WHITE, radius=0.5)
        finger.rotation = rot
        shapes.append(finger)
    group(slide, shapes, "Выплата")


def big_house_family(slide, x, y, s):
    shapes = [freeform(slide, house_poly(x, y, s), fill=None, line=WHITE, line_pt=2.25)]
    shapes += family(slide, x + s * 0.18, y + s * 0.40, s * 0.64, WHITE, count=3)
    group(slide, shapes, "Дом и семья")


def big_chart(slide, x, y, s):
    shapes = glyph_bars(slide, x, y + s * 0.40, s * 0.62, WHITE)
    shaft = rect(slide, x + s * 0.34, y + s * 0.28, s * 0.50, s * 0.07, fill=WHITE)
    shaft.rotation = -36
    shapes.append(shaft)
    head = add_shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, x + s * 0.70, y + s * 0.02, s * 0.26, s * 0.26, fill=WHITE)
    head.rotation = 55
    shapes.append(head)
    group(slide, shapes, "Рост")


def slide_title(prs):
    slide = new_slide(prs)
    write(
        slide,
        0.55,
        2.35,
        8.4,
        2.7,
        [
            {"t": "Многодетные семьи", "size": 36, "bold": True, "spacing": 0.92},
            {"t": "России.", "size": 36, "bold": True, "spacing": 0.92},
            {"t": "Меры их защиты", "size": 36, "bold": True, "spacing": 0.92},
        ],
    )
    chrome(slide)


def slide_contents(prs):
    slide = new_slide(prs)
    title(slide, ["Содержание"], h=0.6, size=30)
    items = [
        "Введение",
        "Понятие и особенности многодетных семей",
        "Роль многодетных семей для страны",
        "Государственная поддержка",
        "Меры социальной защиты",
        "Региональные практики",
        "Проблемы и перспективы",
        "Заключение",
    ]
    for i, text in enumerate(items):
        y = 1.82 + i * 0.60
        write(slide, 0.55, y, 0.5, 0.42, [{"t": f"{i + 1}.", "size": 16, "bold": True}], anchor="middle")
        write(slide, 1.15, y, 7.6, 0.42, [{"t": text, "size": 16}], anchor="middle")
    chrome(slide)


def slide_intro(prs):
    slide = new_slide(prs)
    title(slide, ["Введение"], h=0.6, size=30)
    write(
        slide,
        0.55,
        2.05,
        8.4,
        1.9,
        [
            {
                "t": "Многодетные семьи — это важнейшая социальная группа, от благополучия которой зависят демографическое, экономическое и культурное развитие России.",
                "size": 18,
                "spacing": 1.05,
            }
        ],
    )
    write(
        slide,
        0.55,
        4.15,
        8.4,
        1.7,
        [
            {
                "t": "Государство уделяет особое внимание поддержке многодетных семей, создавая систему мер социальной защиты и помощи.",
                "size": 18,
                "spacing": 1.05,
            }
        ],
    )
    chrome(slide)


def slide_concept(prs):
    slide = new_slide(prs)
    title(slide, ["Понятие и особенности", "многодетных семей"], y=1.02, h=1.05, size=26)
    write(
        slide,
        0.55,
        2.25,
        5.7,
        1.85,
        [
            {
                "t": "Многодетная семья — это семья, в которой воспитываются трое и более детей, независимо от возраста, включая приёмных и усыновлённых.",
                "size": 16,
                "spacing": 1.05,
            }
        ],
    )
    big_family(slide, 6.45, 2.15, 2.35)
    write(slide, 0.55, 4.30, 4.5, 0.4, [{"t": "Особенности", "size": 18, "bold": True}])
    bullets(
        slide,
        0.55,
        4.78,
        [
            "повышенные материальные расходы",
            "большая нагрузка на родителей",
            "необходимость в дополнительной поддержке",
        ],
        w=8.3,
        size=16,
        pitch=0.52,
    )
    chrome(slide)


def slide_role(prs):
    slide = new_slide(prs)
    title(slide, ["Роль многодетных семей", "для страны"], y=1.02, h=1.0, size=26)
    rows = [
        ("family", "Демографическая устойчивость и рост населения"),
        ("child", "Формирование будущего поколения"),
        ("heart", "Укрепление традиционных ценностей"),
        ("bars", "Развитие экономики и снижение социальной нагрузки на систему"),
    ]
    for i, (kind, text) in enumerate(rows):
        icon_row(slide, 0.5, 2.18 + i * 1.08, kind, text, text_w=4.7, size=15, box=0.5, row_h=0.95)
    picture(slide, "sunset-side.jpg", 6.15, 2.15, 2.95, 4.35)
    chrome(slide)


def slide_support(prs):
    slide = new_slide(prs)
    title(slide, ["Государственная поддержка", "многодетных семей"], y=1.02, h=1.05, size=26)
    rows = [
        ("doc", "Указ Президента РФ «О мерах по поддержке многодетных семей»"),
        ("family", "Национальные проекты: «Демография», «Семья»"),
        ("building", "Федеральные и региональные социальные программы"),
    ]
    for i, (kind, text) in enumerate(rows):
        icon_row(slide, 0.5, 2.20 + i * 1.05, kind, text, text_w=5.35, size=16, box=0.52, row_h=0.95)
    picture(slide, "flag.png", 6.15, 5.20, 3.05, 2.15)
    chrome(slide)


def slide_measures(prs):
    slide = new_slide(prs)
    title(slide, ["Основные меры социальной", "защиты"], y=1.02, h=1.0, size=26)
    rows = [
        ("ruble", "Ежемесячные пособия на детей"),
        ("house", "Материнский капитал"),
        ("percent", "Льготы по оплате жилья и коммунальных услуг"),
        ("plus", "Бесплатная медицинская помощь"),
        ("book", "Льготы на образование и отдых"),
    ]
    for i, (kind, text) in enumerate(rows):
        icon_row(slide, 0.5, 2.18 + i * 0.88, kind, text, text_w=7.9, size=16, box=0.5, row_h=0.78)
    chrome(slide)


def slide_payments(prs):
    slide = new_slide(prs)
    title(slide, ["Денежные выплаты и пособия"], h=0.6, size=26)
    bullets(
        slide,
        0.55,
        2.05,
        [
            "Ежемесячное пособие на детей до 17 лет",
            "Единовременное пособие при рождении ребёнка",
            "Пособие по уходу за ребёнком до 1,5 лет",
            "Региональные выплаты многодетным семьям",
        ],
        w=6.15,
        size=16,
        pitch=0.85,
    )
    big_hand(slide, 6.85, 2.55, 2.15)
    chrome(slide)


def slide_benefits(prs):
    slide = new_slide(prs)
    title(slide, ["Льготы и преференции"], h=0.6, size=28)
    bullets(
        slide,
        0.55,
        1.95,
        [
            "Скидки на оплату ЖКХ",
            "Льготы по транспортному налогу",
            "Преимущественное право на зачисление в детские сады и школы",
            "Бесплатное питание для школьников",
            "Льготы при поступлении в вузы",
        ],
        w=6.2,
        size=16,
        pitch=0.72,
    )
    big_house_family(slide, 6.85, 2.35, 2.25)
    chrome(slide)


def slide_regions(prs):
    slide = new_slide(prs)
    title(slide, ["Региональные практики"], h=0.6, size=28)
    picture(slide, "moscow.jpg", 0.48, 2.05, 3.85, 3.85)
    bullets(
        slide,
        4.55,
        2.15,
        [
            "Москва — городская программа «Многодетная семья»",
            "Санкт-Петербург — льготы на проезд и бесплатные кружки для детей",
            "Регионы — дополнительные выплаты, земельные участки, поддержка в трудоустройстве",
        ],
        w=4.45,
        size=15,
        pitch=1.05,
    )
    chrome(slide)


def slide_problems(prs):
    slide = new_slide(prs)
    title(slide, ["Проблемы"], h=0.6, size=30)
    rows = [
        ("wallet", "Недостаточный уровень доходов у многих семей"),
        ("doc", "Сложности с получением выплат и льгот"),
        ("house", "Доступность жилья и детских садов"),
        ("heart", "Потребность в развитии инфраструктуры и социальной поддержки"),
    ]
    for i, (kind, text) in enumerate(rows):
        icon_row(slide, 0.5, 2.05 + i * 1.12, kind, text, text_w=8.0, size=16, box=0.52, row_h=0.95)
    chrome(slide)


def slide_prospects(prs):
    slide = new_slide(prs)
    title(slide, ["Перспективы"], h=0.6, size=30)
    bullets(
        slide,
        0.55,
        2.0,
        [
            "Расширение и совершенствование мер социальной поддержки",
            "Увеличение доступности жилья и образования",
            "Развитие семейной инфраструктуры",
            "Укрепление института семьи и традиционных ценностей",
        ],
        w=6.2,
        size=16,
        pitch=0.9,
    )
    big_chart(slide, 7.05, 2.7, 2.05)
    chrome(slide)


def slide_conclusion(prs):
    slide = new_slide(prs)
    title(slide, ["Заключение"], h=0.6, size=30)
    write(
        slide,
        0.55,
        2.35,
        8.45,
        3.4,
        [
            {
                "t": "Поддержка многодетных семей — это инвестиция в будущее России. Развитие системы социальной защиты, повышение доступности льгот и создание комфортных условий жизни помогают семьям реализовывать свой потенциал и укрепляют страну.",
                "size": 20,
                "spacing": 1.12,
            }
        ],
    )
    chrome(slide)


def slide_slogan(prs):
    slide = new_slide(prs)
    write(
        slide,
        0.7,
        2.55,
        8.5,
        1.9,
        [
            {"t": "Сильные семьи —", "size": 34, "bold": True, "align": "center", "spacing": 0.95},
            {"t": "сильная Россия!", "size": 34, "bold": True, "align": "center", "spacing": 0.95},
        ],
    )
    rect(slide, 1.7, 4.7, 6.5, 0.018, fill=WHITE)
    chrome(slide)


def slide_thanks(prs):
    slide = new_slide(prs)
    picture(slide, "sunset-wide.jpg", 0.36, 1.08, 8.78, 5.55)
    write(
        slide,
        0.62,
        1.85,
        5.4,
        1.7,
        [
            {"t": "Спасибо", "size": 36, "bold": True, "spacing": 0.92},
            {"t": "за внимание!", "size": 36, "bold": True, "spacing": 0.92},
        ],
    )
    chrome(slide)


def slide_blank(prs):
    slide = new_slide(prs)
    chrome(slide)


def build(path: Path = OUT) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    prs.core_properties.title = "Многодетные семьи России. Меры их защиты"
    prs.core_properties.subject = "Президентская академия, РАНХиГС, 2025"
    prs.core_properties.category = "Группа 4-1Б"
    prs.core_properties.author = "Группа 4-1Б"
    for maker in (
        slide_title,
        slide_contents,
        slide_intro,
        slide_concept,
        slide_role,
        slide_support,
        slide_measures,
        slide_payments,
        slide_benefits,
        slide_regions,
        slide_problems,
        slide_prospects,
        slide_conclusion,
        slide_slogan,
        slide_thanks,
        slide_blank,
    ):
        maker(prs)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    embed_noto(path)
    return path


def _obfuscate(data: bytes, guid: str) -> bytes:
    hexed = guid.strip("{}").replace("-", "")
    raw = bytes.fromhex(hexed)
    key = bytes(
        [
            raw[3],
            raw[2],
            raw[1],
            raw[0],
            raw[5],
            raw[4],
            raw[7],
            raw[6],
            raw[8],
            raw[9],
            raw[10],
            raw[11],
            raw[12],
            raw[13],
            raw[14],
            raw[15],
        ]
    )
    out = bytearray(data)
    for i in range(min(32, len(out))):
        out[i] ^= key[i % 16]
    return bytes(out)


def embed_noto(path: Path) -> None:
    """Вкладывает Noto Sans, чтобы кириллица и знак рубля открывались одинаково."""
    regular = Path("/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf").read_bytes()
    bold = Path("/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf").read_bytes()
    guid_r = "{" + str(uuid.uuid4()).upper() + "}"
    guid_b = "{" + str(uuid.uuid4()).upper() + "}"
    tmp = path.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "ppt/presentation.xml":
                data = _insert_font_list(data)
            elif item.filename == "ppt/_rels/presentation.xml.rels":
                data = _insert_font_rels(data, guid_r, guid_b)
            elif item.filename == "[Content_Types].xml":
                data = _insert_font_type(data)
            dst.writestr(item, data)
        dst.writestr(f"ppt/fonts/{guid_r}.odttf", _obfuscate(regular, guid_r))
        dst.writestr(f"ppt/fonts/{guid_b}.odttf", _obfuscate(bold, guid_b))
    tmp.replace(path)


def _insert_font_list(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    if root.find("p:embeddedFontLst", ns) is not None:
        return xml
    lst = etree.Element(qn("p:embeddedFontLst"))
    font = etree.SubElement(lst, qn("p:embeddedFont"))
    face = etree.SubElement(font, qn("p:font"))
    face.set("typeface", FONT)
    face.set("pitchFamily", "34")
    face.set("charset", "0")
    regular = etree.SubElement(font, qn("p:regular"))
    regular.set(qn("r:id"), "rIdFontRegular")
    bold = etree.SubElement(font, qn("p:bold"))
    bold.set(qn("r:id"), "rIdFontBold")
    root.append(lst)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _insert_font_rels(xml: bytes, guid_r: str, guid_b: str) -> bytes:
    root = etree.fromstring(xml)
    ns = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
    for rid, guid in (("rIdFontRegular", guid_r), ("rIdFontBold", guid_b)):
        rel = etree.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rel_type)
        rel.set("Target", f"fonts/{guid}.odttf")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _insert_font_type(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
    exists = any(el.get("Extension") == "odttf" for el in root)
    if not exists:
        default = etree.Element("{http://schemas.openxmlformats.org/package/2006/content-types}Default")
        default.set("Extension", "odttf")
        default.set("ContentType", "application/vnd.openxmlformats-officedocument.obfuscatedFont")
        root.insert(0, default)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


if __name__ == "__main__":
    target = build()
    print(target)
