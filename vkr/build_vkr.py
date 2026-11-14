# -*- coding: utf-8 -*-
"""Сборка выпускной квалификационной работы в формате документа Ворд."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Mm, Pt, RGBColor, Twips

from sources import SOURCES

ROOT = Path(__file__).resolve().parent
FIG = ROOT / "figures"
OUT = ROOT / "ВКР_Аллабергенов_А_2026.docx"

FONT = "Times New Roman"
BLACK = RGBColor(0, 0, 0)


def _rfonts(rPr, name=FONT):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)


def set_run(run, size=14, bold=False, italic=False, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = BLACK
    run.font.italic = italic
    _rfonts(run._element.get_or_add_rPr(), name)


def set_paragraph_format(p, *, align="justify", indent=True, before=0, after=0, line=1.5, keep_next=False):
    pf = p.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Cm(1.25) if indent else Cm(0)
    align_map = {
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    p.alignment = align_map[align]
    pPr = p._p.get_or_add_pPr()
    if keep_next:
        kn = OxmlElement("w:keepNext")
        pPr.append(kn)


def add_text(p, text, size=14, bold=False, italic=False):
    run = p.add_run(text)
    set_run(run, size=size, bold=bold, italic=italic)
    return run


def add_page_number(paragraph):
    run = paragraph.add_run()
    set_run(run, size=14)
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)


def shade_cell(cell, color="E7E6E6"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    trPr.append(cant)


class Builder:
    def __init__(self):
        self.doc = Document()
        self._setup()

    def _setup(self):
        section = self.doc.sections[0]
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(1.5)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.header_distance = Cm(0.6)
        section.footer_distance = Cm(0.6)
        section.different_first_page_header_footer = False

        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.paragraph_format.first_line_indent = Cm(0)
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(0)
        add_page_number(hp)

        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.paragraph_format.first_line_indent = Cm(0)
        # пустой нижний колонтитул: номер стоит сверху, как в исходном файле

        normal = self.doc.styles["Normal"]
        normal.font.name = FONT
        normal.font.size = Pt(14)
        normal.font.color.rgb = BLACK
        _rfonts(normal.element.get_or_add_rPr())
        normal.paragraph_format.line_spacing = 1.5
        normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.space_before = Pt(0)

    def page_break(self):
        self.doc.add_page_break()

    def center(self, text, *, size=14, bold=False, before=0, after=0, italic=False):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="center", indent=False, before=before, after=after)
        add_text(p, text, size=size, bold=bold, italic=italic)
        return p

    def left(self, text, *, size=14, bold=False, before=0, after=0):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="left", indent=False, before=before, after=after)
        add_text(p, text, size=size, bold=bold)
        return p

    def body(self, text, *, before=0, after=0, first=True):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="justify", indent=first, before=before, after=after)
        add_text(p, text)
        return p

    def h1(self, text):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="center", indent=False, before=0, after=12, keep_next=True)
        add_text(p, text, bold=True)
        return p

    def h2(self, text):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="center", indent=False, before=12, after=12, keep_next=True)
        add_text(p, text, bold=True)
        return p

    def caption(self, text):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="center", indent=False, before=6, after=12)
        add_text(p, text, size=14, bold=False)
        return p

    def image(self, path: Path, caption: str, width_cm=15.5):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="center", indent=False, before=6, after=0, keep_next=True)
        run = p.add_run()
        set_run(run, size=14)
        run.add_picture(str(path), width=Cm(width_cm))
        self.caption(caption)

    def table(self, caption, headers, rows, font=11):
        cap = self.doc.add_paragraph()
        set_paragraph_format(cap, align="left", indent=False, before=12, after=6, keep_next=True)
        add_text(cap, caption, size=14, bold=True)

        tbl = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        tbl.autofit = True
        tbl.allow_autofit = True
        for j, h in enumerate(headers):
            cell = tbl.rows[0].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            set_paragraph_format(p, align="center", indent=False, line=1.0)
            add_text(p, h, size=font, bold=True)
            shade_cell(cell)
            set_cell_border(cell)
        prevent_row_split(tbl.rows[0])
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                cell = tbl.rows[i + 1].cells[j]
                cell.text = ""
                p = cell.paragraphs[0]
                set_paragraph_format(p, align="left", indent=False, line=1.0)
                add_text(p, val, size=font, bold=False)
                set_cell_border(cell)
            prevent_row_split(tbl.rows[i + 1])
        # отступ после таблицы
        self.body("", first=False)

    def bibliography_item(self, number, text):
        p = self.doc.add_paragraph()
        set_paragraph_format(p, align="justify", indent=False, before=0, after=2, line=1.5)
        pf = p.paragraph_format
        pf.left_indent = Cm(1.0)
        pf.first_line_indent = Cm(-1.0)
        add_text(p, f"{number}. {text}")

    def save(self, path=OUT):
        self.doc.save(str(path))
        return path


def render(blocks, path=OUT):
    b = Builder()
    for kind, *rest in blocks:
        if kind == "break":
            b.page_break()
        elif kind == "center":
            text, kwargs = rest[0], (rest[1] if len(rest) > 1 else {})
            b.center(text, **kwargs)
        elif kind == "left":
            text, kwargs = rest[0], (rest[1] if len(rest) > 1 else {})
            b.left(text, **kwargs)
        elif kind == "p":
            b.body(rest[0])
        elif kind == "h1":
            b.h1(rest[0])
        elif kind == "h2":
            b.h2(rest[0])
        elif kind == "img":
            b.image(rest[0], rest[1], rest[2] if len(rest) > 2 else 15.5)
        elif kind == "table":
            b.table(rest[0], rest[1], rest[2])
        elif kind == "bib":
            b.bibliography_item(rest[0], rest[1])
        else:
            raise ValueError(kind)
    return b.save(path)


if __name__ == "__main__":
    from text_main import make_blocks

    pages = {
        "intro": 9,
        "ch1": 11,
        "s11": 11,
        "s12": 41,
        "v1": 53,
        "ch2": 54,
        "s21": 54,
        "s22": 66,
        "v2": 74,
        "end": 76,
        "bib": 78,
    }
    path = render(make_blocks(pages))
    print(path)
