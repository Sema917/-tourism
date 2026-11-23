# -*- coding: utf-8 -*-
"""Сборка пояснительной записки ВКР в формате, близком к исходной."""

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Cm, Mm, Pt, Twips, Emu
import copy
import re

FIG = "/agent/repos/-tourism/vkr_src/figures"


def _set_run_font(run, name="Times New Roman", size=14, bold=False, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.name = name
    run.font.size = Pt(size)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)


def _paragraph_format(p, *, align="justify", first=1.25, before=0, after=0,
                      line=1.5, center=False, left_indent=0):
    pf = p.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if center or align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Cm(0)
    elif align == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Cm(first) if first else Cm(0)
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pf.first_line_indent = Cm(0)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(first) if first else Cm(0)
    if left_indent:
        pf.left_indent = Cm(left_indent)


def add_text(doc, text, *, align="justify", first=1.25, before=0, after=0,
             line=1.5, size=14, bold=False, italic=False, center=False):
    p = doc.add_paragraph()
    _paragraph_format(p, align=align, first=first, before=before, after=after,
                      line=line, center=center)
    run = p.add_run(text)
    _set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_mixed(doc, parts, *, align="justify", first=1.25, before=0, after=0,
              line=1.5, size=14, center=False):
    """parts: list of str or (str, bold, italic)."""
    p = doc.add_paragraph()
    _paragraph_format(p, align=align, first=first, before=before, after=after,
                      line=line, center=center)
    for part in parts:
        if isinstance(part, str):
            run = p.add_run(part)
            _set_run_font(run, size=size)
        else:
            text, bold, italic = part
            run = p.add_run(text)
            _set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_page_break(doc):
    p = doc.add_paragraph()
    _paragraph_format(p, first=0, after=0, before=0, line=1.0)
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._element.append(br)


def _shade_header(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "F2F2F2")
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


def add_table(doc, caption, headers, rows, col_widths=None, font=12):
    add_text(doc, caption, align="left", first=0, before=6, after=4, size=14, bold=False)
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(h)
        _set_run_font(run, size=font, bold=True)
        set_cell_border(cell)
        _shade_header(cell)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.0
            run = p.add_run(str(val))
            _set_run_font(run, size=font, bold=False)
            set_cell_border(cell)
    if col_widths:
        for row in table.rows:
            for j, w in enumerate(col_widths):
                row.cells[j].width = Cm(w)
    # empty line after
    add_text(doc, "", first=0, after=0, before=0, line=1.0)
    return table


def add_picture(doc, path, caption, width_cm=15.5):
    p = doc.add_paragraph()
    _paragraph_format(p, center=True, first=0, before=6, after=2)
    run = p.add_run()
    run.add_picture(path, width=Cm(width_cm))
    add_text(doc, caption, center=True, first=0, before=2, after=8, size=14)


def add_footer_page_number(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    _set_run_font(run, size=14)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def setup_section(section):
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(30)
    section.right_margin = Mm(15)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.header_distance = Mm(10)
    section.footer_distance = Mm(10)
    section.different_first_page_header_footer = False


def new_document():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), "Times New Roman")
    pf = style.paragraph_format
    pf.line_spacing = 1.5
    setup_section(doc.sections[0])
    add_footer_page_number(doc.sections[0])
    return doc


def quote_stats(blocks_text):
    total = len(blocks_text)
    quotes = re.findall(r"«[^»]*»", blocks_text)
    qlen = sum(len(q) for q in quotes)
    return total, qlen, (qlen / total if total else 0), len(quotes)
