"""Общая вёрстка документов Word для комплекта."""

from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor, Twips

FONT = "Calibri"
INK = "1C2830"
MUTED = "5E6C76"
NAVY = "16324A"
TEAL = "0F7A6C"
CREAM = "F7F4EE"
LINE = "E4DCD0"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def set_run_font(run, size=11, color=INK, bold=False, italic=False, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)


def _set_style_font(style, size, color, bold=False):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = rgb(color)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name if False else FONT)


def new_doc(footer_left: str, first_page_plain: bool = True):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.55)
    section.bottom_margin = Cm(1.55)
    section.header_distance = Cm(0.55)
    section.footer_distance = Cm(0.5)
    section.different_first_page_header_footer = first_page_plain

    normal = doc.styles["Normal"]
    _set_style_font(normal, 11, INK)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.line_spacing = 1.08

    for name, size, before, after in (
        ("Heading 1", 18, 0, 8),
        ("Heading 2", 14, 12, 4),
        ("Heading 3", 12, 8, 2),
    ):
        style = doc.styles[name]
        _set_style_font(style, size, NAVY, bold=True)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.05
        style.paragraph_format.keep_with_next = True

    _header_footer(section.header, footer_left, True)
    _header_footer(section.footer, footer_left, False)
    if first_page_plain:
        section.first_page_header.paragraphs[0].text = ""
        section.first_page_footer.paragraphs[0].text = ""
    core = doc.core_properties
    core.author = "Комплект «Путь в профессию», фестиваль «Наука 0+»"
    core.language = "ru-RU"
    core.category = "Наука 0+"
    return doc


def _header_footer(container, left_text, is_header):
    p = container.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if is_header else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if is_header:
        run = p.add_run(left_text)
        set_run_font(run, 9, MUTED)
    else:
        run = p.add_run(left_text + "     ")
        set_run_font(run, 9, MUTED)
        _add_page_field(p)


def _add_page_field(paragraph):
    run = paragraph.add_run()
    set_run_font(run, 9, MUTED)
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


def content_width_cm() -> float:
    return 18.0


def shade_cell(cell, fill: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    tcMar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_valign(cell, value="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:vAlign")):
        tcPr.remove(old)
    v = OxmlElement("w:vAlign")
    v.set(qn("w:val"), value)
    tcPr.append(v)


def set_table_widths(table, widths_cm):
    table.autofit = False
    table.allow_autofit = False
    total = int(sum(widths_cm) * 567)
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(total))
    tblW.set(qn("w:type"), "dxa")
    tblInd = tblPr.find(qn("w:tblInd"))
    if tblInd is None:
        tblInd = OxmlElement("w:tblInd")
        tblPr.append(tblInd)
    tblInd.set(qn("w:w"), "0")
    tblInd.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
        for w in widths_cm:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(int(w * 567)))
            grid.append(gc)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(int(widths_cm[i] * 567)))
            tcW.set(qn("w:type"), "dxa")


def set_table_borders(table, color="D9D0C4", sz="4", mode="single"):
    tbl = table._tbl
    tblPr = tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), mode)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def prevent_row_split(row):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))


def set_row_height(row, cm, rule="atLeast"):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    for old in trPr.findall(qn("w:trHeight")):
        trPr.remove(old)
    trh = OxmlElement("w:trHeight")
    trh.set(qn("w:val"), str(int(cm * 567)))
    trh.set(qn("w:hRule"), rule)
    trPr.append(trh)


def spacing(paragraph, before=0, after=6, line=1.08, keep_next=False):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.keep_together = True
    if keep_next:
        pf.keep_with_next = True


def add_text(paragraph, text, size=11, color=INK, bold=False, italic=False):
    run = paragraph.add_run(text)
    set_run_font(run, size, color, bold, italic)
    return run


def p(doc, text="", size=11, color=INK, bold=False, italic=False, before=0, after=6, align="left", line=1.08):
    paragraph = doc.add_paragraph()
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]
    spacing(paragraph, before, after, line)
    if text:
        add_text(paragraph, text, size, color, bold, italic)
    return paragraph


def rich(doc, parts, before=0, after=6, align="left", line=1.08, size=11):
    paragraph = doc.add_paragraph()
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }[align]
    spacing(paragraph, before, after, line)
    for text, kwargs in parts:
        add_text(paragraph, text, size=kwargs.get("size", size), **{k: v for k, v in kwargs.items() if k != "size"})
    return paragraph


def h1(doc, text):
    paragraph = doc.add_heading(text, level=1)
    spacing(paragraph, 0, 6, 1.05, keep_next=True)
    return paragraph


def h2(doc, text):
    paragraph = doc.add_heading(text, level=2)
    spacing(paragraph, 10, 3, 1.05, keep_next=True)
    return paragraph


def h3(doc, text):
    paragraph = doc.add_heading(text, level=3)
    spacing(paragraph, 8, 2, 1.0, keep_next=True)
    return paragraph


def bullet(doc, text, level=0):
    paragraph = p(doc, text, size=11, before=0, after=2)
    paragraph.paragraph_format.left_indent = Cm(0.75 + 0.4 * level)
    paragraph.paragraph_format.first_line_indent = Cm(-0.35)
    if paragraph.runs:
        paragraph.runs[0].text = "•  " + paragraph.runs[0].text
    return paragraph


def page_break(doc):
    doc.add_page_break()


def banner(doc, text, fill=NAVY, color="FFFFFF", size=13):
    table = doc.add_table(rows=1, cols=1)
    set_table_widths(table, [content_width_cm()])
    set_table_borders(table, fill, "0", "nil")
    cell = table.cell(0, 0)
    shade_cell(cell, fill)
    set_cell_margins(cell, 70, 70, 120, 120)
    set_valign(cell, "center")
    paragraph = cell.paragraphs[0]
    spacing(paragraph, 0, 0, 1.0)
    add_text(paragraph, text, size, color, True)
    p(doc, "", size=6, after=6, before=0)
    return table


def callout(doc, title, body, fill="E7F3F1", bar=TEAL):
    table = doc.add_table(rows=1, cols=2)
    set_table_widths(table, [0.28, content_width_cm() - 0.28])
    set_table_borders(table, fill, "0", "nil")
    left, right = table.rows[0].cells
    shade_cell(left, bar)
    shade_cell(right, fill)
    set_cell_margins(left, 40, 40, 40, 40)
    set_cell_margins(right, 80, 80, 120, 120)
    set_valign(left, "center")
    paragraph = right.paragraphs[0]
    spacing(paragraph, 0, 0, 1.05)
    if title:
        add_text(paragraph, title, 11, NAVY, True)
        paragraph2 = right.add_paragraph()
        spacing(paragraph2, 2, 0, 1.05)
        add_text(paragraph2, body, 10.5, INK)
    else:
        add_text(paragraph, body, 10.5, INK)
    prevent_row_split(table.rows[0])
    p(doc, "", size=4, after=4)
    return table


def lines(doc, prompt, count=2, size=11):
    paragraph = p(doc, prompt, size=size, bold=True, color=NAVY, before=2, after=1)
    for _ in range(count):
        line = doc.add_paragraph()
        spacing(line, 0, 0, 1.0)
        line.paragraph_format.space_before = Pt(1)
        _bottom_border(line)
        add_text(line, " ", 12, INK)
    return paragraph


def _bottom_border(paragraph, color="B7C3C9"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "dotted")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)


def cell_text(cell, blocks):
    """blocks: list of paragraph dicts with runs or t/size/color/bold."""
    first = True
    for block in blocks:
        paragraph = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        align = block.get("align", "left")
        paragraph.alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
        }[align]
        spacing(paragraph, block.get("before", 0), block.get("after", 2), block.get("line", 1.02))
        runs = block.get("runs")
        if runs is None:
            runs = [{"t": block.get("t", ""), "size": block.get("size", 11), "color": block.get("color", INK), "bold": block.get("bold", False), "italic": block.get("italic", False)}]
        for part in runs:
            add_text(
                paragraph,
                part.get("t", ""),
                part.get("size", 11),
                part.get("color", INK),
                part.get("bold", False),
                part.get("italic", False),
            )


def spacer(doc, pts=6):
    p(doc, "", size=pts, after=0, before=0)
