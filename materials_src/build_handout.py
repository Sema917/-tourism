"""Одностраничная раздатка: графы «хочу–могу–умею–надо» + 15 пар Климова."""

from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

from content import FORMULA, VERSION
from ddo import QUESTIONS
from docx_kit import (
    INK,
    MUTED,
    add_text,
    cell_text,
    new_doc,
    prevent_row_split,
    set_cell_margins,
    set_row_height,
    set_run_font,
    set_table_borders,
    set_table_widths,
    set_valign,
    shade_cell,
    spacing,
)

NAVY = "6B1E2A"
TEAL = "0F6B5C"
W = 18.5  # usable width at ~1 cm margins

FORMULA_COLORS = {
    "Хочу": ("F8E8EA", "7A1F2B"),
    "Могу": ("E5F2EF", "0F6B5C"),
    "Умею": ("F8EED8", "A86B12"),
    "Надо": ("E8EAF6", "3A46A0"),
}


def build():
    doc = new_doc(f"Наука 0+  ·  Путь в профессию  ·  {VERSION}", first_page_plain=True)
    doc.core_properties.title = "Путь в профессию. Раздатка участника"
    doc.core_properties.subject = f"Одна страница. Комплект {VERSION}"

    normal = doc.styles["Normal"]
    normal.paragraph_format.space_after = Cm(0)
    normal.paragraph_format.space_before = Cm(0)
    normal.paragraph_format.line_spacing = 1.0

    section = doc.sections[0]
    section.left_margin = Cm(0.9)
    section.right_margin = Cm(0.9)
    section.top_margin = Cm(0.7)
    section.bottom_margin = Cm(0.7)
    section.header_distance = Cm(0.3)
    section.footer_distance = Cm(0.3)

    _header(doc)
    _formula(doc)
    _questions(doc)

    path = "/workspace/materials/Путь_в_профессию_раздаточный_материал.docx"
    doc.save(path)
    return path


def _header(doc):
    table = doc.add_table(rows=1, cols=3)
    set_table_widths(table, [7.2, 5.8, 5.5])
    set_table_borders(table, "FFFFFF", "2", "nil")
    left, mid, right = table.rows[0].cells
    shade_cell(left, "FFFFFF")
    shade_cell(mid, "FFFFFF")
    shade_cell(right, "FFFFFF")
    set_cell_margins(left, 20, 20, 20, 20)
    set_cell_margins(mid, 20, 20, 20, 20)
    set_cell_margins(right, 20, 20, 20, 20)
    cell_text(
        left,
        [
            {"t": "НАУКА 0+", "size": 8, "color": NAVY, "bold": True, "after": 0},
            {"t": "Путь в профессию", "size": 14, "color": NAVY, "bold": True, "after": 0},
        ],
    )
    cell_text(
        mid,
        [
            {"t": "Раздатка · 1 страница", "size": 9, "color": MUTED, "after": 0},
            {"t": "Формула + опросник Климова (15)", "size": 9, "color": MUTED, "after": 0},
        ],
    )
    cell_text(
        right,
        [
            {"t": "Класс ________  Имя ____________", "size": 9, "color": INK, "after": 0},
            {"t": "Дата ________", "size": 9, "color": INK, "after": 0},
        ],
    )
    prevent_row_split(table.rows[0])
    # тонкая линия
    line = doc.add_paragraph()
    spacing(line, 2, 4, 1.0)
    pPr = line._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), NAVY)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _formula(doc):
    title = doc.add_paragraph()
    spacing(title, 0, 3, 1.0)
    add_text(title, "Хочу  ·  Могу  ·  Умею  ·  Надо", 11, NAVY, True)
    hint = doc.add_paragraph()
    spacing(hint, 0, 3, 1.0)
    add_text(
        hint,
        "Заполните коротко. Не название профессии — занятие. Пересечение клеток важнее одной мечты.",
        8,
        MUTED,
        False,
    )

    table = doc.add_table(rows=2, cols=2)
    set_table_widths(table, [W / 2, W / 2])
    set_table_borders(table, "FFFFFF", "8", "single")
    prompts = {
        "Хочу": "Что интересно делать даже без отметки?",
        "Могу": "К чему есть силы и склонность?",
        "Умею": "Что уже получалось на деле?",
        "Надо": "Кому это полезно — и вам это не противно?",
    }
    for i, (name, _, _) in enumerate(FORMULA):
        cell = table.cell(i // 2, i % 2)
        fill, bar = FORMULA_COLORS[name]
        shade_cell(cell, fill)
        set_cell_margins(cell, 50, 45, 70, 70)
        set_valign(cell, "top")
        set_row_height(table.rows[i // 2], 2.15)
        cell_text(
            cell,
            [
                {"t": name.upper(), "size": 10, "color": bar, "bold": True, "after": 1},
                {"t": prompts[name], "size": 8, "color": INK, "after": 2},
                {"t": "_" * 42, "size": 8, "color": MUTED, "after": 1},
                {"t": "_" * 42, "size": 8, "color": MUTED, "after": 0},
            ],
        )
    for row in table.rows:
        prevent_row_split(row)

    gap = doc.add_paragraph()
    spacing(gap, 2, 2, 1.0)


def _questions(doc):
    title = doc.add_paragraph()
    spacing(title, 2, 2, 1.0)
    add_text(title, "Опросник Климова · 15 пар", 11, NAVY, True)
    rules = doc.add_paragraph()
    spacing(rules, 0, 3, 1.0)
    add_text(
        rules,
        "Представьте: можете выполнять любую работу. В каждой паре выберите только одно — А или Б. "
        "Первый отклик. Правильных ответов нет. Ключ откроет ведущий.",
        8,
        MUTED,
        False,
    )

    # две колонки по вопросам
    table = doc.add_table(rows=1, cols=2)
    set_table_widths(table, [W / 2, W / 2])
    set_table_borders(table, "FFFFFF", "4", "nil")
    left, right = table.rows[0].cells
    set_cell_margins(left, 20, 20, 40, 80)
    set_cell_margins(right, 20, 20, 80, 40)
    set_valign(left, "top")
    set_valign(right, "top")

    left_qs = QUESTIONS[:8]
    right_qs = QUESTIONS[8:]
    _fill_questions(left, left_qs)
    _fill_questions(right, right_qs)
    prevent_row_split(table.rows[0])

    foot = doc.add_paragraph()
    spacing(foot, 4, 0, 1.0)
    add_text(
        foot,
        "Ведущий тип: ______   Второй: ______   Шаг на месяц: ________________________________",
        8,
        INK,
        False,
    )


def _fill_questions(cell, questions):
    cell.text = ""
    first = True
    for q in questions:
        para = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        spacing(para, 0, 0, 0.98)
        run = para.add_run(f"{q['n']}. ")
        set_run_font(run, 7, NAVY, True)
        run = para.add_run(f"А {q['a']}  ")
        set_run_font(run, 7, INK, False)
        run = para.add_run("☐А")
        set_run_font(run, 7, TEAL, True)

        para2 = cell.add_paragraph()
        spacing(para2, 0, 1, 0.98)
        run = para2.add_run(f"    Б {q['b']}  ")
        set_run_font(run, 7, INK, False)
        run = para2.add_run("☐Б")
        set_run_font(run, 7, TEAL, True)


if __name__ == "__main__":
    print(build())
