#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сборка ВКР в формате DOCX (ГОСТ-подобное оформление)."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
TMP = ROOT / "source_text"
if not TMP.exists():
    TMP = Path("/tmp")
OUT = ROOT / "ВКР_Аллабергенов_А_Деятельность_специалиста_по_социальной_работе_2026.docx"
FIGURES = {
    "page43": ROOT / "figures" / "рис1.png",
    "page53": ROOT / "figures" / "рис2.png",
    "page63": ROOT / "figures" / "рис3.png",
    "page66": ROOT / "figures" / "рис4.png",
    "page68": ROOT / "figures" / "рис5.png",
}
FIGURE_CAPTIONS = {
    "page43": "Рисунок 1 – Структура социально-педагогической технологии",
    "page53": "Рисунок 2 – Работа по профилактике социального сиротства",
    "page63": "Рисунок 3 – Данные с сайта Правительства Орловской области",
    "page66": "Рисунок 4 – Статистика обеспечения жильём детей-сирот",
    "page68": "Рисунок 5 – Данные с сайта Департамента социальной защиты",
}


def set_run_font(run, size=14, bold=False, italic=False, all_caps=False):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    if all_caps:
        run.font.all_caps = True


def set_paragraph_format(
    p,
    *,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    first_line=Cm(1.25),
    space_before=0,
    space_after=0,
    line_spacing=1.5,
):
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    pf.first_line_indent = first_line


def add_page_number(doc: Document):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
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
    set_run_font(run, 12)


def add_centered_lines(doc, lines, *, size=14, bold=False, space_after=0, first_line=Cm(0)):
    for line in lines:
        p = doc.add_paragraph()
        set_paragraph_format(
            p,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            first_line=first_line,
            space_after=space_after,
        )
        run = p.add_run(line)
        set_run_font(run, size=size, bold=bold)


def add_right_block(doc, lines, *, size=12):
    for line in lines:
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.RIGHT, first_line=Cm(0))
        run = p.add_run(line)
        set_run_font(run, size=size)


def add_empty(doc, n=1):
    for _ in range(n):
        p = doc.add_paragraph()
        set_paragraph_format(p, first_line=Cm(0))
        run = p.add_run("")
        set_run_font(run)


def page_break(doc):
    doc.add_page_break()


def add_heading_custom(doc, text, *, size=14, bold=True, center=True, space_before=6, space_after=6):
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        align=WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line=Cm(0),
        space_before=space_before,
        space_after=space_after,
    )
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold)
    return p


def add_body_paragraph(doc, text):
    text = text.strip()
    if not text:
        return
    p = doc.add_paragraph()
    set_paragraph_format(p)
    # Split citation markers «...»[N] and keep them inline
    parts = re.split(r"(«[^»]+»\[\d+\])", text)
    for part in parts:
        if not part:
            continue
        m = re.fullmatch(r"«([^»]+)»\[(\d+)\]", part)
        if m:
            run_q = p.add_run(f"«{m.group(1)}»")
            set_run_font(run_q, italic=False)
            run_n = p.add_run(f"[{m.group(2)}]")
            set_run_font(run_n)
        else:
            run = p.add_run(part)
            set_run_font(run)


def add_figure(doc, key: str):
    path = FIGURES[key]
    p = doc.add_paragraph()
    set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=Cm(0), space_before=6, space_after=6)
    run = p.add_run()
    # width ~15 cm to fit page
    run.add_picture(str(path), width=Cm(12.0))
    cap = doc.add_paragraph()
    set_paragraph_format(cap, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=Cm(0), space_after=12)
    r = cap.add_run(FIGURE_CAPTIONS[key])
    set_run_font(r, size=12)


def load_text(name: str) -> str:
    return (TMP / name).read_text(encoding="utf-8").strip()


def emit_section_text(doc, text: str, *, skip_first_heading=False):
    """Emit paragraphs; handle figure placeholders and headings."""
    # Normalize
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    blocks = re.split(r"\n\s*\n", text)
    first = True
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        # Figure placeholder
        fig = re.fullmatch(r"\[\[РИСУНОК\s+(page\d+)\]\]", block)
        if fig:
            add_figure(doc, fig.group(1))
            continue
        # Inline figure in a longer block
        if "[[РИСУНОК" in block:
            pieces = re.split(r"(\[\[РИСУНОК\s+page\d+\]\])", block)
            buf = []
            for piece in pieces:
                m = re.fullmatch(r"\[\[РИСУНОК\s+(page\d+)\]\]", piece.strip())
                if m:
                    if buf:
                        emit_plain_block(doc, "\n".join(buf), first_heading=first and not skip_first_heading)
                        first = False
                        buf = []
                    add_figure(doc, m.group(1))
                else:
                    if piece.strip():
                        buf.append(piece.strip())
            if buf:
                emit_plain_block(doc, "\n".join(buf), first_heading=first and not skip_first_heading)
                first = False
            continue
        emit_plain_block(doc, block, first_heading=first and not skip_first_heading)
        first = False


def emit_plain_block(doc, block: str, *, first_heading=False):
    lines = [ln.strip() for ln in block.split("\n") if ln.strip()]
    if not lines:
        return
    # Detect headings
    joined = " ".join(lines)
    is_heading = False
    if re.match(r"^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|АННОТАЦИЯ|СОДЕРЖАНИЕ|СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ)\b", joined):
        is_heading = True
    elif re.match(r"^ГЛАВА\s+\d+", joined):
        is_heading = True
    elif re.match(r"^\d+\.\d+\s+", joined) and len(joined) < 220:
        is_heading = True
    elif re.match(r"^Выводы?\s+по\s+главе", joined, re.I):
        is_heading = True

    if is_heading and (first_heading or len(lines) <= 3 or len(joined) < 220):
        # If heading spilled across lines, join
        add_heading_custom(doc, joined, size=14, bold=True, center=True)
        return

    # Numbered list item as separate short para without big first indent change
    if re.match(r"^\d+\.\s+", joined) and len(joined) < 500:
        p = doc.add_paragraph()
        set_paragraph_format(p, first_line=Cm(1.25))
        # citations inside
        parts = re.split(r"(«[^»]+»\[\d+\])", joined)
        for part in parts:
            if not part:
                continue
            m = re.fullmatch(r"«([^»]+)»\[(\d+)\]", part)
            if m:
                p.add_run(f"«{m.group(1)}»")
                set_run_font(p.runs[-1])
                p.add_run(f"[{m.group(2)}]")
                set_run_font(p.runs[-1])
            else:
                p.add_run(part)
                set_run_font(p.runs[-1])
        return

    # Regular paragraph(s): if many lines that look like continuous text, join
    add_body_paragraph(doc, joined)


def build_title_page(doc):
    add_centered_lines(
        doc,
        [
            "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ",
            "РОССИЙСКОЙ ФЕДЕРАЦИИ",
            "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ",
            "УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ",
            "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ",
            "имени И.С. ТУРГЕНЕВА»",
        ],
        size=14,
        bold=True,
    )
    add_empty(doc, 2)
    add_centered_lines(doc, ["ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА"], size=16, bold=True)
    add_empty(doc, 1)
    add_centered_lines(
        doc,
        [
            "по направлению подготовки 39.04.02 Социальная работа",
            "направленность (профиль): Организация социальной работы с разными группами",
            "населения и в различных сферах жизнедеятельности",
        ],
        size=13,
    )
    add_empty(doc, 1)
    add_centered_lines(doc, ["Студента Аллабергенова Атаджана                       шифр 245625"], size=13)
    add_empty(doc, 1)
    add_centered_lines(
        doc,
        [
            "Отделение заочного и очно-заочного образования",
            "Кафедра политологии, социологии и социальной работы",
        ],
        size=13,
    )
    add_empty(doc, 2)
    add_centered_lines(doc, ["Тема выпускной квалификационной работы"], size=13, bold=True)
    add_empty(doc, 1)
    add_centered_lines(
        doc,
        [
            "ДЕЯТЕЛЬНОСТЬ СПЕЦИАЛИСТА ПО СОЦИАЛЬНОЙ",
            "РАБОТЕ С ДЕТЬМИ-СИРОТАМИ И ДЕТЬМИ, ОСТАВШИМИСЯ БЕЗ",
            "ПОПЕЧЕНИЯ РОДИТЕЛЕЙ",
        ],
        size=14,
        bold=True,
    )
    add_empty(doc, 1)
    # signatures
    lines = [
        "Студент                   ______________________   Аллабергенов А.",
        "",
        "Научный",
        "руководитель        ______________________    к.пед.н., доцент Тюрикова Г.Н.",
        "",
        "Зав. кафедрой         ______________________   д.полит.н., профессор Старостенко К.В.",
    ]
    for line in lines:
        p = doc.add_paragraph()
        set_paragraph_format(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=Cm(0))
        run = p.add_run(line)
        set_run_font(run, size=13)
    add_empty(doc, 3)
    add_centered_lines(doc, ["Орёл 2026"], size=14, bold=True)


def build_assignment(doc):
    add_centered_lines(
        doc,
        [
            "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ",
            "РОССИЙСКОЙ ФЕДЕРАЦИИ",
            "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ",
            "УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ",
            "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ",
            "имени И.С. ТУРГЕНЕВА»",
        ],
        size=12,
        bold=True,
    )
    add_empty(doc, 1)
    add_body_paragraph(
        doc,
        "Институт заочного и очно-заочного образования. Кафедра теории и истории социальной педагогики и социальной работы. Направление подготовки 39.04.02 Социальная работа. Направленность (профиль): Организация социальной работы с разными группами населения и в различных сферах жизнедеятельности.",
    )
    add_right_block(
        doc,
        [
            "УТВЕРЖДАЮ:",
            "Зав. кафедрой",
            "Г.Н. Тюрикова",
            "__________________",
            "«25» февраля 2026 г.",
        ],
    )
    add_heading_custom(doc, "ЗАДАНИЕ", size=14, bold=True)
    add_heading_custom(doc, "на выполнение выпускной квалификационной работы", size=13, bold=False, space_before=0)
    add_body_paragraph(doc, "Студента: Аллабергенова Атаджана                                  шифр 245625")
    add_body_paragraph(
        doc,
        "1. Тема ВКР: Деятельность специалиста по социальной работе с детьми-сиротами и детьми, оставшимися без попечения родителей, утверждена приказом по университету «Об утверждении тем выпускных квалификационных работ» № 8-37 от 25.02.2026.",
    )
    add_body_paragraph(doc, "2. Срок сдачи студентом законченной работы «29» октября 2026 г.")
    add_body_paragraph(
        doc,
        "3. Исходные данные к работе: учебная литература, периодические издания, интернет-ресурсы, отчётная документация организации (с базы практики).",
    )
    add_body_paragraph(
        doc,
        "4. Содержание ВКР (перечень подлежащих разработке вопросов): изучить теоретические основы деятельности специалиста по социальной работе с детьми-сиротами и детьми, оставшимися без попечения родителей; охарактеризовать профессиональный портрет и функции специалиста по социальной работе; определить особенности организации деятельности специалиста по работе с детьми-сиротами и детьми, оставшимися без попечения родителей; проанализировать современное состояние жизнеустройства детей-сирот в области; определить опыт организации деятельности службы по устройству детей в семью; рассмотреть перспективы функционирования служб по жизнеустройству детей-сирот в области.",
    )
    add_body_paragraph(doc, "5. Перечень графического материала: 5 рисунков.")
    add_body_paragraph(doc, "Дата выдачи задания «25» февраля 2026 года.")
    add_empty(doc, 1)
    add_body_paragraph(doc, "Научный руководитель ВКР __________________ Тюрикова Г.Н.")
    add_body_paragraph(doc, "Задание принял к исполнению __________________ Аллабергенов А.")
    add_empty(doc, 1)
    add_heading_custom(doc, "КАЛЕНДАРНЫЙ ПЛАН", size=14, bold=True)
    plan = [
        ("Выбор темы, закрепление руководителя", "25 февраля 2026 г."),
        ("Утверждение плана исследования", "27 февраля 2026 г."),
        ("Обзор литературы, представление первой главы", "18 мая 2026 г."),
        ("Представление второй главы", "20 сентября 2026 г."),
        ("Предоставление окончательного варианта работы", "29 октября 2026 г."),
    ]
    for name, date in plan:
        add_body_paragraph(doc, f"{name} — {date}")
    add_empty(doc, 1)
    add_body_paragraph(doc, "Научный руководитель ВКР __________________ Тюрикова Г.Н.")
    add_body_paragraph(doc, "Задание принял к исполнению __________________ Аллабергенов А.")


def build_toc(doc):
    add_heading_custom(doc, "СОДЕРЖАНИЕ", size=14, bold=True)
    items = [
        ("ВВЕДЕНИЕ", "8"),
        ("ГЛАВА 1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ДЕЯТЕЛЬНОСТИ СПЕЦИАЛИСТА ПО СОЦИАЛЬНОЙ РАБОТЕ С ДЕТЬМИ-СИРОТАМИ И ДЕТЬМИ, ОСТАВШИМИСЯ БЕЗ ПОПЕЧЕНИЯ РОДИТЕЛЕЙ", "11"),
        ("1.1 Дети-сироты и дети, оставшиеся без попечения родителей, как объект социальной работы", "11"),
        ("1.2 Профессиональный портрет и функции специалиста по социальной работе", "33"),
        ("Выводы по главе 1", "52"),
        ("ГЛАВА 2. АНАЛИЗ ПРАКТИЧЕСКОГО ОПЫТА И РАЗРАБОТКА ИНДИВИДУАЛЬНОГО ПЛАНА ДЛЯ РАЗВИТИЯ ДЕТЕЙ-СИРОТ", "54"),
        ("2.1 Современное состояние системы жизнеустройства детей-сирот в регионе и опыт организации деятельности службы по устройству детей в семью", "54"),
        ("2.2 Разработка дорожной карты и индивидуального плана развития детей-сирот", "60"),
        ("Выводы по главе 2", "75"),
        ("ЗАКЛЮЧЕНИЕ", "79"),
        ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "83"),
    ]
    for title, page in items:
        p = doc.add_paragraph()
        set_paragraph_format(p, first_line=Cm(0), align=WD_ALIGN_PARAGRAPH.LEFT)
        # tab-like dots
        run = p.add_run(f"{title}")
        set_run_font(run, size=14)
        # page number right-ish via spaces (simple)
        dots = "." * max(3, 80 - len(title))
        run2 = p.add_run(f" {dots} {page}")
        set_run_font(run2, size=14)


def emit_bibliography(doc):
    add_heading_custom(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", size=14, bold=True)
    text = load_text("out_bibliography.txt")
    # drop title line if present
    lines = text.splitlines()
    if lines and "СПИСОК" in lines[0].upper():
        lines = lines[1:]
    # join wrapped entries: entries start with N.
    entries = []
    cur = []
    for ln in lines:
        ln = ln.rstrip()
        if not ln.strip():
            continue
        if re.match(r"^\d+\.\s", ln.strip()):
            if cur:
                entries.append(" ".join(cur))
            cur = [ln.strip()]
        else:
            if cur:
                cur.append(ln.strip())
            else:
                cur = [ln.strip()]
    if cur:
        entries.append(" ".join(cur))
    for entry in entries:
        p = doc.add_paragraph()
        set_paragraph_format(p, first_line=Cm(0), align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        # hanging-ish: no first indent, left aligned block
        pf = p.paragraph_format
        pf.left_indent = Cm(0.75)
        pf.first_line_indent = Cm(-0.75)
        run = p.add_run(entry)
        set_run_font(run, size=12)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(0)


def build():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    add_page_number(doc)

    # Default style
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    build_title_page(doc)
    page_break(doc)
    build_assignment(doc)
    page_break(doc)
    # Annotation
    emit_section_text(doc, load_text("out_annotation.txt"))
    page_break(doc)
    build_toc(doc)
    page_break(doc)
    # Intro + chapters flow continuously (fewer forced blank pages)
    emit_section_text(doc, load_text("out_intro.txt"))
    # Chapter 1
    add_heading_custom(
        doc,
        "ГЛАВА 1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ДЕЯТЕЛЬНОСТИ СПЕЦИАЛИСТА ПО СОЦИАЛЬНОЙ РАБОТЕ С ДЕТЬМИ-СИРОТАМИ И ДЕТЬМИ, ОСТАВШИМИСЯ БЕЗ ПОПЕЧЕНИЯ РОДИТЕЛЕЙ",
        size=14,
        bold=True,
        center=True,
        space_before=10,
        space_after=6,
    )
    emit_section_text(doc, load_text("out_ch1_1.txt"))
    emit_section_text(doc, load_text("out_ch1_2.txt"))
    ch1c = load_text("out_ch1_conclusion.txt")
    if not re.match(r"(?i)выводы?\s+по\s+главе", ch1c):
        ch1c = "Выводы по главе 1\n\n" + ch1c
    emit_section_text(doc, ch1c)
    # Chapter 2
    emit_section_text(doc, load_text("out_ch2_1.txt"))
    emit_section_text(doc, load_text("out_ch2_2.txt"))
    ch2c = load_text("out_ch2_conclusion.txt")
    if not re.match(r"(?i)выводы?\s+по\s+главе", ch2c):
        ch2c = "Выводы по главе 2\n\n" + ch2c
    emit_section_text(doc, ch2c)
    emit_section_text(doc, load_text("out_conclusion.txt"))
    page_break(doc)
    emit_bibliography(doc)

    doc.save(OUT)
    print(f"Saved: {OUT}")


if __name__ == "__main__":
    build()
