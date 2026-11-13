#!/usr/bin/env python3
"""Assemble VKR DOCX from rewritten section files."""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Twips

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
FIG = ROOT / "figures"
OUT = ROOT / "VKR_Logvinova.docx"

FONT = "Times New Roman"
# Fallback registered via python-docx if OS lacks TNR; LibreOffice maps Liberation Serif.
FALLBACK_FONT_FILE = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"


def set_run_font(run, size=14, bold=False, italic=False):
    run.bold = bold
    run.italic = italic
    run.font.name = FONT
    run.font.size = Pt(size)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), FONT)
    rFonts.set(qn("w:hAnsi"), FONT)
    rFonts.set(qn("w:cs"), FONT)
    rFonts.set(qn("w:eastAsia"), FONT)


def configure_document(doc: Document):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(14)
    style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.first_line_indent = Cm(1.25)


def add_para(
    doc,
    text,
    *,
    bold=False,
    center=False,
    first_indent=True,
    size=14,
    space_before=0,
    space_after=0,
):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.first_line_indent = Cm(1.25) if first_indent and not center else Cm(0)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Split bold lead-ins like "Актуальность исследования. rest"
    m = re.match(r"^([^.]{3,80}\.)\s+(.*)$", text)
    if (
        not bold
        and not center
        and m
        and m.group(1).rstrip(".").split()[-1:]
        and any(
            m.group(1).startswith(x)
            for x in (
                "Актуальность",
                "Степень изученности",
                "Объект исследования",
                "Предмет исследования",
                "Цель",
                "Методы",
                "Теоретическая значимость",
                "Практическая значимость",
                "Экспериментальная база",
                "Структура работы",
            )
        )
    ):
        r1 = p.add_run(m.group(1) + " ")
        set_run_font(r1, size=size, bold=True)
        add_inline_with_quotes(p, m.group(2), size=size)
    else:
        if bold:
            r = p.add_run(text)
            set_run_font(r, size=size, bold=True)
        else:
            add_inline_with_quotes(p, text, size=size)
    return p


def add_inline_with_quotes(paragraph, text, size=14):
    """Render «...» slightly italic-neutral; keep as normal runs."""
    r = paragraph.add_run(text)
    set_run_font(r, size=size, bold=False)


FIGURE_FILES = {
    1: "fig01.png",
    2: "fig02.png",
    3: "fig03.png",
    4: "fig04.png",
    5: "fig05.png",
    6: "fig06.png",
    7: "fig07.png",
    8: "fig08.png",
    9: "fig09.png",
}


def is_heading(text: str) -> bool:
    t = text.strip()
    if t.startswith("ГЛАВА ") or t.startswith("ВВЕДЕНИЕ") or t.startswith("ЗАКЛЮЧЕНИЕ"):
        return True
    if t.startswith("СПИСОК ИСПОЛЬЗОВАННЫХ"):
        return True
    if t.startswith("ПРИЛОЖЕН"):
        return True
    if t.startswith("Выводы по главе"):
        return True
    if re.match(r"^\d+\.\d+\.\s", t):
        return True
    if t.startswith("АННОТАЦИЯ") or t.startswith("СОДЕРЖАНИЕ"):
        return True
    return False


def is_figure_caption(text: str) -> bool:
    return bool(re.match(r"^Рисунок\s+\d+", text.strip()))


def is_table_caption(text: str) -> bool:
    return bool(re.match(r"^Таблица\s+\d+", text.strip()))


def add_figure(doc, num: int, caption: str):
    path = FIG / FIGURE_FILES[num]
    if path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        run = p.add_run()
        # width ~15cm to fit margins
        run.add_picture(str(path), width=Cm(15.0))
    add_para(doc, caption, center=True, first_indent=False, size=14, space_before=6, space_after=6)


def add_body_from_file(doc, path: Path):
    text = path.read_text(encoding="utf-8").strip()
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    for para in paragraphs:
        if is_figure_caption(para):
            m = re.match(r"Рисунок\s+(\d+)", para)
            num = int(m.group(1)) if m else None
            if num in FIGURE_FILES:
                add_figure(doc, num, para)
            else:
                add_para(doc, para, center=True, first_indent=False)
            continue
        if is_table_caption(para):
            add_para(doc, para, center=True, first_indent=False, bold=True, space_before=6)
            # embed program table collage if table 5
            if para.startswith("Таблица 5"):
                for name in ("tab05a.png", "tab05b.png"):
                    fp = FIG / name
                    if fp.exists():
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.first_line_indent = Cm(0)
                        p.add_run().add_picture(str(fp), width=Cm(15.0))
            continue
        if is_heading(para):
            # long headings may include body on same line for 1.1/1.2 if poorly split
            if re.match(r"^\d+\.\d+\.\s", para) and len(para) > 120:
                # split heading and first sentence if glued
                m = re.match(r"^(\d+\.\d+\.\s[^.]{10,120})\s+(.*)$", para)
                if m and m.group(2)[0].isupper():
                    add_para(doc, m.group(1), bold=True, center=True, first_indent=False, space_before=12, space_after=6)
                    add_para(doc, m.group(2))
                    continue
            add_para(
                doc,
                para,
                bold=True,
                center=not para.startswith("Выводы"),
                first_indent=False,
                space_before=12,
                space_after=6,
            )
            continue
        add_para(doc, para)


def add_title_pages(doc: Document):
    lines_center = [
        "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ",
        "РОССИЙСКОЙ ФЕДЕРАЦИИ",
        "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ",
        "УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ",
        "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ",
        "имени И.С. ТУРГЕНЕВА»",
    ]
    for t in lines_center:
        add_para(doc, t, center=True, first_indent=False, bold=False, space_after=0)
    for _ in range(3):
        add_para(doc, "", center=True, first_indent=False)
    add_para(doc, "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА", center=True, first_indent=False, bold=True)
    add_para(doc, "", center=True, first_indent=False)
    add_para(
        doc,
        "по направлению подготовки 44.04.02 Психолого-педагогическое образование",
        center=True,
        first_indent=False,
    )
    add_para(
        doc,
        "направленность (профиль): Менеджмент социально-педагогической деятельности",
        center=True,
        first_indent=False,
    )
    add_para(doc, "Студента Логвиновой Анастасии Владимировны                    шифр 215980", center=True, first_indent=False)
    add_para(doc, "", center=True, first_indent=False)
    add_para(doc, "Отделение заочного и очно-заочного образования", center=True, first_indent=False)
    add_para(doc, "Кафедра педагогики", center=True, first_indent=False)
    for _ in range(2):
        add_para(doc, "", center=True, first_indent=False)
    add_para(doc, "Тема выпускной квалификационной работы", center=True, first_indent=False)
    add_para(doc, "", center=True, first_indent=False)
    add_para(
        doc,
        "ФОРМИРОВАНИЕ ГАРМОНИЧНЫХ ДЕТСКО-РОДИТЕЛЬСКИХ ОТНОШЕНИЙ В ПРОЦЕССЕ СОВМЕСТНОЙ ДЕЯТЕЛЬНОСТИ РОДИТЕЛЕЙ И ДЕТЕЙ",
        center=True,
        first_indent=False,
        bold=True,
    )
    for _ in range(3):
        add_para(doc, "", center=True, first_indent=False)
    add_para(doc, "Студент                       ____________________   Логвинова А.В.", first_indent=False)
    add_para(doc, "Научный руководитель             ____________________   к.п.н., доцент Гришина Ю.В.", first_indent=False)
    add_para(doc, "И.о. зав. кафедрой      ____________________   д.п.н., профессор Алдошина М.И.", first_indent=False)
    for _ in range(4):
        add_para(doc, "", center=True, first_indent=False)
    add_para(doc, "Орел 2026", center=True, first_indent=False)
    doc.add_page_break()

    # Assignment page (condensed but complete)
    for t in lines_center[:6]:
        add_para(doc, t, center=True, first_indent=False)
    add_para(doc, "", center=True, first_indent=False)
    add_para(doc, "Институт заочного и очно-заочного образования", first_indent=False)
    add_para(doc, "Кафедра теории и истории социальной педагогики и социальной работы", first_indent=False)
    add_para(doc, "Направление подготовки 44.04.02 Психолого-педагогическое образование", first_indent=False)
    add_para(doc, "Направленность (профиль): Менеджмент социально-педагогической деятельности", first_indent=False)
    add_para(doc, "", first_indent=False)
    add_para(doc, "УТВЕРЖДАЮ:", first_indent=False)
    add_para(doc, "Зав. кафедрой Г.Н. Тюрикова", first_indent=False)
    add_para(doc, "_______________  «25» февраля 2026 г.", first_indent=False)
    add_para(doc, "", first_indent=False)
    add_para(doc, "ЗАДАНИЕ", center=True, first_indent=False, bold=True)
    add_para(doc, "на выполнение выпускной квалификационной работы", center=True, first_indent=False, bold=True)
    add_para(doc, "", first_indent=False)
    add_para(doc, "студента Логвиновой Анастасии Владимировны                    шифр 215980", first_indent=False)
    add_para(
        doc,
        "1. Тема ВКР: Формирование гармоничных детско-родительских отношений в процессе совместной деятельности родителей и детей, утверждена приказом по университету «Об утверждении тем выпускных квалификационных работ» № 8-35 от 25.02.2026 г.",
        first_indent=False,
    )
    add_para(doc, "2. Срок сдачи студентом законченной работы «27» октября 2026 г.", first_indent=False)
    add_para(
        doc,
        "3. Исходные данные к работе: учебная литература, периодические издания, интернет-ресурсы, отчетная документация с базы практики.",
        first_indent=False,
    )
    add_para(
        doc,
        "4. Содержание ВКР: теоретические основы формирования гармоничных детско-родительских отношений; опытно-экспериментальная работа по формированию гармоничных детско-родительских отношений в процессе совместной деятельности.",
        first_indent=False,
    )
    add_para(doc, "5. Перечень графического материала: таблицы, 9 рисунков.", first_indent=False)
    add_para(doc, "", first_indent=False)
    add_para(doc, "Дата выдачи задания «25» февраля 2026 г.", first_indent=False)
    add_para(doc, "Научный руководитель ВКР __________ Гришина Ю.В.", first_indent=False)
    add_para(doc, "Задание принял к исполнению __________ Логвинова А.В.", first_indent=False)
    add_para(doc, "", first_indent=False)
    add_para(doc, "КАЛЕНДАРНЫЙ ПЛАН", center=True, first_indent=False, bold=True)
    plan = [
        ("Выбор темы, закрепление руководителя", "25 февраля 2026 г."),
        ("Утверждение плана исследования", "26 февраля 2026 г."),
        ("Обзор литературы, представление первой главы", "9 июня 2026 г."),
        ("Представление второй главы", "8 сентября 2026 г."),
        ("Предоставление окончательного варианта работы", "27 октября 2026 г."),
    ]
    for a, b in plan:
        add_para(doc, f"{a} — {b}", first_indent=False)
    doc.add_page_break()


def add_annotation(doc: Document):
    add_para(doc, "АННОТАЦИЯ", center=True, first_indent=False, bold=True, space_after=12)
    add_para(
        doc,
        "Объем ВКР составляет около 100 страниц и включает введение, две главы, заключение, список использованных источников, приложения.",
    )
    add_para(
        doc,
        "Перечень ключевых слов: родительское отношение, детско-родительские отношения, взаимодействие, совместная деятельность, дошкольный возраст, гармонизация, программа.",
    )
    add_para(doc, "Краткая характеристика работы.", bold=True, first_indent=True)
    add_para(
        doc,
        "Особое значение детско-родительские отношения приобретают в дошкольном возрасте, поскольку этот период связан с установлением эмоциональной связи со значимыми взрослыми, потребностью в поддержке и одобрении, а также с приобретением опыта непосредственного взаимодействия. Важную роль в формировании гармоничных детско-родительских отношений играет совместная деятельность родителей и детей. Именно совместная деятельность создает «зону ближайшего развития», в которой родители лучше понимают индивидуальные особенности, интересы и потребности ребенка, поддерживают его инициативу и самостоятельность, а ребенок приобретает опыт сотрудничества, эмоциональной поддержки и взаимного уважения.",
    )
    add_para(
        doc,
        "Практическая значимость исследования заключается в разработке и апробации программы совместной деятельности родителей и детей, ориентированной на формирование гармоничных детско-родительских отношений. Программа, формы организации совместной деятельности и методические материалы могут использоваться воспитателями и педагогами-психологами дошкольных образовательных организаций, а также родителями.",
    )
    doc.add_page_break()


def add_toc(doc: Document):
    add_para(doc, "СОДЕРЖАНИЕ", center=True, first_indent=False, bold=True, space_after=12)
    items = [
        "ВВЕДЕНИЕ",
        "ГЛАВА 1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ФОРМИРОВАНИЯ ГАРМОНИЧНЫХ ДЕТСКО-РОДИТЕЛЬСКИХ ОТНОШЕНИЙ",
        "1.1. Понятие и особенности формирования детско-родительских отношений в современной семье",
        "1.2. Совместная деятельность родителей и детей как средство формирования гармоничных детско-родительских отношений",
        "ГЛАВА 2. ОПЫТНО-ЭКСПЕРИМЕНТАЛЬНАЯ РАБОТА ПО ФОРМИРОВАНИЮ ГАРМОНИЧНЫХ ДЕТСКО-РОДИТЕЛЬСКИХ ОТНОШЕНИЙ В ПРОЦЕССЕ СОВМЕСТНОЙ ДЕЯТЕЛЬНОСТИ",
        "2.1. Диагностика особенностей детско-родительских отношений в семьях детей дошкольного возраста",
        "2.2. Разработка, реализация и оценка результатов программы совместной деятельности родителей и детей",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
        "ПРИЛОЖЕНИЯ",
    ]
    for it in items:
        add_para(doc, it, first_indent=False)
    doc.add_page_break()


def add_biblio_lines(doc: Document):
    lines = (SRC / "biblio_master.txt").read_text(encoding="utf-8").splitlines()
    buf = []
    for line in lines:
        if not line.strip():
            if buf:
                add_para(doc, " ".join(buf), first_indent=False)
                buf = []
            continue
        if line.strip().startswith("СПИСОК"):
            if buf:
                add_para(doc, " ".join(buf), first_indent=False)
                buf = []
            add_para(doc, line.strip(), center=True, first_indent=False, bold=True, space_after=12)
            continue
        if re.match(r"^\d+\.\s", line.strip()) and buf:
            add_para(doc, " ".join(buf), first_indent=False)
            buf = [line.strip()]
        else:
            buf.append(line.strip())
    if buf:
        add_para(doc, " ".join(buf), first_indent=False)


def add_appendix(doc: Document):
    add_para(doc, "ПРИЛОЖЕНИЯ", center=True, first_indent=False, bold=True, space_after=12)
    app = (SRC / "app_clean.txt").read_text(encoding="utf-8") if (SRC / "app_clean.txt").exists() else (SRC / "app.txt").read_text(encoding="utf-8")
    for para in [p.strip() for p in re.split(r"\n\s*\n", app) if p.strip()]:
        if para.startswith("Приложение") or para.startswith("ПРИЛОЖЕН"):
            add_para(doc, para, center=True, first_indent=False, bold=True, space_before=12)
        else:
            add_para(doc, para, first_indent=False)


def main():
    required = [
        "out_intro.txt",
        "out_ch1_1.txt",
        "out_ch1_2.txt",
        "out_ch2_1.txt",
        "out_ch2_2.txt",
        "out_concl.txt",
    ]
    missing = [f for f in required if not (SRC / f).exists()]
    if missing:
        raise SystemExit(f"Missing rewritten sections: {missing}")

    doc = Document()
    configure_document(doc)
    add_title_pages(doc)
    add_annotation(doc)
    add_toc(doc)

    for fname in required:
        add_body_from_file(doc, SRC / fname)
        # page break between major parts except inside chapter
        if fname in {"out_intro.txt", "out_ch1_2.txt", "out_ch2_2.txt", "out_concl.txt"}:
            doc.add_page_break()

    add_biblio_lines(doc)
    doc.add_page_break()
    add_appendix(doc)

    doc.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
