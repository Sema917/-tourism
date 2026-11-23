#!/usr/bin/env python3
"""Полный текст курсового проекта Смирнова Д. А. с дописанной программой
и технико-экономическим обоснованием по методике дипломов кафедры."""

from __future__ import annotations

import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

PDF = Path(
    "/home/ubuntu/.cursor/projects/agent-repos-tourism/uploads/"
    "__________________1__d60a.pdf"
)
IMG = Path("/tmp/vkr/course_img")
OUT = Path(
    "/agent/repos/-tourism/coursework/"
    "Курсовой_проект_Смирнов_Д.А_пожарная_безопасность.docx"
)

FONT = "Times New Roman"
MONO = "Courier New"


def rub(value) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def money(n: int) -> str:
    return f"{n:,}".replace(",", " ")


# --- комплектующие по структурной и принципиальной схемам проекта ---
# Количество: 3 оповещателя Астра и 2 Маяка — рисунок 2.1;
# 4 клапана — рисунок 4.1, по одному на зону; 4 оросителя — по числу клапанов;
# 2 реле групп оповещателей на рисунке 2.1 и 4 реле клапанов по тексту раздела 4.
ITEMS = [
    ("ПЛК110-24.30.Р-М", 1, 48251, "прайс ОВЕН, модификация 24 В"),
    ("Оповещатель Астра-10 исп. 3", 3, 810, "розничная цена"),
    ("Оповещатель Маяк-24-3М", 2, 445, "рекомендованная розничная цена"),
    ("Реле промежуточное РП slim, 24 В", 6, 275, "прайс EKF AVERES"),
    ("Извещатель тепловой ИП 103-5/1", 1, 160, "розничная цена"),
    ("Извещатель пламени ИПП-Ex (ИП 330-8)", 1, 10553, "прайс поставщика"),
    ("Реле давления воды, 24 В", 1, 2180, "изделие того же класса, открытый прайс"),
    ("Насос пожарный НЦПН-70/100.Э24", 1, 539000, "прайс поставщика"),
    ("Реле силовое Charge Mate 2502", 1, 38200, "прайс поставщика"),
    ("Блок питания 12 В, 60 Вт", 1, 2650, "блок питания того же класса"),
    ("Блок питания 24 В, 60 Вт", 1, 2980, "блок питания того же класса"),
    ("Клапан электромагнитный DN15, 220 В, НЗ", 4, 4287, "латунный, прайс"),
    ("Ороситель спринклерный СВН-10", 4, 271, "прайс, исполнение 57–68 °C"),
]

LEAD_MONTH, ENG_MONTH, DAYS_IN_MONTH = 70000, 45000, 23
LEAD_DAY = rub(LEAD_MONTH / DAYS_IN_MONTH)  # 3043, как в образцах кафедры
ENG_DAY = rub(ENG_MONTH / DAYS_IN_MONTH)  # 1957
# этапы курсового проекта, дни: исполнитель / руководитель
STAGES = [
    ("Анализ предметной области", 10, 1),
    ("Разработка структурной схемы", 5, 1),
    ("Подбор элементной базы", 8, 1),
    ("Принципиальная электрическая схема", 8, 1),
    ("Функциональная схема автоматизации", 5, 0),
    ("Алгоритм и программа для ПЛК", 12, 1),
    ("Технико-экономическое обоснование", 4, 0),
    ("Оформление пояснительной записки", 8, 1),
]
ENG_DAYS = sum(s[1] for s in STAGES)
LEAD_DAYS = sum(s[2] for s in STAGES)
assert ENG_DAYS == 60 and LEAD_DAYS == 6

ZP_ENG = ENG_DAY * ENG_DAYS
ZP_LEAD = LEAD_DAY * LEAD_DAYS
ZP = ZP_ENG + ZP_LEAD

TAXES = [
    ("НДФЛ", Decimal("0.13")),
    ("Обязательные пенсионные выплаты", Decimal("0.22")),
    ("Медицинское страхование", Decimal("0.051")),
    ("Страхование от несчастных случаев", Decimal("0.002")),
    ("Социальное страхование", Decimal("0.029")),
]
TAX_ROWS = [(name, rate, rub(Decimal(ZP) * rate)) for name, rate in TAXES]
TAX_SUM = sum(row[2] for row in TAX_ROWS)
# образцы сводят ставку к 43,2 %; расхождение округления гасится в последней строке
TARGET_TAX = rub(Decimal(ZP) * Decimal("0.432"))
TAX_ROWS[-1] = (TAX_ROWS[-1][0], TAX_ROWS[-1][1], TAX_ROWS[-1][2] + (TARGET_TAX - TAX_SUM))
TAX_SUM = sum(row[2] for row in TAX_ROWS)
assert TAX_SUM == TARGET_TAX

ZP_FULL = ZP + TAX_SUM
MAT = sum(q * p for _, q, p, _ in ITEMS)
PUMP = 539000
MOUNT = rub(Decimal(MAT) * Decimal("0.10"))
OTHER = rub(Decimal(ZP_FULL + MAT + MOUNT) * Decimal("0.05"))
TOTAL = ZP_FULL + MAT + MOUNT + OTHER
AUTO = MAT - PUMP  # всё, кроме насоса

SHARES = {
    "Заработная плата с начислениями": ZP_FULL,
    "Материальные затраты": MAT,
    "Затраты на монтаж": MOUNT,
    "Прочие затраты": OTHER,
}


def shares_percent():
    raw = {k: Decimal(v) / Decimal(TOTAL) * 100 for k, v in SHARES.items()}
    rounded = {k: rub(x * 10) / 10 for k, x in raw.items()}  # one decimal via tenths
    # rub() is int; do one decimal properly
    one = []
    for k, v in SHARES.items():
        one.append((k, (Decimal(v) / Decimal(TOTAL) * Decimal(100)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)))
    drift = Decimal("100.0") - sum(x[1] for x in one)
    one[-1] = (one[-1][0], one[-1][1] + drift)
    return one


SHARE_ROWS = shares_percent()
assert sum(x[1] for x in SHARE_ROWS) == Decimal("100.0")

PUMP_SHARE = (Decimal(PUMP) / Decimal(MAT) * 100).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
AUTO_SHARE = (Decimal(AUTO) / Decimal(MAT) * 100).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
AUTO_TO_PUMP = (Decimal(AUTO) / Decimal(PUMP) * 100).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
KIT = MAT + MOUNT  # порог прямого ущерба, при котором комплект окупается одним пожаром
FACTORY_PUMP = 1_679_000  # ЛЗПМ, «от» для НЦПН-70/100.М.Э24
MARGIN_VS_FACTORY = FACTORY_PUMP - TOTAL


def pct(value) -> str:
    return f"{value}".replace(".", ",")


def set_run_font(run, name, size, bold=False, italic=False, color=None):
    run.bold = bold
    run.italic = italic
    run.font.name = name
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), name)


def set_paragraph_format(p, align="justify", first=True, before=0, after=0, line=1.5, center=False):
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
        pf.first_line_indent = Cm(1.25) if first else Cm(0)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(1.25) if first else Cm(0)


def add_text(p, text, size=14, bold=False, italic=False, font=FONT):
    r = p.add_run(text)
    set_run_font(r, font, size, bold, italic)
    return r


def body(doc, text, first=True, size=14, bold=False, italic=False, align="justify", before=0, after=0, center=False):
    p = doc.add_paragraph()
    set_paragraph_format(p, align=align, first=first, before=before, after=after, center=center)
    add_text(p, text, size=size, bold=bold, italic=italic)
    return p


def heading(doc, text, level=1):
    p = doc.add_paragraph()
    set_paragraph_format(p, center=True, before=12 if level == 1 else 10, after=8, first=False)
    add_text(p, text, size=14, bold=True)
    return p


def bullet(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(p, align="justify", first=False, before=0, after=0)
    p.paragraph_format.left_indent = Cm(1.25)
    add_text(p, "• " + text)
    return p


def caption(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(p, center=True, first=False, before=4, after=8)
    add_text(p, text, size=14)
    return p


def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
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
        el.set(qn("w:color"), "000000")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def add_table(doc, headers, rows, col_widths=None, font=11):
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
        p.paragraph_format.first_line_indent = Cm(0)
        add_text(p, h, size=font, bold=True)
        shade_cell(cell, "E8EEF4")
        set_cell_border(cell)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.first_line_indent = Cm(0)
            bold = str(val).startswith("Итого") or str(val).startswith("Всего") or str(val).startswith("Общая")
            add_text(p, str(val), size=font, bold=bold)
            set_cell_border(cell)
            if bold:
                shade_cell(cell, "F4F7FA")
    if col_widths:
        for row in table.rows:
            for j, w in enumerate(col_widths):
                row.cells[j].width = Cm(w)
    doc.add_paragraph()
    return table


def add_page_number(paragraph):
    run1 = paragraph.add_run()
    set_run_font(run1, FONT, 12)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run1._r.append(fld_begin)
    run2 = paragraph.add_run()
    set_run_font(run2, FONT, 12)
    run2._r.append(instr)
    run3 = paragraph.add_run()
    set_run_font(run3, FONT, 12)
    run3._r.append(fld_end)


def setup_section(doc):
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.left_margin = Mm(30)
    sec.right_margin = Mm(15)
    sec.top_margin = Mm(20)
    sec.bottom_margin = Mm(20)
    sec.header_distance = Mm(10)
    sec.footer_distance = Mm(10)
    footer = sec.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(p)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(14)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)


HEAD_RE = re.compile(
    r"^(\d+\.(?:\d+\.)?\s+.+|"
    r"\d+\.[А-ЯA-Z].+|"
    r"ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК ЛИТЕРАТУРЫ|СОДЕРЖАНИЕ|"
    r"ПРИЛОЖЕНИЕ\s+[А-Я])$"
)


def is_heading_line(s: str) -> bool:
    if s in {"ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ЛИТЕРАТУРЫ", "СОДЕРЖАНИЕ"}:
        return True
    if re.match(r"^ПРИЛОЖЕНИЕ\s+[А-ЯA-Z]", s):
        return True
    if re.match(r"^\d+\.\s*[А-ЯA-Z]", s) and len(s) < 110:
        return True
    if re.match(r"^\d+\.\d+\.?\s+\S", s) and len(s) < 120:
        return True
    return False


def iter_page_rows(page):
    """Строки страницы сверху вниз.

    В PDF выровненный по ширине абзац записан отдельными словами с одной
    базовой линией. Слова с одной высоты собираются обратно в строку.
    Две колонки с большим зазором — строка таблицы характеристик.
    """
    d = page.get_text("dict")
    raw = []
    for b in d["blocks"]:
        if b.get("type") == 1:
            w = b["bbox"][2] - b["bbox"][0]
            h = b["bbox"][3] - b["bbox"][1]
            if w > 80 and h > 60:
                raw.append((b["bbox"][1], "img", b))
            continue
        for line in b.get("lines", []):
            text = "".join(s["text"] for s in line["spans"]).strip()
            if not text:
                continue
            x0, y0, x1, _y1 = line["bbox"]
            if y0 < 55 and re.fullmatch(r"\d{1,3}", text):
                continue
            raw.append((y0, "frag", (x0, x1, text)))
    raw.sort(key=lambda item: (item[0], 0 if item[1] == "img" else 1))

    rows = []
    bucket = []
    bucket_y = None

    def flush_bucket():
        nonlocal bucket
        if not bucket:
            return
        bucket.sort(key=lambda frag: frag[0])
        clusters = [[bucket[0]]]
        for frag in bucket[1:]:
            if frag[0] - clusters[-1][-1][1] > 55:
                clusters.append([frag])
            else:
                clusters[-1].append(frag)

        def cluster_text(cluster):
            return re.sub(r"\s+", " ", " ".join(frag[2] for frag in cluster)).strip()

        left_x = clusters[0][0][0]
        # правая колонка таблиц характеристик начинается около x=320
        left_frags = [frag for frag in bucket if frag[0] < 318]
        right_frags = [frag for frag in bucket if frag[0] >= 318]
        right_text = cluster_text(right_frags) if right_frags else ""
        if (
            left_frags
            and right_frags
            and left_frags[0][0] < 120
            and right_frags[0][0] < 360
            and len(right_text) <= 40
            and not any(frag[0] >= 380 for frag in bucket)
        ):
            rows.append(("pair", cluster_text(left_frags), right_text))
        else:
            text = re.sub(r"\s+", " ", " ".join(cluster_text(c) for c in clusters)).strip()
            rows.append(("line", text, left_x))
        bucket = []

    for y, kind, payload in raw:
        if kind == "img":
            flush_bucket()
            rows.append(("img", payload))
            bucket_y = None
            continue
        if bucket_y is None or abs(y - bucket_y) <= 2.2:
            bucket.append(payload)
            if bucket_y is None:
                bucket_y = y
        else:
            flush_bucket()
            bucket = [payload]
            bucket_y = y
    flush_bucket()
    return rows


def save_image(page, block, name):
    IMG.mkdir(parents=True, exist_ok=True)
    # берём xref страницы, если блок не содержит байтов
    xref = None
    if "image" in block:
        # dict image blocks from get_text don't include xref; use page images by order
        pass
    path = IMG / name
    # rendered clip of the image bbox
    clip = pymupdf.Rect(block["bbox"])
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), clip=clip, alpha=False)
    pix.save(path)
    return path


def normalize_heading(text):
    return re.sub(r"^(\d+)\.([А-ЯA-Z])", r"\1. \2", text)


def extract_body(doc_pdf):
    """Страницы 4–35: введение и главы 1–6, с рисунками и таблицами характеристик."""
    flow = []
    img_n = 0
    para = None
    table = None

    def flush_para():
        nonlocal para
        if para:
            flow.append(("p", re.sub(r"\s+", " ", " ".join(para)).strip()))
            para = None

    def flush_table():
        nonlocal table
        if table:
            flow.append(("tbl", table))
            table = None

    for i in range(3, 35):
        page = doc_pdf[i]
        for row in iter_page_rows(page):
            if row[0] == "img":
                flush_para()
                flush_table()
                img_n += 1
                path = save_image(page, row[1], f"fig_{i+1}_{img_n}.png")
                flow.append(("img", path))
                continue
            if row[0] == "pair":
                flush_para()
                _kind, left, right = row
                if table is None:
                    table = []
                table.append([left, right])
                continue

            _kind, text, x0 = row
            if table and x0 >= 280 and len(text) < 40 and table[-1][1] == "":
                table[-1][1] = text
                continue
            if (
                table is not None
                and 80 <= x0 < 160
                and len(text) < 45
                and not text.startswith(("Рисунок", "Таблица", "Продолжение", "•"))
                and not is_heading_line(text)
            ):
                # перенос ячейки пишется со строчной; новый параметр — с прописной
                if table and text[:1].islower() and not re.search(r"[.!?]", text):
                    table[-1][0] = (table[-1][0] + " " + text).strip()
                    continue
                table.append([text, ""])
                continue

            if text.startswith(("Таблица", "Продолжение таблицы", "Продолжение")):
                flush_para()
                flush_table()
                flow.append(("c", text))
                table = []
                continue

            flush_table()
            if text.startswith("Рисунок"):
                flush_para()
                flow.append(("c", text))
                continue
            if flow and flow[-1][0] == "c" and len(text) < 24 and not text[:1].isupper():
                flow[-1] = ("c", flow[-1][1] + " " + text)
                continue
            if re.match(r"^[A-Z]{1,3}\s*[–-]\s", text):
                flush_para()
                flow.append(("p", text))
                continue
            if is_heading_line(text):
                flush_para()
                flow.append(("h", normalize_heading(text)))
                continue
            if (
                flow
                and flow[-1][0] == "h"
                and x0 > 160
                and text.isupper()
                and not flow[-1][1].endswith(".")
            ):
                flow[-1] = ("h", flow[-1][1] + " " + text)
                continue
            if text.startswith("•") or re.match(r"^\d+\)\s", text):
                flush_para()
                flow.append(("b", text.lstrip("•").strip()))
                continue
            if flow and flow[-1][0] == "b" and x0 >= 150:
                flow[-1] = ("b", flow[-1][1] + " " + text)
                continue
            if flow and flow[-1][0] == "b" and flow[-1][1][:1].isdigit() and x0 < 110:
                flow[-1] = ("b", flow[-1][1] + " " + text)
                continue
            if para and para[-1].rstrip().endswith((".", "!", "?")) and text[:1].isupper():
                flush_para()
                para = [text]
            elif x0 >= 110 and (para is None or para[-1].rstrip().endswith((".", "!", "?", ":", ";"))):
                flush_para()
                para = [text]
            elif para is None:
                para = [text]
            else:
                para.append(text)
    flush_para()
    flush_table()
    return flow


def emit_flow(doc, flow):
    for kind, payload in flow:
        if kind == "h":
            level = 2
            if payload in {"ВВЕДЕНИЕ"} or re.match(r"^\d+\.\s*[А-ЯA-Z]", payload):
                level = 1
            heading(doc, payload, level)
            if payload.startswith("6."):
                heading(doc, "6.1. Блок-схема алгоритма", 2)
        elif kind == "b":
            if re.match(r"^\d+\)", payload):
                p = doc.add_paragraph()
                set_paragraph_format(p, align="justify", first=False)
                p.paragraph_format.left_indent = Cm(1.25)
                add_text(p, payload)
            else:
                bullet(doc, payload)
            if payload.startswith("Разработка алгоритмов"):
                bullet(doc, "Технико-экономическое обоснование разработанной системы.")
        elif kind == "c":
            caption(doc, payload)
        elif kind == "img":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.first_line_indent = Cm(0)
            p.add_run().add_picture(str(payload), width=Cm(15.5))
        elif kind == "tbl":
            add_table(doc, ["Параметр", "Значение"], payload, col_widths=[9.5, 6.5], font=12)
        else:
            body(doc, payload)


def add_title(doc):
    lines_center = [
        "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
        "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ",
        "УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ",
        "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ",
        "имени И. С. ТУРГЕНЕВА»",
    ]
    for i, line in enumerate(lines_center):
        body(doc, line, center=True, first=False, size=14, bold=(i == 0), before=0, after=0)
    body(doc, "", center=True, first=False)
    body(doc, "Институт приборостроения, автоматизации и информационных технологий", center=True, first=False, size=14)
    body(doc, "Кафедра автоматизированных систем управления и кибернетики", center=True, first=False, size=14)
    body(doc, "", center=True, first=False)
    body(doc, "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА К КУРСОВОМУ ПРОЕКТУ", center=True, first=False, size=16, bold=True, before=12)
    body(doc, "по дисциплине «Проектная деятельность»", center=True, first=False, before=8)
    body(doc, "Тема курсового проекта", center=True, first=False, before=10)
    body(doc, "«Разработка системы пожарной безопасности»", center=True, first=False, bold=True)
    body(doc, "", center=True, first=False)
    body(doc, "Студент                                              Смирнов Дмитрий Алексеевич", first=False, align="left")
    body(doc, "Шифр                                                 211297", first=False, align="left")
    body(
        doc,
        "Направление подготовки 15.03.04 Автоматизация технологических процессов и производств",
        first=False,
        align="left",
    )
    body(doc, "Направленность (профиль) «Киберфизические системы»", first=False, align="left")
    body(doc, "Руководитель                                        Семин А. Г.", first=False, align="left", before=8)
    body(doc, "Нормоконтроль                                       к.э.н. Марганова О. Н.", first=False, align="left")
    body(doc, "Орёл 2023", center=True, first=False, before=18)
    doc.add_page_break()


def add_assignment(doc):
    heading(doc, "ЗАДАНИЕ", 1)
    body(doc, "на курсовой проект по дисциплине «Проектная деятельность»", center=True, first=False)
    body(
        doc,
        "Студенту Смирнову Дмитрию Алексеевичу, шифр 211297, направление подготовки 15.03.04 "
        "«Автоматизация технологических процессов и производств», направленность (профиль) "
        "«Киберфизические системы».",
    )
    body(doc, "1. Тема: «Разработка системы пожарной безопасности».", first=False, align="left")
    body(
        doc,
        "2. Исходные данные: материалы кафедры автоматизированных систем управления и кибернетики; "
        "ГОСТ Р 59638-2021, ГОСТ 34698-2020; руководство по эксплуатации ПЛК ОВЕН 110; "
        "открытые прайсы на выбранную элементную базу.",
        first=False,
        align="left",
    )
    body(doc, "3. Содержание пояснительной записки:", first=False, align="left")
    for item in [
        "анализ предметной области;",
        "разработка структурной схемы системы;",
        "подбор элементной базы;",
        "разработка принципиальной электрической схемы;",
        "разработка функциональной схемы автоматизации;",
        "разработка алгоритма в виде блок-схемы и программного обеспечения;",
        "технико-экономическое обоснование.",
    ]:
        bullet(doc, item)
    body(doc, "4. Перечень графического материала:", first=False, align="left")
    for item in [
        "структурная схема системы;",
        "принципиальная электрическая схема;",
        "функциональная схема автоматизации;",
        "блок-схема программы.",
    ]:
        bullet(doc, item)
    body(
        doc,
        "5. Задание выдал руководитель курсового проекта Семин А. Г. Задание принял к исполнению студент Смирнов Д. А.",
    )
    doc.add_page_break()


def add_contents(doc):
    heading(doc, "СОДЕРЖАНИЕ", 1)
    items = [
        "ВВЕДЕНИЕ",
        "1. АНАЛИЗ ПРЕДМЕТНОЙ ОБЛАСТИ",
        "1.1. Функции и структура систем охранно-пожарной сигнализации",
        "1.2. Типы датчиков охранной сигнализации",
        "2. РАЗРАБОТКА СТРУКТУРНОЙ СХЕМЫ СИСТЕМЫ",
        "3. ПОДБОР ЭЛЕМЕНТНОЙ БАЗЫ",
        "4. РАЗРАБОТКА ПРИНЦИПИАЛЬНОЙ ЭЛЕКТРИЧЕСКОЙ СХЕМЫ",
        "5. РАЗРАБОТКА ФУНКЦИОНАЛЬНОЙ СХЕМЫ АВТОМАТИЗАЦИИ",
        "6. РАЗРАБОТКА АЛГОРИТМОВ В ВИДЕ БЛОК-СХЕМЫ И ПРОГРАММНОГО ОБЕСПЕЧЕНИЯ",
        "6.1. Блок-схема алгоритма",
        "6.2. Программа управления для ПЛК ОВЕН 110",
        "7. ТЕХНИКО-ЭКОНОМИЧЕСКОЕ ОБОСНОВАНИЕ",
        "7.1. Перечень работ и трудоёмкость",
        "7.2. Расчёт стоимости комплектующих",
        "7.3. Расчёт заработной платы и начислений",
        "7.4. Бюджет затрат на объект автоматизации",
        "7.5. Оценка экономического результата",
        "7.6. SWOT-анализ",
        "ЗАКЛЮЧЕНИЕ",
        "СПИСОК ЛИТЕРАТУРЫ",
        "ПРИЛОЖЕНИЕ А",
        "ПРИЛОЖЕНИЕ Б",
        "ПРИЛОЖЕНИЕ В",
    ]
    for it in items:
        body(doc, it, first=False, align="left", size=14)
    doc.add_page_break()


CODE = """PROGRAM PLC_PRG
VAR
    (* Входы. Тепловой ИП 103-5/1 нормально замкнутый *)
    Flame    : BOOL; (* ИП 330-8: TRUE — обнаружено пламя *)
    HeatNC   : BOOL; (* ИП 103-5/1: TRUE — цепь цела, FALSE — нагрев *)
    (* Выходы *)
    Astra    : BOOL; (* Астра-10, свет и звук *)
    Mayak    : BOOL; (* Маяк-24-3М *)
    Valve    : ARRAY[1..4] OF BOOL; (* клапаны DN15 по зонам *)
    Pump     : BOOL; (* пуск насоса через силовое реле *)
    Fire     : ARRAY[1..4] OF BOOL;
    AnyFire  : BOOL;
    i        : INT;
END_VAR

(* Извещатели, показанные на структурной схеме, отнесены к зоне 1.
   Зоны 2–4 повторяют ту же логику, когда на их входы будут
   заведены свои шлейфы. *)
Fire[1] := Flame OR (NOT HeatNC);
Fire[2] := FALSE;
Fire[3] := FALSE;
Fire[4] := FALSE;

AnyFire := FALSE;
FOR i := 1 TO 4 DO
    Valve[i] := Fire[i];
    IF Fire[i] THEN
        AnyFire := TRUE;
    END_IF;
END_FOR;

Astra := AnyFire;
Mayak := AnyFire;
Pump := AnyFire;
"""


def add_software(doc):
    heading(doc, "6.2. Программа управления для ПЛК ОВЕН 110", 2)
    body(
        doc,
        "Блок-схема на рисунке 6.1 задаёт цикл программы, но самого текста программы в записке не было. "
        "Ниже он дописан на языке структурированного текста среды CODESYS 2.3, которой, как указано в разделе 3.3, "
        "программируется ПЛК110. Программа повторяет блок-схему: пока возгорания нет, оповещение выключено, "
        "клапаны закрыты и насос стоит; при сигнале зоны включаются оповещатели этой тревоги, открывается клапан "
        "зоны и запускается насос.",
    )
    body(
        doc,
        "На структурной схеме показаны один извещатель пламени ИП 330-8 и один тепловой извещатель. "
        "Оба отнесены к первой зоне. Тепловой ИП 103-5/1 нормально замкнутый, поэтому нагрев читается как размыкание. "
        "На принципиальной схеме четыре электромагнитных клапана, поэтому в программе четыре зоны. "
        "Входы зон 2–4 оставлены выключенными до подключения своих извещателей: логика для них та же, что для зоны 1. "
        "Датчик давления воды на схемах есть, но блок-схема на рисунке 6.1 пускает насос сразу по факту возгорания, "
        "без проверки давления. Программа повторяет блок-схему и датчик давления в условии пуска не использует.",
    )
    for line in CODE.strip().splitlines():
        p = doc.add_paragraph()
        set_paragraph_format(p, align="left", first=False, line=1.0, before=0, after=0)
        p.paragraph_format.left_indent = Cm(0.5)
        add_text(p, line if line else " ", size=10, font=MONO)
    body(
        doc,
        "После загрузки программы в энергонезависимую память ПЛК110 она запускается вместе с питанием шкафа. "
        "Пока Fire[1] ложно, Astra, Mayak, все Valve и Pump остаются выключенными. "
        "Как только пламя или нагрев дают истину, клапан первой зоны открывается, оба оповещателя включаются "
        "и через силовое реле получает команду насос. Когда извещатели возвращаются в норму, выходы снимаются "
        "и опрос продолжается. Это и есть цикл, нарисованный на рисунке 6.1.",
        before=8,
    )


def add_economics(doc):
    heading(doc, "7. ТЕХНИКО-ЭКОНОМИЧЕСКОЕ ОБОСНОВАНИЕ", 1)
    body(
        doc,
        "Курсовой проект посвящён системе пожарной безопасности на ПЛК ОВЕН 110: обнаружение возгорания, "
        "световое и звуковое оповещение, пуск воды по зонам. Готовых шкафов с тем же составом на рынке много, "
        "поэтому экономический раздел считает не «выгоду вообще», а стоимость именно этого комплекта: "
        "приборы с рисунков 2.1 и 4.1, труд на проект и затраты на монтаж.",
    )
    body(
        doc,
        "Расчёт построен так же, как технико-экономическое обоснование выпускных работ кафедры. "
        "Сначала фиксируются этапы и их длительность, затем цена комплектующих, основная заработная плата, "
        "начисления, монтаж и прочие затраты. Монтаж принят в размере 10 % от стоимости оборудования, "
        "прочие затраты — 5 % от суммы заработной платы с начислениями, материалов и монтажа.",
    )

    heading(doc, "7.1. Перечень работ и трудоёмкость", 2)
    body(
        doc,
        "Над проектом работают два человека: руководитель и исполнитель (студент). "
        "Трудоёмкость оценена экспертно в человеко-днях по фактическому составу записки. "
        "Сумма по исполнителю — 60 дней, по руководителю — 6 дней.",
    )
    cap = caption(doc, "Таблица 7.1 – Этапы курсового проекта и загрузка исполнителей")
    add_table(
        doc,
        ["Этап", "Исполнитель, дни", "Руководитель, дни"],
        [[name, str(a), str(b)] for name, a, b in STAGES]
        + [[f"Итого", str(ENG_DAYS), str(LEAD_DAYS)]],
        col_widths=[9.5, 3.5, 3.5],
    )

    heading(doc, "7.2. Расчёт стоимости комплектующих", 2)
    body(
        doc,
        "Количество приборов взято из схем проекта, а не «на помещение вообще». "
        "На рисунке 2.1 три оповещателя Астра-10 исп. 3, два Маяка-24-3М, по одному ПЛК, извещателю пламени, "
        "тепловому извещателю, датчику давления, силовому реле и насосу, два промежуточных реле на группы оповещателей. "
        "На рисунке 4.1 четыре электромагнитных клапана, поэтому в спецификацию добавлены четыре реле клапанов "
        "и четыре оросителя СВН-10 — по одному на зону. Блоков питания два, на 12 В и на 24 В, как описано в разделе 4.",
    )
    body(
        doc,
        "Цены сняты с открытых прайсов на ноябрь 2026 года и округлены до рубля. "
        "Для реле давления воды и двух блоков питания взят открытый прайс изделия того же класса и мощности: "
        "в записке нет торговой марки этих двух позиций. Остальные цены относятся к тем приборам, которые названы в главе 3.",
    )
    caption(doc, "Таблица 7.2 – Расчёт стоимости комплектующих")
    bom_rows = []
    for n, (name, qty, price, _src) in enumerate(ITEMS, 1):
        bom_rows.append([f"{n}. {name}", str(qty), money(price), money(qty * price)])
    bom_rows.append(["Итого", "—", "—", money(MAT)])
    add_table(
        doc,
        ["Наименование", "Кол-во, шт.", "Цена, руб.", "Сумма, руб."],
        bom_rows,
        col_widths=[8.2, 2.4, 3.0, 3.2],
        font=10,
    )
    body(
        doc,
        f"Общая стоимость комплектующих — {money(MAT)} руб. Насос НЦПН-70/100.Э24 занимает {money(PUMP)} руб., "
        f"это {pct(PUMP_SHARE)} % комплекта. Остальная автоматика, датчики, оповещение, клапаны и оросители "
        f"стоят {money(AUTO)} руб., или {pct(AUTO_SHARE)} % комплекта. На фоне цены насоса это {pct(AUTO_TO_PUMP)} %: "
        f"автоматика добавляет обнаружение, тревогу и пуск воды по зоне.",
    )
    body(
        doc,
        "Для сравнения по открытым карточкам тот же ряд насосов оценивается по-разному: НЦПН-70/100.Э24 — "
        f"{money(PUMP)} руб. у поставщика, а НЦПН-70/100.М.Э24 завод ЛЗПМ указывает от {money(FACTORY_PUMP)} руб. "
        "Отдельный приёмно-контрольный комплект без насоса такой подачи обычно заметно дешевле полевой автоматики "
        "этого проекта, но он и не подаёт 70 л/с.",
    )

    heading(doc, "7.3. Расчёт заработной платы и начислений", 2)
    body(
        doc,
        "Месячный оклад руководителя принят 70 000 руб., оклад исполнителя — 45 000 руб. "
        "В месяце 23 рабочих дня, дневная ставка округляется до рубля до умножения на дни:",
    )
    body(doc, f"Здн.рук = {money(LEAD_MONTH)} / {DAYS_IN_MONTH} = {money(LEAD_DAY)} руб./день;", first=False, align="left")
    body(doc, f"Здн.исп = {money(ENG_MONTH)} / {DAYS_IN_MONTH} = {money(ENG_DAY)} руб./день.", first=False, align="left")
    body(
        doc,
        f"Основная заработная плата равна дневной ставке, умноженной на дни из таблицы 7.1: "
        f"руководитель {money(LEAD_DAY)} × {LEAD_DAYS} = {money(ZP_LEAD)} руб., "
        f"исполнитель {money(ENG_DAY)} × {ENG_DAYS} = {money(ZP_ENG)} руб.",
    )
    caption(doc, "Таблица 7.3 – Основная заработная плата")
    add_table(
        doc,
        ["Исполнитель", "Оклад, руб./мес.", "Дневная ставка, руб.", "Дни", "Зарплата, руб."],
        [
            ["Руководитель", money(LEAD_MONTH), money(LEAD_DAY), str(LEAD_DAYS), money(ZP_LEAD)],
            ["Исполнитель", money(ENG_MONTH), money(ENG_DAY), str(ENG_DAYS), money(ZP_ENG)],
            ["Итого", "—", "—", str(ENG_DAYS + LEAD_DAYS), money(ZP)],
        ],
        col_widths=[3.4, 3.2, 3.4, 1.8, 3.2],
        font=10,
    )
    body(
        doc,
        "Начисления взяты той же долей, что и в расчётах выпускных работ кафедры: 43,2 % к основной заработной плате. "
        "В эту долю входят НДФЛ 13 %, пенсионные взносы 22 %, медицинское страхование 5,1 %, "
        "социальное страхование 2,9 % и страхование от несчастных случаев 0,2 %. "
        f"База начислений — {money(ZP)} руб.",
    )
    caption(doc, "Таблица 7.4 – Начисления на заработную плату")
    tax_table_rows = [
        [f"{i}. {name}", f"{str((rate * 100).normalize()).replace('.', ',')} %", money(amount)]
        for i, (name, rate, amount) in enumerate(TAX_ROWS, 1)
    ]
    tax_table_rows.append(["Всего", "43,2 %", money(TAX_SUM)])
    add_table(
        doc,
        ["Вид начисления", "Ставка", "Сумма, руб."],
        tax_table_rows,
        col_widths=[9.5, 3.0, 3.5],
    )
    body(
        doc,
        f"Заработная плата с начислениями: {money(ZP)} + {money(TAX_SUM)} = {money(ZP_FULL)} руб.",
    )

    heading(doc, "7.4. Бюджет затрат на объект автоматизации", 2)
    body(
        doc,
        f"Монтаж — 10 % от комплектующих: 0,10 × {money(MAT)} = {money(MOUNT)} руб. "
        f"Прочие затраты — 5 % от суммы заработной платы с начислениями, материалов и монтажа: "
        f"0,05 × ({money(ZP_FULL)} + {money(MAT)} + {money(MOUNT)}) = {money(OTHER)} руб.",
    )
    body(
        doc,
        f"Полная стоимость проекта К = {money(ZP_FULL)} + {money(MAT)} + {money(MOUNT)} + {money(OTHER)} "
        f"= {money(TOTAL)} руб.",
    )
    caption(doc, "Таблица 7.5 – Бюджет затрат")
    add_table(
        doc,
        ["Статья", "Сумма, руб.", "Доля, %"],
        [[name, money(SHARES[name]), str(pct).replace(".", ",")] for name, pct in SHARE_ROWS]
        + [["Общая сумма затрат", money(TOTAL), "100,0"]],
        col_widths=[9.0, 3.8, 3.0],
    )
    mat_pct = next(pct for name, pct in SHARE_ROWS if name.startswith("Материальные"))
    zp_pct = next(pct for name, pct in SHARE_ROWS if name.startswith("Заработная"))
    body(
        doc,
        f"Как видно из таблицы 7.5, проект стоит {money(TOTAL)} руб. "
        f"Больше всего занимают материальные затраты — {str(mat_pct).replace('.', ',')} %. "
        f"Заработная плата с начислениями — {str(zp_pct).replace('.', ',')} %. "
        "Такое соотношение получается из-за насоса: без него структура была бы ближе к обычной проектной работе, "
        "где главная статья — труд.",
    )

    heading(doc, "7.5. Оценка экономического результата", 2)
    body(
        doc,
        "Сравнивать весь проект с дешёвым бытовым комплектом датчиков некорректно: в спецификации стоит "
        f"насос подачи 70 л/с за {money(PUMP)} руб. Корректная опора — цена этого насоса и цена старшей модификации того же ряда.",
    )
    body(
        doc,
        f"Комплект приборов стоит {money(MAT)} руб., из них автоматика сверх насоса — {money(AUTO)} руб. "
        f"Монтаж комплекта — {money(MOUNT)} руб. Вместе оборудование и монтаж, то есть то, что остаётся на объекте, "
        f"составляют {money(KIT)} руб. Это порог прямого ущерба: один предотвращённый пожар с ущербом не ниже "
        f"{money(KIT)} руб. окупает поставленный комплект без повторного счёта проектного труда.",
    )
    body(
        doc,
        f"Полная стоимость проекта с заработной платой — {money(TOTAL)} руб. "
        f"Открытая цена насоса НЦПН-70/100.М.Э24 у завода-изготовителя начинается от {money(FACTORY_PUMP)} руб. "
        f"Даже весь проект, включая оплату 66 человеко-дней, меньше этой цены на {money(MARGIN_VS_FACTORY)} руб., "
        "хотя в проект уже входят извещатели, оповещение, клапаны и программа пуска по зоне. "
        "Значит, собирать систему на ПЛК110 вокруг насоса нижней ценовой позиции выгоднее, чем ориентироваться "
        "на верхнюю заводскую цену насоса без этих функций.",
    )
    body(
        doc,
        "Отдельно стоит ограничение силового реле. Charge Mate 2502 в спецификации коммутирует 500 А, "
        "но это реле аккумуляторных цепей 12/24 В, а не пускатель двигателя насоса от сети. "
        "Для реального пуска насоса мощностью порядка 100 кВт понадобятся контактор и устройство плавного пуска. "
        "Их цена в таблицу 7.2 не входила, поэтому бюджет — нижняя оценка силовой части. "
        "На вывод по сигнализации и зонным клапанам эта оговорка не влияет: их стоимость уже посчитана.",
    )

    heading(doc, "7.6. SWOT-анализ", 2)
    body(
        doc,
        "SWOT-анализ сводит внутренние и внешние факторы проекта в четыре группы: сильные стороны, слабые стороны, "
        "возможности и угрозы. Для этой системы они такие.",
    )
    caption(doc, "Таблица 7.6 – SWOT-анализ системы пожарной безопасности")
    swot = [
        ["Сильные стороны", "Отечественный ПЛК ОВЕН, зонный пуск воды, свет и звук, открытая спецификация."],
        ["Слабые стороны", "Цена насоса. Charge Mate 2502 коммутирует аккумуляторную цепь и не заменяет пускатель двигателя около 100 кВт."],
        ["Возможности", "Повтор схем и программы на похожем объекте. Силовую часть можно заменить штатным контактором, не переписывая алгоритм."],
        ["Угрозы", "Рост цены насоса и импортного реле, более жёсткие требования к пуску электродвигателя."],
    ]
    add_table(doc, ["Фактор", "Содержание"], swot, col_widths=[4.2, 12.0], font=11)
    body(
        doc,
        f"Технико-экономическое обоснование даёт один итог. Проект стоит {money(TOTAL)} руб., "
        f"из них комплект на объекте с монтажом — {money(KIT)} руб. Основная сумма сидит в насосе, "
        f"автоматика добавляет {money(AUTO)} руб. и закрывает обнаружение, оповещение и пуск по зоне. "
        "Разработка для одного объекта дорога относительно мелких приборов, но при повторении схем на следующем объекте "
        "заново оплачивается уже не труд проекта, а комплект и монтаж.",
    )


def add_conclusion(doc):
    heading(doc, "ЗАКЛЮЧЕНИЕ", 1)
    body(
        doc,
        "В курсовом проекте разработана система пожарной безопасности на ПЛК ОВЕН 110. "
        "Цель работы — фиксация возгорания, объявление тревоги и пуск тушения — закрыта составом приборов и программой.",
    )
    body(doc, "Сделано следующее.", first=False, align="left")
    for item in [
        "разобраны функции охранно-пожарной сигнализации и типы извещателей;",
        "собрана структурная схема: ПЛК, Астра-10, Маяк-24-3М, извещатели, насос НЦПН-70/100 и пункт контроля;",
        "подобрана элементная база и расписаны характеристики приборов;",
        "построена принципиальная электрическая схема с питанием 12 В и 24 В, оповещением, шлейфами и четырьмя клапанами;",
        "построена функциональная схема автоматизации;",
        "по блок-схеме дописана программа на структурированном тексте CODESYS: опрос извещателей, тревога, клапан зоны и пуск насоса;",
        f"выполнено технико-экономическое обоснование: комплект {money(MAT)} руб., проект целиком {money(TOTAL)} руб., порог ущерба одного пожара {money(KIT)} руб.",
    ]:
        bullet(doc, item)
    body(
        doc,
        "Систему можно ставить как проектное решение обнаружения и зонного пуска. "
        "Перед реальным включением насоса силовое реле Charge Mate 2502 нужно заменить контактором и устройством плавного пуска: "
        "это ограничение отмечено в экономическом разделе и в стоимость комплекта не входило. "
        "Дальше схему можно нарастить шлейфами зон 2–4, не меняя уже написанный цикл программы.",
    )


def collapse_url_breaks(text):
    """Убирает пробелы, которые перенос строки вставил внутрь адреса."""

    def repl(match):
        url = re.sub(r"\s+", "", match.group(1))
        return url + match.group(2)

    return re.sub(r"(https?://.*?)(\s*\(дата|\s*$)", repl, text)


def add_bibliography(doc, pdf):
    heading(doc, "СПИСОК ЛИТЕРАТУРЫ", 1)
    lines = []
    for i in (36, 37):
        for row in iter_page_rows(pdf[i]):
            if row[0] != "line":
                continue
            text = row[1].strip()
            if not text or text == "СПИСОК ЛИТЕРАТУРЫ" or re.fullmatch(r"\d{1,3}", text):
                continue
            lines.append((row[2], text))
    items = []
    for x0, text in lines:
        if re.match(r"^\d+\.\s", text) or not items:
            items.append(text)
        else:
            items[-1] = items[-1] + " " + text
    for it in items:
        body(doc, collapse_url_breaks(it), first=False, align="left")
    extra = [
        "10. ПЛК110 [М02]. Каталог ОВЕН [Электронный ресурс]. URL: https://owen-energo.ru/product/plk110_m02/ (дата обращения: 23.11.2026).",
        "11. Оповещатель Астра-10 исп. 3 [Электронный ресурс]. URL: https://www.dssl.ru/products/astra-10-isp-3-opoveshchatel/ (дата обращения: 23.11.2026).",
        "12. Оповещатель Маяк-24-3М [Электронный ресурс]. URL: https://omelta.com/store/67-mayak_24_zm/ (дата обращения: 23.11.2026).",
        "13. Реле промежуточное РП slim 24 В [Электронный ресурс]. URL: https://xmart24.ru/catalog/EG000019/EC001437/rele-promezutocnoe-rp-slim-251-10a-24v-dc-averes-ekf-rps25124dc/ (дата обращения: 23.11.2026).",
        "14. Извещатель тепловой ИП 103-5/1 [Электронный ресурс]. URL: https://tinko-sb.ru/products/ip-103-5-1s-v-5376/ (дата обращения: 23.11.2026).",
        "15. Извещатель пламени ИПП-Ex (ИП 330-8) [Электронный ресурс]. URL: https://www.gradsb.ru/sistema-ops-vzryvoopasnyh-obyektov-quot-ladogaEx-quot-/23203.html (дата обращения: 23.11.2026).",
        "16. Насос НЦПН-70/100.Э24 [Электронный ресурс]. URL: https://pum-p.ru/catalog/nasosy/spetsialnye-nasosy/nasosy-ntspn/ (дата обращения: 23.11.2026).",
        "17. Насос НЦПН-70/100.М.Э24. Ливенский завод противопожарного машиностроения [Электронный ресурс]. URL: https://www.lzpm.ru/products/pro/nasos-pozharnyy-ntspn-70-100-m-e24/ (дата обращения: 23.11.2026).",
        "18. Charge Mate 2502 [Электронный ресурс]. URL: https://www.rvone.ru/catalog/elektrichestvo/charge-mate-2502/ (дата обращения: 23.11.2026).",
        "19. Клапан электромагнитный DN15, 220 В [Электронный ресурс]. URL: https://gremir.ru/klapan/elektromagnitnye/ (дата обращения: 23.11.2026).",
        "20. Ороситель СВН-10 [Электронный ресурс]. URL: https://sib-msk.ru/katalog/sprinklernye-orositeli/orositel-sprinklernyy-vodyanoy-obshchego-naznacheniya-svv-i-svn/ (дата обращения: 23.11.2026).",
    ]
    for it in extra:
        body(doc, it, first=False, align="left")


def add_appendices(doc, pdf):
    titles = {
        38: "ПРИЛОЖЕНИЕ А",
        39: "ПРИЛОЖЕНИЕ Б",
        40: "ПРИЛОЖЕНИЕ В",
    }
    notes = {
        38: "Графический материал структурной схемы.",
        39: "Графический материал принципиальной электрической схемы.",
        40: "Графический материал блок-схемы программы.",
    }
    for i in (38, 39, 40):
        doc.add_page_break()
        heading(doc, titles[i], 1)
        body(doc, notes[i], center=True, first=False)
        pix = pdf[i].get_pixmap(matrix=pymupdf.Matrix(1.6, 1.6), alpha=False)
        path = IMG / f"app_{i+1}.png"
        pix.save(path)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.add_run().add_picture(str(path), width=Cm(16.0))


def main():
    assert MAT > 0 and TOTAL > MAT
    print("MAT", MAT)
    print("ZP", ZP, "TAX", TAX_SUM, "ZP_FULL", ZP_FULL)
    print("MOUNT", MOUNT, "OTHER", OTHER, "TOTAL", TOTAL)
    print("KIT", KIT, "AUTO", AUTO, "MARGIN", MARGIN_VS_FACTORY)
    print("SHARES", SHARE_ROWS)

    pdf = pymupdf.open(PDF)
    doc = Document()
    setup_section(doc)
    add_title(doc)
    add_assignment(doc)
    add_contents(doc)
    flow = extract_body(pdf)
    emit_flow(doc, flow)
    add_software(doc)
    add_economics(doc)
    add_conclusion(doc)
    add_bibliography(doc, pdf)
    add_appendices(doc, pdf)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print("saved", OUT, OUT.stat().st_size)


if __name__ == "__main__":
    main()
