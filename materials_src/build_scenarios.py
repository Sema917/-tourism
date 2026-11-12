"""Сценарий экспресс 20 минут: формула, типы, направления, опросник."""

from __future__ import annotations

from content import SOURCES, VERSION
from ddo import DROPPED_CLASSIC, QUESTIONS, TYPE_ORDER, TYPE_SHORT, answers_by_type, expected_sum, max_score
from docx_kit import (
    INK,
    MUTED,
    NAVY,
    TEAL,
    add_text,
    bullet,
    callout,
    h1,
    h2,
    new_doc,
    p,
    page_break,
    rich,
    set_cell_margins,
    set_table_borders,
    set_table_widths,
    shade_cell,
    spacing,
)

W = 18.0
BURGUNDY = "6B1E2A"


def build_all():
    path = _build(
        "/workspace/materials/Путь_в_профессию_сценарий_20_минут.docx",
        "Экспресс 20 минут",
        BLOCKS,
        8,
    )
    # для совместимости прежнего комплекта — тот же сценарий
    alias = "/workspace/materials/Путь_в_профессию_сценарий_60_минут.docx"
    _build(alias, "Экспресс 20 минут (основной формат комплекта)", BLOCKS, 8)
    return path, alias


def _build(filename, subtitle, blocks, slides):
    doc = new_doc(f"Наука 0+   ·   Путь в профессию   ·   сценарий 20 минут   ·   {VERSION}")
    doc.core_properties.title = "Путь в профессию. Сценарий на 20 минут"
    doc.core_properties.subject = subtitle
    _passport(doc, subtitle, slides)
    h1(doc, "Ход встречи")
    for block in blocks:
        _block(doc, block)
    page_break(doc)
    _appendix(doc)
    doc.save(filename)
    return filename


def _passport(doc, subtitle, slides):
    p(doc, "ФЕСТИВАЛЬ  «НАУКА 0+»", size=12, color=TEAL, bold=True, after=2)
    p(doc, "Путь в профессию", size=26, color=BURGUNDY, bold=True, after=2, line=1.0)
    p(doc, subtitle, size=13, after=4)
    p(doc, f"Сценарий  ·  20 минут  ·  5–11 класс  ·  комплект {VERSION}", size=11, color=MUTED, after=8)
    rows = [
        ("Формат", "Только 20 минут. Раздатка — одна страница: формула + 15 пар Климова."),
        ("Оставить", "Смысл пути в профессию, формула «хочу–могу–умею–надо», типы, направления подготовки, опросник."),
        ("Слайды", f"{slides} слайдов в стиле квиза (бордо / крем / золото)."),
        ("Раздатка", "Только графы формулы и опросник. Без кейса, письма и длинных справок."),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    set_table_widths(table, [3.2, W - 3.2])
    set_table_borders(table, "E4DCD0", "4")
    for i, (k, v) in enumerate(rows):
        left, right = table.rows[i].cells
        shade_cell(left, BURGUNDY)
        shade_cell(right, "F7F4EE" if i % 2 == 0 else "FFFFFF")
        set_cell_margins(left, 50, 50, 70, 70)
        set_cell_margins(right, 50, 50, 80, 80)
        _cell(left, k, 10, "FFFFFF", True)
        _cell(right, v, 10.5, INK, False)


def _cell(cell, text, size, color, bold):
    paragraph = cell.paragraphs[0]
    spacing(paragraph, 0, 0, 1.02)
    add_text(paragraph, text, size, color, bold)


def _block(doc, block):
    h2(doc, f"{block['time']}   {block['title']}")
    rich(doc, [("Слайд. ", {"bold": True, "color": TEAL}), (block["slide"], {"color": INK})], after=2)
    rich(doc, [("Задача. ", {"bold": True, "color": BURGUNDY}), (block["goal"], {})], after=3)
    if block.get("interactive"):
        callout(doc, "ИНТЕРАКТИВ", block["interactive"], "FDE8E2", "E15A3A")
    if block.get("info"):
        callout(doc, "ИНФОРМАЦИЯ", block["info"], "E7F3F1", "0F7A6C")
    for para in block["say"]:
        p(doc, para, size=11, after=3)
    if block.get("cut"):
        p(doc, "Если не успеваете: " + block["cut"], size=11, color="A86B12")


def _appendix(doc):
    h1(doc, "Приложение")
    h2(doc, "Ключ ДДО · 15 пар")
    for code in TYPE_ORDER:
        bullet(doc, f"{code} · {TYPE_SHORT[code]}: {', '.join(answers_by_type(False)[code])}")
    p(doc, f"Сумма {expected_sum(False)}. Максимум столбца {max_score(False)}.")
    h2(doc, "Сокращение")
    p(doc, "Не вошли классические пары " + ", ".join(DROPPED_CLASSIC) + ".")
    for q in QUESTIONS:
        note = " Адаптация." if q["note"] else ""
        bullet(doc, f"{q['n']} ← классический {q['classic']}. А={q['a_type']}, Б={q['b_type']}.{note}")
    h2(doc, "Источники")
    for item in SOURCES:
        bullet(doc, item)


BLOCKS = [
    {
        "time": "0:00–0:02",
        "title": "Старт",
        "slide": "1–2",
        "goal": "Задать смысл «пути в профессию» и карту 20 минут.",
        "say": [
            "Сегодня не угадываем должность. Берём формулу выбора, типы Климова, четыре направления и опросник.",
            "Раздатка одна страница: графы формулы и 15 пар. Тетрадь не сдаём.",
        ],
    },
    {
        "time": "0:02–0:06",
        "title": "Формула выбора",
        "slide": "3",
        "goal": "Заполнить хочу — могу — умею — надо.",
        "interactive": "4 минуты: графы в раздатке.",
        "say": [
            "Только «хочу» — мечта. Только «надо» — чужой выбор. Нужно пересечение клеток.",
        ],
    },
    {
        "time": "0:06–0:09",
        "title": "Типы и направления",
        "slide": "4–5",
        "goal": "Показать предмет труда и четыре направления подготовки.",
        "info": "Социология, социальная работа, политология, организация работы с молодёжью.",
        "say": [
            "Климов: природа, техника, человек, знаки, образ.",
            "Если интересны люди и общество — четыре направления. Назовите вопрос каждого направления, не читайте длинно.",
        ],
        "cut": "Типы — одной минутой, направления — двумя.",
    },
    {
        "time": "0:09–0:16",
        "title": "Опросник Климова",
        "slide": "6",
        "goal": "Пройти 15 пар.",
        "interactive": "7 минут: А или Б в раздатке. Ключ закрыт.",
        "say": [
            "Это гипотеза интереса, не приговор и не назначение профессии.",
        ],
    },
    {
        "time": "0:16–0:20",
        "title": "Ключ и выход",
        "slide": "7–8",
        "goal": "Посчитать, связать с формулой и направлениями, назвать шаг.",
        "interactive": "Ключ на слайде. Внизу раздатки — ведущий тип и шаг на месяц.",
        "say": [
            "Смотрите на один-два верхних столбца. Если рядом человек/знаки — вспомните направления.",
            "Один шаг на месяц: проба, разговор или страница «Билета в будущее». Спасибо.",
        ],
    },
]


if __name__ == "__main__":
    print("\n".join(build_all()))
