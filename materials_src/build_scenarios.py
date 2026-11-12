"""Сценарий экспресс 20 минут: формула, направления, опрос по слайдам, типы после ключа."""

from __future__ import annotations

from content import SOURCES, VERSION
from ddo import DROPPED_CLASSIC, QUESTIONS, TYPE_ORDER, TYPE_SHORT, answers_by_type, expected_sum, max_score
from docx_kit import (
    INK,
    MUTED,
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
INK_DEEP = "14202B"


def build_all():
    path = _build(
        "/workspace/materials/Путь_в_профессию_сценарий_20_минут.docx",
        "Экспресс 20 минут",
        BLOCKS,
        26,
    )
    alias = "/workspace/materials/Путь_в_профессию_сценарий_60_минут.docx"
    _build(alias, "Экспресс 20 минут (основной формат комплекта)", BLOCKS, 26)
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
    p(doc, "Путь в профессию", size=26, color=INK_DEEP, bold=True, after=2, line=1.0)
    p(doc, subtitle, size=13, after=4)
    p(doc, f"Сценарий  ·  20 минут  ·  5–11 класс  ·  комплект {VERSION}", size=11, color=MUTED, after=8)
    rows = [
        ("Формат", "20 минут. Раздатка — одна страница: формула + 15 пар Климова."),
        ("Опрос", "Вопросы показывают на слайдах. Ответ отмечают в раздатке."),
        ("Типы", "Подробный разбор пяти типов — после ключа, не до опроса."),
        ("Слайды", f"{slides} слайдов: формула, подробные направления, 15 вопросов со стикерами, ключ, типы."),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    set_table_widths(table, [3.2, W - 3.2])
    set_table_borders(table, "D2DBE3", "4")
    for i, (k, v) in enumerate(rows):
        left, right = table.rows[i].cells
        shade_cell(left, INK_DEEP)
        shade_cell(right, "F7FAFB" if i % 2 == 0 else "FFFFFF")
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
    rich(doc, [("Слайды. ", {"bold": True, "color": TEAL}), (block["slide"], {"color": INK})], after=2)
    rich(doc, [("Задача. ", {"bold": True, "color": INK_DEEP}), (block["goal"], {})], after=3)
    if block.get("interactive"):
        callout(doc, "ИНТЕРАКТИВ", block["interactive"], "FDECEA", "E05A45")
    if block.get("info"):
        callout(doc, "ИНФОРМАЦИЯ", block["info"], "E6F5F2", "0B8A7A")
    for para in block["say"]:
        p(doc, para, size=11, after=3)
    if block.get("cut"):
        p(doc, "Если не успеваете: " + block["cut"], size=11, color="C47A1A")


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
        "goal": "Задать смысл пути и карту встречи.",
        "say": [
            "Сегодня не угадываем должность. Формула выбора, направления, опросник по слайдам. Типы разберём после подсчёта.",
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
        "time": "0:06–0:10",
        "title": "Направления подготовки",
        "slide": "4–6",
        "goal": "Подробно рассказать о четырёх направлениях.",
        "info": "Слайд 4 — обзор. Слайды 5–6 — карточки: чему учат, что изучают, куда идут, кому близко.",
        "say": [
            "Если интересны люди и общество — четыре пути. На каждом подробном слайде держитесь вопроса направления и пары ключевых фраз.",
        ],
        "cut": "Обзор + по одной минуте на пару направлений.",
    },
    {
        "time": "0:10–0:16",
        "title": "Опросник по слайдам",
        "slide": "7–22",
        "goal": "Пройти 15 пар. Ответ — в раздатке.",
        "interactive": "Слайд 7 — правило. Слайды 8–22 — пары 1–15. По 20–25 секунд.",
        "say": [
            "Листайте спокойно. Без обсуждения вслух. Стикеры на слайдах — для настроения, не для подсказки ответа.",
        ],
        "cut": "Держите темп: не комментируйте пары.",
    },
    {
        "time": "0:16–0:18",
        "title": "Ключ",
        "slide": "23",
        "goal": "Посчитать баллы по пяти столбцам.",
        "interactive": "Ключ на слайде. В раздатке — ведущий и второй тип.",
        "say": [
            "Смотрите на один-два верхних столбца. Разница в один балл почти ничего не значит.",
        ],
    },
    {
        "time": "0:18–0:20",
        "title": "Типы подробно и выход",
        "slide": "24–26",
        "goal": "Объяснить смысл типов и связать с направлениями; назвать шаг.",
        "info": "Пять типов по предмету труда + границы типов.",
        "say": [
            "Теперь, когда столбец уже есть, расскажите, что он означает. Свяжите «человек» и «знаки» с направлениями.",
            "Один шаг на месяц. Спасибо.",
        ],
    },
]


if __name__ == "__main__":
    print("\n".join(build_all()))
