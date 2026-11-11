"""Сценарии v2026.3: кейс про Леру; формула профессии ≠ тест; направления обязательно."""

from __future__ import annotations

from content import CASE_SAMPLES, SOURCES, VERSION
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


def build_all():
    return (
        _build(60, "/workspace/materials/Путь_в_профессию_сценарий_60_минут.docx", "Сюжет на час: кейс про Леру", BLOCKS_60, 17, False),
        _build(20, "/workspace/materials/Путь_в_профессию_сценарий_20_минут.docx", "Экспресс: кейс про Леру", BLOCKS_20, 12, True),
    )


def _build(minutes, filename, subtitle, blocks, slides, express):
    doc = new_doc(f"Наука 0+   ·   Путь в профессию   ·   сценарий {minutes} минут   ·   комплект {VERSION}")
    doc.core_properties.title = f"Путь в профессию. Сценарий на {minutes} минут"
    doc.core_properties.subject = subtitle
    _passport(doc, minutes, subtitle, slides, express)
    h1(doc, "Главное отличие инструментов")
    bullet(doc, "«Хочу — могу — умею — надо» — формула выбора человека.")
    bullet(doc, "Формула профессии по Климову — как описать работу: предмет, цели, средства, условия + пять типов по предмету труда.")
    bullet(doc, "Тест ДДО — отдельный шаг: к какому типу дел тянет именно участник.")
    bullet(doc, "Четыре направления подготовки — обязательный информационный блок для решения кейса.")
    h1(doc, "Ход встречи")
    for block in blocks:
        _block(doc, block)
    page_break(doc)
    _appendix(doc, express)
    doc.save(filename)
    return filename


def _passport(doc, minutes, subtitle, slides, express):
    p(doc, "ФЕСТИВАЛЬ  «НАУКА 0+»", size=12, color=TEAL, bold=True, after=2)
    p(doc, "Путь в профессию", size=28, color=NAVY, bold=True, after=2, line=1.0)
    p(doc, subtitle, size=14, after=4)
    p(doc, f"Сценарий  ·  {minutes} минут  ·  5–11 класс  ·  комплект {VERSION}", size=11, color=MUTED, after=8)
    rows = [
        ("Кейс", "«Помогите Лере выбрать». Выдаём в начале, решаем в конце по четырём пунктам."),
        ("Интерактивы", "1) Сбор фактов о Лере. 2) Хочу — могу — умею — надо. Затем решение кейса и письмо."),
        ("Информация", "Между заданиями: смысл инструментов, рынок/маршруты, формула профессии, типы, направления, поддержка."),
        ("Метод", "ДДО отдельно от формулы профессии. " + ("Вопросы 1–10." if express else "Вопросы 1–15.")),
        ("Слайды", f"{slides} слайдов."),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    set_table_widths(table, [3.2, W - 3.2])
    set_table_borders(table, "E4DCD0", "4")
    for i, (k, v) in enumerate(rows):
        left, right = table.rows[i].cells
        shade_cell(left, "16324A")
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
    rich(doc, [("Слайды. ", {"bold": True, "color": TEAL}), (block["slides"], {"color": INK})], after=2)
    rich(doc, [("Задача. ", {"bold": True, "color": NAVY}), (block["goal"], {})], after=3)
    if block.get("interactive"):
        callout(doc, "ИНТЕРАКТИВ", block["interactive"], "FDE8E2", "E15A3A")
    if block.get("info"):
        callout(doc, "ИНФОРМАЦИЯ", block["info"], "E7F3F1", "0F7A6C")
    for para in block["say"]:
        p(doc, para, size=11, after=4)
    if block.get("cut"):
        p(doc, "Если не успеваете: " + block["cut"], size=11, color="A86B12")


def _appendix(doc, express):
    h1(doc, "Приложение")
    h2(doc, "Ключ ДДО")
    for title, flag in (("Полная 1–15", False), ("Экспресс 1–10", True)):
        p(doc, title, size=12, color=NAVY, bold=True, before=4, after=2)
        for code in TYPE_ORDER:
            bullet(doc, f"{code} · {TYPE_SHORT[code]}: {', '.join(answers_by_type(flag)[code])}")
    p(doc, f"Сумма {expected_sum(False)}/{expected_sum(True)}. Максимум столбца {max_score(False)}/{max_score(True)}.")
    h2(doc, "Сокращение")
    p(doc, "Не вошли классические пары " + ", ".join(DROPPED_CLASSIC) + ".")
    for q in QUESTIONS:
        note = " Адаптация." if q["note"] else ""
        bullet(doc, f"{q['n']} ← классический {q['classic']}. А={q['a_type']}, Б={q['b_type']}.{note}")
    h2(doc, "Пример решения кейса")
    for title, text in CASE_SAMPLES:
        p(doc, title, size=11, color=NAVY, bold=True, before=4, after=1)
        p(doc, text, size=11)
    h2(doc, "Источники")
    for item in SOURCES:
        bullet(doc, item)


BLOCKS_60 = [
    {
        "time": "0:00–0:04",
        "title": "Старт",
        "slides": "1–2",
        "goal": "Показать карту инструментов.",
        "say": [
            "Сегодня мы не угадываем профессию. Мы берём инструменты и в конце помогаем Лере выбрать следующий шаг.",
            "Важно: формула «хочу — могу — умею — надо», формула профессии по Климову и тест ДДО — это разные вещи. Все три понадобятся.",
        ],
    },
    {
        "time": "0:04–0:08",
        "title": "Кейс",
        "slides": "3",
        "goal": "Выдать задачу Леры. Не решать.",
        "say": [
            "Кейс. Лера из 9 класса запуталась между людьми, пользой и названиями направлений. К концу заполним карточку из четырёх пунктов: формула выбора, тип Климова, направление подготовки, шаг на месяц.",
        ],
    },
    {
        "time": "0:08–0:12",
        "title": "Информация + интерактив 1",
        "slides": "4–5",
        "goal": "Собрать факты о Лере.",
        "info": "Четыре инструмента встречи.",
        "interactive": "4 минуты: выписать хочу/могу/умею/надо Леры из текста.",
        "say": [
            "Сначала карта инструментов. Затем интерактив один: только факты из истории Леры, без вывода о вузе.",
        ],
    },
    {
        "time": "0:12–0:20",
        "title": "Информация + интерактив 2",
        "slides": "6–7",
        "goal": "Формула выбора для себя и для Леры.",
        "info": "Клетка «надо» и маршруты по возрастам.",
        "interactive": "5 минут: хочу — могу — умею — надо.",
        "say": [
            "Лера в 9 классе — время развилки. Колледж тоже путь.",
            "Интерактив два — формула выбора. Это не формула профессии Климова и не тест.",
        ],
    },
    {
        "time": "0:20–0:28",
        "title": "Формула профессии и типы",
        "slides": "8–9",
        "goal": "Отделить описание профессии от теста.",
        "info": "Предмет, цели, средства, условия. Пять типов по предмету труда.",
        "say": [
            "Формула профессии отвечает: из чего состоит работа. По предмету труда Климов выделяет природу, технику, человека, знаки, образ.",
            "Для Леры пока только подумайте, какой тип звучит ближе. Тест будет следующим и отдельным шагом.",
        ],
    },
    {
        "time": "0:28–0:38",
        "title": "Тест ДДО",
        "slides": "10–11",
        "goal": "Провести тест отдельно.",
        "interactive": "6 минут: вопросы 1–15, затем ключ.",
        "say": [
            "Тест ДДО не описывает профессию целиком. Он показывает, к какому типу дел вас тянет. Гипотеза, не приговор.",
        ],
    },
    {
        "time": "0:38–0:48",
        "title": "Направления подготовки + поддержка",
        "slides": "12–14",
        "goal": "Обязательно рассказать о четырёх направлениях.",
        "info": "Социология, социальная работа, политология, организация работы с молодёжью.",
        "say": [
            "Это не реклама факультета, но блок обязательный: без него кейс не решить.",
            "Социология — что происходит с людьми на самом деле. Социальная работа — как организовать помощь. Политология — кто и по каким правилам решает. Организация работы с молодёжью — как дать молодым дело и среду.",
            "Для шага Леры назовите простые двери поддержки: «Билет в будущее», «Профессионалитет», целевое только с родителями.",
        ],
        "cut": "Поддержку — одной минутой.",
    },
    {
        "time": "0:48–0:55",
        "title": "Решение кейса",
        "slides": "15",
        "goal": "Собрать ответ Лере по 4 пунктам.",
        "interactive": "6 минут запись + 1–2 голоса.",
        "say": [
            "Заполните карточку: формула выбора Леры, тип Климова, направление, шаг. Можно опираться на свой тест только как на подсказку логики, не как на копию Леры.",
        ],
    },
    {
        "time": "0:55–1:00",
        "title": "Письмо и финал",
        "slides": "16–17",
        "goal": "Личный шаг.",
        "interactive": "Письмо себе.",
        "say": [
            "Письмо себе из будущего. Тетрадь заберите. Спасибо.",
        ],
    },
]


BLOCKS_20 = [
    {
        "time": "0:00–0:03",
        "title": "Кейс и инструменты",
        "slides": "1–3",
        "goal": "Выдать Леру и четыре инструмента.",
        "say": [
            "Кейс про Леру. В конце — 4 пункта ответа. Формула выбора, формула профессии, тест и направления — разные инструменты.",
        ],
    },
    {
        "time": "0:03–0:08",
        "title": "Два интерактива",
        "slides": "4–5",
        "goal": "Факты о Лере и формула выбора.",
        "interactive": "Сбор фактов + хочу/могу/умею/надо.",
        "say": ["Сначала факты, потом формула выбора для себя и для Леры."],
    },
    {
        "time": "0:08–0:14",
        "title": "Формула профессии и тест",
        "slides": "6–8",
        "goal": "Не смешивать.",
        "info": "Формула профессии: предмет, цели, средства, условия.",
        "interactive": "Экспресс-тест 1–10.",
        "say": [
            "Сначала формула профессии. Потом отдельно тест.",
        ],
    },
    {
        "time": "0:14–0:20",
        "title": "Направления, решение, письмо",
        "slides": "9–12",
        "goal": "Направления обязательны, затем решение кейса.",
        "info": "Четыре направления.",
        "interactive": "Карточка Леры + начало письма.",
        "say": [
            "Коротко по четырём направлениям. Затем карточка ответа Лере. Письмо себе.",
        ],
    },
]


if __name__ == "__main__":
    print("\n".join(build_all()))
