"""Презентации v2026.2: кейс → инфо → интерактивы → тест → письмо."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from content import (
    CASE,
    FORMULA,
    FORMULA_TRAPS,
    PATHS,
    PROGRAMS,
    SLIDE_SIMPLE,
    TYPES_KID,
    VERSION,
)
from ddo import answers_by_type, expected_sum, max_score
from pptx_kit import (
    AMBER,
    AMBER_D,
    CORAL,
    CREAM,
    INDIGO,
    INK,
    MUTED,
    NAVY,
    TEAL,
    TYPE_COLORS,
    WHITE,
    add_shape,
    chip,
    eyebrow,
    finish,
    new_blank,
    notes,
    picture,
    title,
    write,
)

W, H = 13.333, 7.5


def _prs(title_text: str):
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = title_text
    prs.core_properties.author = "Комплект «Путь в профессию», фестиваль «Наука 0+»"
    prs.core_properties.subject = f"Комплект {VERSION}"
    prs.core_properties.language = "ru-RU"
    return prs


def _mark(slide, text, color=CORAL):
    write(slide, 0.5, 0.22, 12, 0.28, [{"t": text.upper(), "size": 13, "color": color, "bold": True}])


def _info_badge(slide):
    chip(slide, 11.0, 0.2, 1.85, 0.38, "ИНФОРМАЦИЯ", TEAL, size=12)


def _int_badge(slide, minutes=""):
    label = f"ИНТЕРАКТИВ{('  ·  ' + minutes) if minutes else ''}"
    chip(slide, 10.35, 0.2, 2.5, 0.38, label, CORAL, size=12)


def _hero(prs, minutes: int, dark: set):
    s = new_blank(prs, NAVY)
    dark.add(len(prs.slides))
    picture(s, "hero-path.jpg", 0, 0, W, H, darken=0.28)
    add_shape(s, 0, 0, 7.2, H, NAVY)
    add_shape(s, 7.05, 0, 0.12, H, CORAL)
    write(s, 0.55, 1.35, 6.2, 0.35, [{"t": "ФЕСТИВАЛЬ  «НАУКА 0+»", "size": 15, "color": AMBER, "bold": True}])
    write(s, 0.55, 1.85, 6.3, 1.5, [{"t": "Путь в\nпрофессию", "size": 54, "color": WHITE, "bold": True, "spacing": 0.92}])
    write(s, 0.55, 3.7, 6.1, 0.9, [{"t": "Кейс, понятные вопросы\nи один шаг на месяц", "size": 20, "color": "D8E6EE"}])
    chip(s, 0.55, 5.0, 2.2, 0.48, f"{minutes} МИНУТ", TEAL)
    chip(s, 2.9, 5.0, 2.4, 0.48, "5–11 КЛАСС", "245066")
    notes(s, "Поздоровайтесь. Сегодня не «угадай профессию», а разберём выбор на простые части.")


def _flow(prs):
    s = new_blank(prs)
    picture(s, "hero-path.jpg", 8.5, 0, 4.85, H, darken=0.12)
    add_shape(s, 8.35, 0, 0.15, H, CORAL)
    eyebrow(s, "Как устроена встреча")
    title(s, "Порядок на сегодня", w=7.6)
    write(s, 0.5, 1.35, 7.6, 0.7, [{"t": SLIDE_SIMPLE["agenda"], "size": 16, "color": MUTED}])
    steps = [
        ("1", "Кейс", "Задача, которую будем решать"),
        ("2", "Два интерактива", "Разведка и формула «хочу — могу — умею — надо»"),
        ("3", "Тест Климова", "Типы профессий по интересу"),
        ("4", "Письмо себе", "Личный шаг на месяц"),
    ]
    for i, (n, head, body) in enumerate(steps):
        y = 2.2 + i * 1.1
        add_shape(s, 0.5, y, 7.6, 0.95, WHITE, rounded=True, radius=0.1)
        add_shape(s, 0.5, y, 0.9, 0.95, [CORAL, TEAL, INDIGO, AMBER_D][i])
        write(s, 0.5, y, 0.9, 0.95, [{"t": n, "size": 22, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, 1.65, y + 0.12, 6.2, 0.75, [{"t": head, "size": 18, "color": INK, "bold": True}, {"t": body, "size": 14, "color": MUTED, "before": 2}])
    notes(s, "Один проход. Между интерактивами будут короткие объяснения — не задания.")


def _case_open(prs):
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.15)
    add_shape(s, 0.45, 0.4, 12.4, 6.7, CREAM)
    _mark(s, "Кейс  ·  выдаём в начале")
    write(s, 0.75, 0.7, 11.5, 0.55, [{"t": CASE["title"], "size": 34, "color": NAVY, "bold": True}])
    write(s, 0.75, 1.35, 11.8, 1.15, [{"t": CASE["situation"], "size": 17, "color": INK}])
    write(s, 0.75, 2.6, 11.8, 0.4, [{"t": CASE["task"], "size": 17, "color": TEAL, "bold": True}])
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, (role, code, question) in enumerate(CASE["roles"]):
        x = 0.75 + i * 3.0
        add_shape(s, x, 3.25, 2.85, 3.45, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 3.25, 2.85, 0.14, accents[i])
        write(s, x + 0.15, 3.5, 2.55, 3.0, [
            {"t": role, "size": 14, "color": NAVY, "bold": True},
            {"t": code, "size": 12, "color": accents[i], "before": 2},
            {"t": question, "size": 13, "color": INK, "before": 10},
        ])
    notes(s, "Выдайте кейс и роли. Решение пока не собирайте: сначала будут инструменты. Критерии — в тетради.")


def _info_self(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "self-discovery.jpg", 7.85, 0.85, 4.95, 5.85, darken=0.05)
    add_shape(s, 7.85, 0.85, 4.95, 0.12, TEAL)
    eyebrow(s, "Перед первым заданием")
    title(s, "С чего начинается выбор", w=7.1)
    cards = [
        ("Почему сложно", SLIDE_SIMPLE["why_hard"]),
        ("Что такое интерес", SLIDE_SIMPLE["interest"]),
        ("Зачем это кейсу", "Центру нужно действие, которое помогает школьнику увидеть интерес, а не только «выбери уже профессию»."),
    ]
    for i, (head, body) in enumerate(cards):
        y = 1.5 + i * 1.65
        add_shape(s, 0.5, y, 7.05, 1.5, WHITE, rounded=True, radius=0.1)
        write(s, 0.75, y + 0.18, 6.6, 1.2, [{"t": head, "size": 18, "color": TEAL, "bold": True}, {"t": body, "size": 15, "color": INK, "before": 4}])
    notes(s, "Рассказ, не опрос. 2–3 минуты. Переход к разведке для кейса.")


def _interactive_recon(prs):
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.12)
    add_shape(s, 0.45, 0.45, 12.4, 6.6, CREAM)
    _int_badge(s, "4 МИН")
    write(s, 0.75, 0.7, 9, 0.3, [{"t": "ИНТЕРАКТИВ 1", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.75, 1.05, 11.5, 0.55, [{"t": CASE["interactive1"]["title"], "size": 32, "color": NAVY, "bold": True}])
    write(s, 0.75, 1.7, 11.8, 0.55, [{"t": CASE["interactive1"]["prompt"], "size": 17, "color": INK}])
    for i, q in enumerate(CASE["interactive1"]["questions"]):
        y = 2.5 + i * 1.25
        add_shape(s, 0.75, y, 11.8, 1.1, WHITE, rounded=True, radius=0.1)
        write(s, 1.0, y, 1.0, 1.1, [{"t": f"0{i+1}", "size": 22, "color": CORAL, "bold": True}], anchor="middle")
        write(s, 2.1, y, 10.1, 1.1, [{"t": q, "size": 20, "color": INK, "bold": True}], anchor="middle")
    notes(s, "4 минуты. Пишут по ролям в тетради. Вслух не обязательно. Это черновик решения кейса.")


def _info_world(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "labor-market.jpg", 0, 0.9, 6.15, 5.8, darken=0.05)
    add_shape(s, 0, 0.9, 6.15, 0.12, INDIGO)
    eyebrow(s, "Между интерактивами")
    title(s, "Что полезно знать про «надо»", x=6.45, w=6.4, size=26)
    write(s, 6.45, 1.45, 6.4, 1.0, [{"t": SLIDE_SIMPLE["market"], "size": 15, "color": INK}])
    for i, (head, body) in enumerate(PATHS):
        y = 2.65 + i * 1.25
        add_shape(s, 6.45, y, 6.4, 1.1, WHITE, rounded=True, radius=0.1)
        write(s, 6.7, y + 0.12, 5.95, 0.9, [{"t": head, "size": 16, "color": INDIGO, "bold": True}, {"t": body, "size": 12, "color": MUTED, "before": 2}])
    notes(s, "Коротко. Не уходите в ЕГЭ и лимиты заявлений. Старшим: детали в тетради. Младшим: сейчас важны пробы.")


def _interactive_formula(prs, minutes="5 МИН"):
    s = new_blank(prs)
    picture(s, "formula-visual.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.35, 0.3, 12.6, 6.85, CREAM)
    _int_badge(s, minutes)
    write(s, 0.65, 0.5, 9, 0.28, [{"t": "ИНТЕРАКТИВ 2", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.65, 0.85, 12, 0.5, [{"t": "Хочу · Могу · Умею · Надо", "size": 30, "color": NAVY, "bold": True}])
    write(s, 0.65, 1.4, 12.1, 0.45, [{"t": FORMULA_TRAPS + " Заполните свои клетки и отметьте, какую клетку закроет действие вашего центра.", "size": 14, "color": MUTED}])
    gap = 0.16
    width = (12.1 - 3 * gap) / 4
    x = 0.65
    for name, color, text in FORMULA:
        add_shape(s, x, 2.05, width, 3.75, color, rounded=True, radius=0.08)
        write(s, x + 0.16, 2.25, width - 0.32, 3.4, [
            {"t": name, "size": 22, "color": WHITE if color != AMBER else NAVY, "bold": True},
            {"t": text, "size": 13, "color": WHITE if color != AMBER else NAVY, "before": 8},
            {"t": "Себе: одна строка\nЦентру: ☐ помогает", "size": 12, "color": AMBER if color != AMBER else CORAL, "bold": True, "before": 14},
        ])
        x += width + gap
    notes(s, "5 минут. Личная формула + галочка: какую клетку закрывает идея центра. Это второй инструмент для кейса.")


def _info_klimov(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "five-types.jpg", 0, 0, W, 2.85, darken=0.2)
    add_shape(s, 0, 2.35, W, 0.55, CREAM)
    eyebrow(s, "Перед тестом", y=3.0)
    title(s, "Формула типов профессий", y=3.25, size=28)
    write(s, 0.5, 3.9, 12.3, 0.45, [{"t": SLIDE_SIMPLE["klimov"], "size": 15, "color": MUTED}])
    for i, (code, name, obj, examples) in enumerate(TYPES_KID):
        x = 0.4 + i * 2.55
        add_shape(s, x, 4.5, 2.4, 2.2, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 4.5, 2.4, 0.5, TYPE_COLORS[code], rounded=True, radius=0.1)
        add_shape(s, x, 4.8, 2.4, 0.2, TYPE_COLORS[code])
        write(s, x, 4.5, 2.4, 0.5, [{"t": f"{code}  {name}", "size": 14, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, x + 0.12, 5.15, 2.16, 1.4, [{"t": obj, "size": 13, "color": INK, "align": "center"}, {"t": examples, "size": 11, "color": MUTED, "before": 6, "align": "center"}])
    notes(s, "Объясните просто: тест спросит, с чем интереснее иметь дело. Границы типов бывают смешанными — это нормально.")


def _test(prs, express: bool, dark: set):
    s = new_blank(prs)
    picture(s, "self-discovery.jpg", 0, 0, W, H, darken=0.08)
    add_shape(s, 0.7, 1.15, 11.9, 5.15, CREAM)
    add_shape(s, 0.7, 1.15, 11.9, 0.16, CORAL)
    _int_badge(s, "5–6 МИН")
    write(s, 1.05, 1.55, 11, 0.3, [{"t": "ТЕСТ ДДО  ·  Е. А. КЛИМОВ", "size": 14, "color": CORAL, "bold": True}])
    if express:
        write(s, 1.05, 2.05, 11, 0.9, [{"t": "5 минут · вопросы 1–10", "size": 36, "color": NAVY, "bold": True}])
    else:
        write(s, 1.05, 2.05, 11, 0.9, [{"t": "6 минут · вопросы 1–15", "size": 36, "color": NAVY, "bold": True}])
    write(s, 1.05, 3.2, 11, 1.5, [{"t": SLIDE_SIMPLE["test"] + "\n\nВ каждой паре выберите одно: А или Б. Первый отклик. Ключ пока не открывайте.", "size": 18, "color": INK}])
    write(s, 1.05, 5.2, 11, 0.6, [{"t": "Непонятное слово — тихая рука. Баллы вслух не называем.", "size": 15, "color": TEAL, "bold": True}])
    notes(s, "Тишина. Не подсказывайте выборы.")


def _key(prs, express: bool):
    s = new_blank(prs)
    eyebrow(s, "Считаем")
    title(s, "Ключ теста", size=30)
    mx, sm = max_score(express), expected_sum(express)
    write(s, 0.5, 1.2, 12.3, 0.4, [{"t": f"Совпадение = 1 балл. Максимум в столбце — {mx}. Сумма всех баллов — {sm}. Смотрите на 1–2 самых высоких столбца.", "size": 15, "color": MUTED}])
    buckets = answers_by_type(express)
    gap = 0.14
    width = (12.33 - 4 * gap) / 5
    x = 0.5
    for code, items in buckets.items():
        add_shape(s, x, 1.75, width, 4.55, WHITE, rounded=True, radius=0.08)
        add_shape(s, x, 1.75, width, 0.7, TYPE_COLORS[code], rounded=True, radius=0.08)
        add_shape(s, x, 2.2, width, 0.25, TYPE_COLORS[code])
        name = {"П": "природа", "Т": "техника", "Ч": "человек", "З": "знаки", "Х": "образ"}[code]
        write(s, x, 1.75, width, 0.7, [{"t": f"{code}  {name}", "size": 14, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, x + 0.08, 2.65, width - 0.16, 3.4, [{"t": item, "size": 17, "color": INK, "bold": True, "align": "center", "after": 4} for item in items])
        x += width + gap
    notes(s, "Ноль — не запрет. Разница в 1 балл почти ничего не значит. Запишите ведущий тип рядом с «хочу».")


def _info_doors(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "social-directions.jpg", 0, 0, W, 2.55, darken=0.28)
    write(s, 0.5, 0.7, 12, 0.3, [{"t": "ПОСЛЕ ТЕСТА", "size": 13, "color": AMBER, "bold": True}])
    write(s, 0.5, 1.1, 12, 0.7, [{"t": "Куда может вести интерес к людям и обществу", "size": 26, "color": WHITE, "bold": True}])
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, p in enumerate(PROGRAMS):
        x = 0.4 + (i % 4) * 3.2
        add_shape(s, x, 2.85, 3.05, 2.35, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 2.85, 3.05, 0.12, accents[i])
        write(s, x + 0.15, 3.1, 2.75, 1.95, [
            {"t": p["title"], "size": 15, "color": NAVY, "bold": True},
            {"t": p["code"], "size": 11, "color": MUTED, "before": 2},
            {"t": p["question"], "size": 12, "color": accents[i], "before": 8},
        ])
    add_shape(s, 0.4, 5.4, 12.5, 1.3, "E7F3F1", rounded=True, radius=0.1)
    write(s, 0.65, 5.55, 12.0, 1.05, [
        {"t": "Поддержка, которой уже можно пользоваться", "size": 14, "color": TEAL, "bold": True},
        {"t": SLIDE_SIMPLE["support"] + " Подробности — в тетради, не на слайде.", "size": 14, "color": INK, "before": 4},
    ])
    notes(s, "Это пример, не каталог. Если тянет техника или природа — двери другие, и они не хуже. Не читайте ЕГЭ с экрана.")


def _case_finish(prs):
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.5, 1.1, 12.3, 5.2, CREAM)
    _int_badge(s, "3 МИН")
    write(s, 0.85, 1.4, 11, 0.3, [{"t": "ВОЗВРАТ К КЕЙСУ", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.85, 1.85, 11.5, 0.6, [{"t": "Соберите одно действие центра", "size": 30, "color": NAVY, "bold": True}])
    points = [
        "Возьмите черновик из разведки.",
        "Проверьте формулу: какую клетку закрываете для школьника?",
        "Учтите тип интересов: действие должно помогать разным ребятам, не только «самым громким».",
        "Запишите: название, для кого, первый шаг на этой неделе.",
    ]
    for i, line in enumerate(points):
        y = 2.7 + i * 0.75
        add_shape(s, 0.85, y, 11.6, 0.65, WHITE, rounded=True, radius=0.1)
        write(s, 1.1, y, 11.1, 0.65, [{"t": f"{i+1}.  {line}", "size": 16, "color": INK}], anchor="middle")
    notes(s, "3 минуты на запись. 1–2 команды вслух по желанию. Не устраивайте конкурс.")


def _letter(prs):
    s = new_blank(prs)
    picture(s, "future-letter.jpg", 0, 0, W, H, darken=0.08)
    add_shape(s, 0.7, 1.05, 11.9, 5.35, CREAM)
    add_shape(s, 0.7, 1.05, 11.9, 0.16, AMBER_D)
    _int_badge(s, "4 МИН")
    write(s, 1.0, 1.4, 11, 0.3, [{"t": "ПИСЬМО СЕБЕ", "size": 13, "color": AMBER_D, "bold": True}])
    write(s, 1.0, 1.8, 11, 0.55, [{"t": "Письмо себе из будущего", "size": 30, "color": NAVY, "bold": True}])
    lines = [
        "Мне ____ лет. Я занимаюсь…",
        "Мне по-настоящему нравится…",
        "Из школьного пригодилось…",
        "Одна фраза себе сегодняшнему…",
    ]
    for i, line in enumerate(lines):
        y = 2.6 + i * 0.8
        add_shape(s, 1.0, y, 11.3, 0.7, WHITE, rounded=True, radius=0.12)
        write(s, 1.25, y, 10.8, 0.7, [{"t": line, "size": 18, "color": INK}], anchor="middle")
    notes(s, "Не читайте и не собирайте. Для 5–6 класса хватит одной фразы.")


def _close(prs, dark: set):
    s = new_blank(prs, NAVY)
    dark.add(len(prs.slides))
    picture(s, "hero-path.jpg", 0, 0, W, H, darken=0.4)
    add_shape(s, 0.7, 1.5, 11.9, 4.5, NAVY)
    write(s, 1.0, 1.85, 11, 0.3, [{"t": "НА ВЫХОД", "size": 14, "color": AMBER, "bold": True}])
    write(s, 1.0, 2.3, 11, 0.7, [{"t": "Кейс. Формула. Тип интереса. Шаг.", "size": 28, "color": WHITE, "bold": True}])
    write(s, 1.0, 3.3, 11, 1.5, [{"t": "Мой шаг на этот месяц: …\n\nВопрос домой: какую пробу я сделаю?", "size": 20, "color": "D8E6EE"}])
    write(s, 1.0, 5.1, 11, 0.5, [{"t": "Тетрадь заберите. Спасибо.", "size": 16, "color": AMBER}])
    notes(s, "Короткий финал. Без новой темы.")


def build_hour():
    prs = _prs("Путь в профессию · 60 минут")
    dark: set[int] = set()
    _hero(prs, 60, dark)
    _flow(prs)
    _case_open(prs)
    _info_self(prs)
    _interactive_recon(prs)
    _info_world(prs)
    _interactive_formula(prs)
    _info_klimov(prs)
    _test(prs, express=False, dark=dark)
    _key(prs, express=False)
    _info_doors(prs)
    _case_finish(prs)
    _letter(prs)
    _close(prs, dark)
    return finish(prs, dark)


def build_short():
    prs = _prs("Путь в профессию · 20 минут")
    dark: set[int] = set()
    _hero(prs, 20, dark)

    # compact case
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.12)
    add_shape(s, 0.5, 0.7, 12.3, 6.1, CREAM)
    write(s, 0.85, 1.0, 11, 0.3, [{"t": "КЕЙС", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.85, 1.4, 11.5, 0.55, [{"t": CASE["title"], "size": 30, "color": NAVY, "bold": True}])
    write(s, 0.85, 2.1, 11.5, 1.3, [{"t": CASE["situation"], "size": 16, "color": INK}])
    write(s, 0.85, 3.5, 11.5, 0.5, [{"t": "Выберите одну роль. К концу встречи — одно действие.", "size": 17, "color": TEAL, "bold": True}])
    for i, (role, code, _) in enumerate(CASE["roles"]):
        x = 0.85 + i * 3.0
        add_shape(s, x, 4.3, 2.85, 2.0, WHITE, rounded=True, radius=0.1)
        write(s, x + 0.15, 4.3, 2.55, 2.0, [{"t": role, "size": 14, "color": NAVY, "bold": True, "align": "center"}, {"t": code, "size": 12, "color": MUTED, "align": "center", "before": 6}], anchor="middle")
    notes(s, "Кейс сразу. Решение соберём после инструментов.")

    # info short
    s = new_blank(prs)
    _info_badge(s)
    eyebrow(s, "Коротко")
    title(s, "Два простых вопроса")
    for i, (head, body) in enumerate([
        ("Что интересно делать?", SLIDE_SIMPLE["interest"]),
        ("Кому это может быть нужно?", SLIDE_SIMPLE["market"]),
    ]):
        y = 1.6 + i * 2.3
        add_shape(s, 0.5, y, 12.3, 2.1, WHITE, rounded=True, radius=0.1)
        write(s, 0.85, y + 0.35, 11.7, 1.5, [{"t": head, "size": 24, "color": NAVY, "bold": True}, {"t": body, "size": 17, "color": INK, "before": 8}])
    notes(s, "1.5 минуты рассказа.")

    # interactive 1 mini
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.55, 1.2, 12.2, 5.1, CREAM)
    _int_badge(s, "2 МИН")
    write(s, 0.9, 1.55, 11, 0.3, [{"t": "ИНТЕРАКТИВ 1", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.9, 2.0, 11.5, 0.7, [{"t": "Одно действие для вашего героя", "size": 28, "color": NAVY, "bold": True}])
    write(s, 0.9, 2.9, 11.5, 2.8, [
        {"t": "По выбранной роли напишите:", "size": 18, "color": INK},
        {"t": "1. Кому помогаете", "size": 20, "color": NAVY, "bold": True, "before": 14},
        {"t": "2. Какое одно действие предлагаете", "size": 20, "color": NAVY, "bold": True, "before": 10},
        {"t": "Это черновик. Потом уточните формулой.", "size": 16, "color": MUTED, "before": 14},
    ])
    notes(s, "2 минуты. Без озвучки.")

    _interactive_formula(prs, minutes="3 МИН")
    _info_klimov(prs)
    _test(prs, express=True, dark=dark)
    _key(prs, express=True)

    # doors compressed + letter + close
    s = new_blank(prs)
    _info_badge(s)
    eyebrow(s, "Если близки люди и общество")
    title(s, "Четыре направления — пример", size=28)
    for i, p in enumerate(PROGRAMS):
        y = 1.45 + i * 1.2
        add_shape(s, 0.5, y, 12.3, 1.05, WHITE, rounded=True, radius=0.1)
        write(s, 0.75, y, 4.8, 1.05, [{"t": f"{p['title']}  {p['code']}", "size": 16, "color": NAVY, "bold": True}], anchor="middle")
        write(s, 5.6, y, 6.9, 1.05, [{"t": p["question"], "size": 15, "color": INDIGO}], anchor="middle")
    notes(s, "Четыре вопроса. Поддержка: Билет в будущее, Профессионалитет, целевое с родителями — одной фразой.")

    _letter(prs)
    _close(prs, dark)
    return finish(prs, dark)


def main():
    hour = build_hour()
    short = build_short()
    hour.save("/workspace/materials/Путь_в_профессию_презентация_60_минут.pptx")
    short.save("/workspace/materials/Путь_в_профессию_презентация_20_минут.pptx")
    print(f"hour: {len(hour.slides)}")
    print(f"short: {len(short.slides)}")


if __name__ == "__main__":
    main()
