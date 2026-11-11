"""Презентации v2026.3: кейс про Леру, раздельные формула профессии и тест, направления."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from content import (
    CASE,
    DIRECTIONS_SLIDE,
    FORMULA,
    FORMULA_TRAPS,
    PATHS,
    PROFESSION_FORMULA,
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
    write(s, 0.55, 3.7, 6.1, 1.0, [{"t": "Инструменты выбора\nи кейс, который решим в конце", "size": 20, "color": "D8E6EE"}])
    chip(s, 0.55, 5.1, 2.2, 0.48, f"{minutes} МИНУТ", TEAL)
    chip(s, 2.9, 5.1, 2.4, 0.48, "5–11 КЛАСС", "245066")
    notes(s, "Скажите: сегодня изучим инструменты и в конце применим их к кейсу про Леру.")


def _flow(prs):
    s = new_blank(prs)
    picture(s, "hero-path.jpg", 8.5, 0, 4.85, H, darken=0.12)
    add_shape(s, 8.35, 0, 0.15, H, CORAL)
    eyebrow(s, "Как устроена встреча")
    title(s, "Что изучим и зачем", w=7.6)
    write(s, 0.5, 1.3, 7.6, 0.8, [{"t": SLIDE_SIMPLE["agenda"], "size": 15, "color": MUTED}])
    steps = [
        ("1", "Кейс про Леру", "Задача, которую решим в конце"),
        ("2", "Два интерактива", "Факты о Лере + «хочу — могу — умею — надо»"),
        ("3", "Формула профессии", "Как описать работу по Климову — отдельно от теста"),
        ("4", "Тест ДДО", "Отдельный шаг: ваш тип интереса"),
        ("5", "Направления", "Социология, соцработа, политология, работа с молодёжью"),
        ("6", "Решение + письмо", "Собираем ответ для Леры"),
    ]
    for i, (n, head, body) in enumerate(steps):
        y = 2.15 + i * 0.75
        add_shape(s, 0.5, y, 7.6, 0.68, WHITE, rounded=True, radius=0.1)
        write(s, 0.7, y, 0.6, 0.68, [{"t": n, "size": 18, "color": CORAL, "bold": True}], anchor="middle")
        write(s, 1.4, y, 6.5, 0.68, [{"t": head, "size": 15, "color": INK, "bold": True}, {"t": "  ·  " + body, "size": 13, "color": MUTED}], anchor="middle")
    notes(s, "Подчеркните: формула профессии ≠ тест Климова. Оба инструмента понадобятся для кейса.")


def _case_open(prs):
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.12)
    add_shape(s, 0.4, 0.35, 12.5, 6.8, CREAM)
    write(s, 0.7, 0.55, 12, 0.28, [{"t": "КЕЙС  ·  ВЫДАЁМ СЕЙЧАС, РЕШАЕМ В КОНЦЕ", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.7, 0.95, 12, 0.55, [{"t": CASE["title"], "size": 32, "color": NAVY, "bold": True}])
    write(s, 0.7, 1.55, 12.0, 1.7, [{"t": CASE["situation"], "size": 16, "color": INK}])
    write(s, 0.7, 3.4, 12.0, 0.45, [{"t": CASE["task"], "size": 16, "color": TEAL, "bold": True}])
    for i, step in enumerate(CASE["solve_steps"]):
        y = 4.0 + i * 0.7
        add_shape(s, 0.7, y, 12.0, 0.6, WHITE, rounded=True, radius=0.1)
        write(s, 0.95, y, 11.5, 0.6, [{"t": f"{i+1}.  {step}", "size": 15, "color": INK}], anchor="middle")
    notes(s, "Не решайте кейс сейчас. Скажите: все следующие блоки — инструменты для ответа Лере.")


def _info_self(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "self-discovery.jpg", 7.85, 0.85, 4.95, 5.85, darken=0.05)
    add_shape(s, 7.85, 0.85, 4.95, 0.12, TEAL)
    eyebrow(s, "Перед первым заданием")
    title(s, "Какие инструменты возьмём", w=7.1, size=28)
    cards = [
        ("Формула выбора", "Хочу — могу — умею — надо. Про Леру и про себя."),
        ("Формула профессии", "Как устроена работа: предмет, цели, средства, условия."),
        ("Тест Климова", "Отдельно: к какому типу дел тянет."),
        ("4 направления", "Социология, соцработа, политология, работа с молодёжью."),
    ]
    for i, (head, body) in enumerate(cards):
        y = 1.45 + i * 1.25
        add_shape(s, 0.5, y, 7.05, 1.1, WHITE, rounded=True, radius=0.1)
        write(s, 0.75, y + 0.15, 6.6, 0.85, [{"t": head, "size": 17, "color": TEAL, "bold": True}, {"t": body, "size": 14, "color": INK, "before": 3}])
    notes(s, "2 минуты. Это карта инструментов для кейса.")


def _interactive_facts(prs):
    s = new_blank(prs)
    picture(s, "self-discovery.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.45, 0.45, 12.4, 6.6, CREAM)
    _int_badge(s, "4 МИН")
    write(s, 0.75, 0.7, 9, 0.28, [{"t": "ИНТЕРАКТИВ 1", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.75, 1.05, 11.5, 0.5, [{"t": CASE["interactive1"]["title"], "size": 30, "color": NAVY, "bold": True}])
    write(s, 0.75, 1.65, 11.8, 0.5, [{"t": CASE["interactive1"]["prompt"], "size": 16, "color": INK}])
    colors = [CORAL, TEAL, AMBER_D, INDIGO]
    for i, q in enumerate(CASE["interactive1"]["questions"]):
        col = i % 2
        row = i // 2
        x = 0.75 + col * 6.05
        y = 2.4 + row * 2.1
        add_shape(s, x, y, 5.85, 1.9, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, y, 5.85, 0.14, colors[i])
        write(s, x + 0.25, y + 0.35, 5.35, 1.3, [{"t": q, "size": 18, "color": NAVY, "bold": True}])
    notes(s, "4 минуты. Только факты из текста кейса. Выводов о направлении пока не делаем.")


def _info_world(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "labor-market.jpg", 0, 0.9, 6.15, 5.8, darken=0.05)
    add_shape(s, 0, 0.9, 6.15, 0.12, INDIGO)
    eyebrow(s, "Между интерактивами")
    title(s, "Зачем Лере клетка «надо»", x=6.45, w=6.4, size=26)
    write(s, 6.45, 1.45, 6.4, 1.15, [{"t": SLIDE_SIMPLE["market"], "size": 15, "color": INK}])
    for i, (head, body) in enumerate(PATHS):
        y = 2.8 + i * 1.2
        add_shape(s, 6.45, y, 6.4, 1.05, WHITE, rounded=True, radius=0.1)
        write(s, 6.7, y + 0.12, 5.95, 0.85, [{"t": head, "size": 15, "color": INDIGO, "bold": True}, {"t": body, "size": 12, "color": MUTED, "before": 2}])
    notes(s, "Лера в 9 классе — как раз развилка. Колледж тоже нормальный путь.")


def _interactive_want(prs, minutes="5 МИН"):
    s = new_blank(prs)
    picture(s, "formula-visual.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.35, 0.3, 12.6, 6.85, CREAM)
    _int_badge(s, minutes)
    write(s, 0.65, 0.5, 9, 0.28, [{"t": "ИНТЕРАКТИВ 2", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.65, 0.85, 12, 0.5, [{"t": "Хочу · Могу · Умею · Надо", "size": 30, "color": NAVY, "bold": True}])
    write(s, 0.65, 1.4, 12.1, 0.45, [{"t": FORMULA_TRAPS + " Сначала для себя коротко, затем уточните карточку Леры.", "size": 14, "color": MUTED}])
    gap = 0.16
    width = (12.1 - 3 * gap) / 4
    x = 0.65
    for name, color, text in FORMULA:
        add_shape(s, x, 2.05, width, 3.75, color, rounded=True, radius=0.08)
        write(s, x + 0.16, 2.25, width - 0.32, 3.4, [
            {"t": name, "size": 22, "color": WHITE if color != AMBER else NAVY, "bold": True},
            {"t": text, "size": 13, "color": WHITE if color != AMBER else NAVY, "before": 8},
            {"t": "Себе: строка\nЛере: строка", "size": 12, "color": AMBER if color != AMBER else CORAL, "bold": True, "before": 14},
        ])
        x += width + gap
    notes(s, "Это формула выбора человека, не формула профессии Климова и не тест.")


def _info_profession_formula(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "five-types.jpg", 0, 0, W, 2.55, darken=0.22)
    write(s, 0.5, 0.55, 12, 0.3, [{"t": "ФОРМУЛА ПРОФЕССИИ  ·  НЕ ТЕСТ", "size": 13, "color": AMBER, "bold": True}])
    write(s, 0.5, 0.95, 12, 0.7, [{"t": "Как описать любую профессию", "size": 30, "color": WHITE, "bold": True}])
    write(s, 0.5, 2.75, 12.3, 0.55, [{"t": SLIDE_SIMPLE["profession_formula"], "size": 15, "color": MUTED}])
    for i, (name, text) in enumerate(PROFESSION_FORMULA):
        x = 0.4 + i * 3.2
        add_shape(s, x, 3.45, 3.05, 3.2, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 3.45, 3.05, 0.14, [TEAL, CORAL, INDIGO, AMBER_D][i])
        write(s, x + 0.2, 3.75, 2.65, 2.7, [{"t": name, "size": 20, "color": NAVY, "bold": True}, {"t": text, "size": 14, "color": INK, "before": 10}])
    notes(s, "Отдельно скажите: сейчас не заполняем тест. Сначала понимаем, из чего состоит профессия.")


def _info_types(prs):
    s = new_blank(prs)
    _info_badge(s)
    eyebrow(s, "Часть формулы профессии · предмет труда")
    title(s, "Пять типов по Климову", size=30)
    write(s, 0.5, 1.3, 12.3, 0.45, [{"t": SLIDE_SIMPLE["types"] + " Это всё ещё не тест — это карта.", "size": 15, "color": MUTED}])
    for i, (code, name, obj, examples) in enumerate(TYPES_KID):
        x = 0.4 + i * 2.55
        add_shape(s, x, 1.95, 2.4, 4.55, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 1.95, 2.4, 0.7, TYPE_COLORS[code], rounded=True, radius=0.1)
        add_shape(s, x, 2.4, 2.4, 0.25, TYPE_COLORS[code])
        write(s, x, 1.95, 2.4, 0.7, [{"t": f"{code}\n{name}", "size": 16, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, x + 0.15, 2.9, 2.1, 3.3, [
            {"t": obj, "size": 14, "color": INK, "align": "center"},
            {"t": examples, "size": 12, "color": MUTED, "before": 12, "align": "center"},
            {"t": "Для Леры\nподумайте:\nпохоже?", "size": 13, "color": TEAL, "bold": True, "before": 16, "align": "center"},
        ])
    notes(s, "Свяжите с кейсом: у Леры чаще звучит тип «человек», возможен стык со «знаками».")


def _test(prs, express: bool):
    s = new_blank(prs)
    picture(s, "self-discovery.jpg", 0, 0, W, H, darken=0.08)
    add_shape(s, 0.7, 1.15, 11.9, 5.15, CREAM)
    add_shape(s, 0.7, 1.15, 11.9, 0.16, CORAL)
    _int_badge(s, "5 МИН" if express else "6 МИН")
    write(s, 1.05, 1.55, 11, 0.3, [{"t": "ТЕСТ ДДО  ·  ОТДЕЛЬНО ОТ ФОРМУЛЫ ПРОФЕССИИ", "size": 14, "color": CORAL, "bold": True}])
    write(s, 1.05, 2.05, 11, 0.9, [{"t": ("5 минут · вопросы 1–10" if express else "6 минут · вопросы 1–15"), "size": 34, "color": NAVY, "bold": True}])
    write(s, 1.05, 3.15, 11, 1.6, [{"t": SLIDE_SIMPLE["test"] + "\n\nВ паре выберите одно: А или Б. Ключ пока не открывайте.", "size": 18, "color": INK}])
    write(s, 1.05, 5.15, 11, 0.6, [{"t": "Потом сравните свой результат с гипотезой по Лере.", "size": 15, "color": TEAL, "bold": True}])
    notes(s, "Явно: тест ≠ формула профессии. Тишина.")


def _key(prs, express: bool):
    s = new_blank(prs)
    eyebrow(s, "Считаем тест")
    title(s, "Ключ ДДО", size=30)
    mx, sm = max_score(express), expected_sum(express)
    write(s, 0.5, 1.2, 12.3, 0.4, [{"t": f"Совпадение = 1 балл. Максимум в столбце — {mx}. Сумма — {sm}. Смотрите на 1–2 верхних столбца.", "size": 15, "color": MUTED}])
    buckets = answers_by_type(express)
    gap = 0.14
    width = (12.33 - 4 * gap) / 5
    x = 0.5
    names = {"П": "природа", "Т": "техника", "Ч": "человек", "З": "знаки", "Х": "образ"}
    for code, items in buckets.items():
        add_shape(s, x, 1.75, width, 4.55, WHITE, rounded=True, radius=0.08)
        add_shape(s, x, 1.75, width, 0.7, TYPE_COLORS[code], rounded=True, radius=0.08)
        add_shape(s, x, 2.2, width, 0.25, TYPE_COLORS[code])
        write(s, x, 1.75, width, 0.7, [{"t": f"{code}  {names[code]}", "size": 14, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, x + 0.08, 2.65, width - 0.16, 3.4, [{"t": item, "size": 17, "color": INK, "bold": True, "align": "center", "after": 4} for item in items])
        x += width + gap
    notes(s, "Ноль — не запрет. Для кейса важнее понять логику типа, чем угадать Леру один в один.")


def _directions_overview(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "social-directions.jpg", 0, 0, W, 2.55, darken=0.28)
    write(s, 0.5, 0.65, 12, 0.3, [{"t": "НАПРАВЛЕНИЯ ПОДГОТОВКИ", "size": 13, "color": AMBER, "bold": True}])
    write(s, 0.5, 1.05, 12, 0.8, [{"t": "Четыре пути, если близки люди и общество", "size": 28, "color": WHITE, "bold": True}])
    write(s, 0.5, 2.75, 12.3, 0.5, [{"t": SLIDE_SIMPLE["directions"], "size": 15, "color": MUTED}])
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, d in enumerate(DIRECTIONS_SLIDE):
        x = 0.4 + i * 3.2
        add_shape(s, x, 3.4, 3.05, 3.25, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 3.4, 3.05, 0.14, accents[i])
        write(s, x + 0.15, 3.65, 2.75, 2.85, [
            {"t": d["title"], "size": 16, "color": NAVY, "bold": True},
            {"t": d["code"], "size": 12, "color": MUTED, "before": 2},
            {"t": d["ask"], "size": 13, "color": accents[i], "before": 10},
            {"t": d["type_hint"], "size": 12, "color": INK, "before": 10},
        ])
    notes(s, "Обязательный блок. Не пропускайте. Это материал для ответа Лере.")


def _directions_detail(prs):
    s = new_blank(prs)
    _info_badge(s)
    eyebrow(s, "Что изучают и куда идут")
    title(s, "Разберём направления подробнее", size=28)
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, d in enumerate(DIRECTIONS_SLIDE):
        col = i % 2
        row = i // 2
        x = 0.45 + col * 6.4
        y = 1.4 + row * 2.7
        add_shape(s, x, y, 6.2, 2.5, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, y, 0.14, 2.5, accents[i])
        write(s, x + 0.35, y + 0.2, 5.6, 2.15, [
            {"t": f"{d['title']}  ·  {d['code']}", "size": 18, "color": NAVY, "bold": True},
            {"t": d["ask"], "size": 14, "color": accents[i], "before": 4},
            {"t": "Чему учат: " + d["do"], "size": 13, "color": INK, "before": 8},
            {"t": "Где работают: " + d["work"], "size": 13, "color": MUTED, "before": 4},
            {"t": d["fit"], "size": 13, "color": INK, "before": 4},
        ])
    notes(s, "По одной минуте на карточку. Для Леры пока не объявляйте «правильный» ответ — его соберут сами.")


def _info_support(prs):
    s = new_blank(prs)
    _info_badge(s)
    picture(s, "youth-support.jpg", 8.4, 0.9, 4.45, 5.7, darken=0.05)
    add_shape(s, 8.4, 0.9, 4.45, 0.12, AMBER_D)
    eyebrow(s, "Для шага в кейсе")
    title(s, "Какая поддержка уже есть", w=7.6, size=28)
    items = [
        ("«Билет в будущее»", "Диагностики и профпробы. Спросите педагога или зайдите на bvbinfo.ru."),
        ("«Профессионалитет»", "Колледжи вместе с работодателями. Важно после 9 класса."),
        ("Целевое обучение", "Договор с заказчиком. Читать только с родителями."),
        ("Движение Первых", "Проекты и команды для школьников."),
    ]
    for i, (head, body) in enumerate(items):
        y = 1.5 + i * 1.2
        add_shape(s, 0.5, y, 7.6, 1.05, WHITE, rounded=True, radius=0.1)
        write(s, 0.75, y + 0.15, 7.15, 0.8, [{"t": head, "size": 16, "color": NAVY, "bold": True}, {"t": body, "size": 13, "color": MUTED, "before": 2}])
    notes(s, "Нужно для 4-го пункта решения кейса: один реальный шаг.")


def _case_solve(prs):
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.45, 0.4, 12.4, 6.7, CREAM)
    _int_badge(s, "6 МИН")
    write(s, 0.75, 0.6, 10, 0.28, [{"t": "РЕШАЕМ КЕЙС", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.75, 0.95, 11.5, 0.5, [{"t": "Карточка ответа для Леры", "size": 30, "color": NAVY, "bold": True}])
    for i, step in enumerate(CASE["solve_steps"]):
        y = 1.65 + i * 1.05
        add_shape(s, 0.75, y, 12.0, 0.9, WHITE, rounded=True, radius=0.1)
        write(s, 1.0, y, 1.0, 0.9, [{"t": f"0{i+1}", "size": 22, "color": CORAL, "bold": True}], anchor="middle")
        write(s, 2.1, y, 10.4, 0.9, [{"t": step, "size": 17, "color": INK, "bold": True}], anchor="middle")
    notes(s, "6 минут на запись в тетради. 2–3 ответа вслух. Критерии: все 4 инструмента, понятный язык, без «просто определись».")


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
    notes(s, "Не читайте письма. Можно связать с шагом, который предложили Лере.")


def _close(prs, dark: set):
    s = new_blank(prs, NAVY)
    dark.add(len(prs.slides))
    picture(s, "hero-path.jpg", 0, 0, W, H, darken=0.4)
    add_shape(s, 0.7, 1.4, 11.9, 4.7, NAVY)
    write(s, 1.0, 1.7, 11, 0.3, [{"t": "НА ВЫХОД", "size": 14, "color": AMBER, "bold": True}])
    write(s, 1.0, 2.15, 11, 1.0, [{"t": "Четыре инструмента и один шаг", "size": 28, "color": WHITE, "bold": True}])
    write(s, 1.0, 3.3, 11, 1.8, [{"t": "1) Хочу — могу — умею — надо\n2) Формула профессии по Климову\n3) Тест ДДО\n4) Направления подготовки\n\nМой шаг на месяц: …", "size": 18, "color": "D8E6EE"}])
    notes(s, "Короткий финал.")


def build_hour():
    prs = _prs("Путь в профессию · 60 минут")
    dark: set[int] = set()
    _hero(prs, 60, dark)
    _flow(prs)
    _case_open(prs)
    _info_self(prs)
    _interactive_facts(prs)
    _info_world(prs)
    _interactive_want(prs)
    _info_profession_formula(prs)
    _info_types(prs)
    _test(prs, express=False)
    _key(prs, express=False)
    _directions_overview(prs)
    _directions_detail(prs)
    _info_support(prs)
    _case_solve(prs)
    _letter(prs)
    _close(prs, dark)
    return finish(prs, dark)


def build_short():
    prs = _prs("Путь в профессию · 20 минут")
    dark: set[int] = set()
    _hero(prs, 20, dark)

    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.5, 0.55, 12.3, 6.4, CREAM)
    write(s, 0.85, 0.8, 11, 0.28, [{"t": "КЕЙС", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.85, 1.2, 11.5, 0.5, [{"t": CASE["title"], "size": 28, "color": NAVY, "bold": True}])
    write(s, 0.85, 1.85, 11.5, 2.4, [{"t": CASE["situation"], "size": 15, "color": INK}])
    write(s, 0.85, 4.4, 11.5, 2.1, [
        {"t": "В конце заполним 4 пункта:", "size": 16, "color": TEAL, "bold": True},
        {"t": "формула выбора · тип Климова · направление · шаг", "size": 16, "color": INK, "before": 8},
    ])
    notes(s, "Кейс сразу. Решение — после инструментов.")

    s = new_blank(prs)
    _info_badge(s)
    eyebrow(s, "Инструменты")
    title(s, "Четыре вещи, которые нужны для ответа")
    for i, (h, b) in enumerate([
        ("Хочу — могу — умею — надо", "Формула выбора человека"),
        ("Формула профессии", "Предмет, цели, средства, условия"),
        ("Тест ДДО", "Отдельно: тип интереса"),
        ("4 направления", "Социология, соцработа, политология, ОРМ"),
    ]):
        y = 1.5 + i * 1.2
        add_shape(s, 0.5, y, 12.3, 1.05, WHITE, rounded=True, radius=0.1)
        write(s, 0.8, y, 11.8, 1.05, [{"t": h, "size": 18, "color": NAVY, "bold": True}, {"t": "  ·  " + b, "size": 15, "color": MUTED}], anchor="middle")

    _interactive_facts(prs)
    _interactive_want(prs, minutes="3 МИН")
    _info_profession_formula(prs)
    _test(prs, express=True)
    _key(prs, express=True)
    _directions_overview(prs)
    _case_solve(prs)
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
