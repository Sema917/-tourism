"""Презентации: визуальный рассказ + несколько интерактивов."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from content import (
    CAN_OPTIONS,
    CASE,
    FORMULA,
    FORMULA_TRAPS,
    MYTHS,
    PATHS,
    PROGRAMS,
    SKILLS,
    SUPPORT,
    TYPES,
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
    NAVY2,
    SOFT_AMBER,
    SOFT_CORAL,
    SOFT_INDIGO,
    SOFT_TEAL,
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


def _hero(prs, minutes: int, dark: set):
    s = new_blank(prs, NAVY)
    dark.add(len(prs.slides))
    picture(s, "hero-path.jpg", 0, 0, W, H, darken=0.28)
    add_shape(s, 0, 0, 7.6, H, None)
    # left gradient panel via solid translucent substitute
    add_shape(s, 0, 0, 7.2, H, NAVY)
    # soften with cream strip accent
    add_shape(s, 7.05, 0, 0.12, H, CORAL)
    write(s, 0.55, 1.35, 6.2, 0.35, [{"t": "ФЕСТИВАЛЬ  «НАУКА 0+»", "size": 15, "color": AMBER, "bold": True}])
    write(s, 0.55, 1.85, 6.3, 1.5, [{"t": "Путь в\nпрофессию", "size": 54, "color": WHITE, "bold": True, "spacing": 0.92}])
    write(
        s,
        0.55,
        3.7,
        6.1,
        0.9,
        [{"t": "Как выбрать дело по интересам,\nумениям и запросу мира вокруг", "size": 20, "color": "D8E6EE"}],
    )
    chip(s, 0.55, 5.0, 2.2, 0.48, f"{minutes} МИНУТ", TEAL)
    chip(s, 2.9, 5.0, 2.4, 0.48, "5–11 КЛАСС", "245066")
    notes(s, "Поздоровайтесь. Скажите: сегодня исследование, а не вердикт. К концу — формула и один шаг, не «единственная профессия».")


def _agenda(prs, minutes: int):
    s = new_blank(prs)
    picture(s, "hero-path.jpg", 8.4, 0, 4.95, 7.5, darken=0.15)
    add_shape(s, 8.25, 0, 0.15, H, CORAL)
    eyebrow(s, "Карта встречи")
    title(s, "О чём поговорим")
    write(s, 0.5, 1.35, 7.5, 0.45, [{"t": "Большая часть часа — рассказ и картинка. Интерактивов немного, и они отмечены отдельно.", "size": 15, "color": MUTED}])
    items = [
        ("01", "Самопознание", "Интересы и типы профессий по Климову", CORAL),
        ("02", "Рынок труда", "Что нужно людям и стране сегодня", INDIGO),
        ("03", "Учёба и навыки", "Маршруты после школы и умения XXI века", TEAL),
        ("04", "План и поддержка", "Шаг на месяц и государственные двери", AMBER_D),
    ]
    for i, (n, head, body, color) in enumerate(items):
        y = 1.95 + i * 1.15
        add_shape(s, 0.5, y, 7.5, 1.0, WHITE, rounded=True, radius=0.1)
        add_shape(s, 0.5, y, 0.16, 1.0, color)
        write(s, 0.9, y, 1.0, 1.0, [{"t": n, "size": 22, "color": color, "bold": True}], anchor="middle")
        write(s, 2.0, y + 0.12, 5.7, 0.8, [{"t": head, "size": 20, "color": INK, "bold": True}, {"t": body, "size": 14, "color": MUTED, "before": 2}])
    notes(s, f"Один проход. На {minutes} минут не обещайте разобрать всё до деталей тетради — тетрадь унесут с собой.")


def _section(prs, dark: set, number: str, kicker: str, heading: str, line: str, image: str, accent: str):
    s = new_blank(prs, NAVY)
    dark.add(len(prs.slides))
    picture(s, image, 0, 0, W, H, darken=0.35)
    add_shape(s, 0, 0, W, H, None)
    add_shape(s, 0.55, 1.7, 8.5, 4.0, NAVY)
    write(s, 0.85, 1.95, 3, 0.35, [{"t": f"ТЕМА  {number}", "size": 14, "color": accent, "bold": True}])
    write(s, 0.85, 2.4, 7.8, 1.6, [{"t": heading, "size": 36, "color": WHITE, "bold": True, "spacing": 0.95}])
    write(s, 0.85, 4.2, 7.6, 1.0, [{"t": line, "size": 18, "color": "D8E6EE"}])
    notes(s, "10 секунд. Не задерживайтесь — сразу к содержанию темы.")


def _content_self(prs):
    s = new_blank(prs)
    picture(s, "self-discovery.jpg", 7.7, 0.9, 5.1, 5.7, darken=0.05)
    add_shape(s, 7.7, 0.9, 5.1, 0.12, TEAL)
    eyebrow(s, "Самопознание")
    title(s, "Сначала занятие, не должность")
    bullets = [
        "«Юрист» или «блогер» — это вывеска. Важнее, что вы делаете: спорите по правилам, объясняете, считаете, чините, ухаживаете.",
        "Интерес проверяют делом и вопросами, а не только мечтой о зарплате.",
        "Сегодняшний опросник Климова спрашивает: с чем вам интереснее иметь дело — с природой, техникой, людьми, знаками или образом.",
    ]
    for i, text in enumerate(bullets):
        y = 1.55 + i * 1.55
        add_shape(s, 0.5, y, 6.9, 1.4, WHITE, rounded=True, radius=0.1)
        write(s, 0.75, y + 0.2, 6.4, 1.05, [{"t": text, "size": 16, "color": INK, "spacing": 1.02}])
    notes(s, "Расскажите спокойно. Не превращайте в опрос зала. Переход к формуле — следующий слайд.")


def _formula(prs):
    s = new_blank(prs)
    picture(s, "formula-visual.jpg", 0, 0, W, H, darken=0.12)
    add_shape(s, 0.35, 0.3, 12.6, 6.85, CREAM)
    write(s, 0.65, 0.5, 8, 0.3, [{"t": "ИНТЕРАКТИВ  ·  3 МИНУТЫ", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.65, 0.85, 12, 0.55, [{"t": "Формула профессии", "size": 34, "color": NAVY, "bold": True}])
    write(s, 0.65, 1.45, 12.1, 0.45, [{"t": FORMULA_TRAPS, "size": 15, "color": MUTED}])
    gap = 0.18
    width = (12.1 - 3 * gap) / 4
    x = 0.65
    for name, color, text in FORMULA:
        add_shape(s, x, 2.15, width, 3.7, color, rounded=True, radius=0.08)
        write(s, x + 0.18, 2.35, width - 0.36, 3.3, [
            {"t": name, "size": 24, "color": WHITE if color != AMBER else NAVY, "bold": True},
            {"t": text, "size": 14, "color": WHITE if color != AMBER else NAVY, "before": 10, "spacing": 1.0},
            {"t": "Запишите строку\nв тетради", "size": 13, "color": AMBER if color != AMBER else CORAL, "bold": True, "before": 18},
        ])
        x += width + gap
    notes(s, "Единственный интерактив блока. 3 минуты: кратко поясните четыре слова и дайте записать по одной строке. Баллы и должности не собирайте.")


def _types(prs):
    s = new_blank(prs)
    picture(s, "five-types.jpg", 0, 0, W, 3.15, darken=0.2)
    add_shape(s, 0, 2.55, W, 0.6, CREAM)
    eyebrow(s, "Метод Е. А. Климова", y=3.2)
    title(s, "Пять типов по предмету труда", y=3.45, size=28)
    for i, t in enumerate(TYPES):
        x = 0.4 + i * 2.55
        add_shape(s, x, 4.25, 2.4, 2.5, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 4.25, 2.4, 0.55, TYPE_COLORS[t["code"]], rounded=True, radius=0.1)
        add_shape(s, x, 4.55, 2.4, 0.25, TYPE_COLORS[t["code"]])
        write(s, x, 4.25, 2.4, 0.55, [{"t": t["code"], "size": 18, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, x + 0.12, 4.95, 2.16, 1.6, [
            {"t": t["name"].replace("Человек — ", ""), "size": 14, "color": INK, "bold": True, "align": "center"},
            {"t": t["slide"], "size": 11, "color": MUTED, "before": 6, "align": "center"},
        ])
    notes(s, "Коротко по пяти. Факт: лаборатория в больнице — природа, беседа с больным — человек. Границы типов размыты — это нормально.")


def _ddo(prs, express: bool, dark: set):
    s = new_blank(prs)
    picture(s, "self-discovery.jpg", 0, 0, W, H, darken=0.08)
    add_shape(s, 0.7, 1.15, 11.9, 5.15, CREAM)
    add_shape(s, 0.7, 1.15, 11.9, 0.16, CORAL)
    write(s, 1.05, 1.55, 11, 0.35, [{"t": "ИНТЕРАКТИВ  ·  ТЕСТ ДДО", "size": 14, "color": CORAL, "bold": True}])
    if express:
        write(s, 1.05, 2.1, 11, 1.0, [{"t": "5 минут. Вопросы 1–10", "size": 40, "color": NAVY, "bold": True}])
        write(s, 1.05, 3.3, 11, 1.4, [{"t": "В каждой паре выберите одно занятие — А или Б.\nПервый отклик. Ключ пока не открывайте.", "size": 22, "color": INK}])
    else:
        write(s, 1.05, 2.1, 11, 1.0, [{"t": "6 минут. Вопросы 1–15", "size": 40, "color": NAVY, "bold": True}])
        write(s, 1.05, 3.3, 11, 1.4, [{"t": "Сокращённый опросник Е. А. Климова.\nВ паре один выбор. Это гипотеза, не приговор.", "size": 22, "color": INK}])
    write(s, 1.05, 5.15, 11, 0.7, [{"t": "Непонятное слово — тихая рука. Баллы вслух не собираем.", "size": 16, "color": TEAL, "bold": True}])
    notes(s, "Молчите во время теста. Подсказывайте значения слов, не выборы. За минуту: поставьте выбор даже при сомнении.")


def _key(prs, express: bool):
    s = new_blank(prs)
    eyebrow(s, "Считаем вместе")
    title(s, "Ключ" + (" экспресс" if express else " сокращённого варианта"), size=30)
    mx = max_score(express)
    sm = expected_sum(express)
    write(s, 0.5, 1.25, 12.3, 0.35, [{"t": f"Один балл за совпадение. Максимум в столбце — {mx}. Сумма — {sm}.", "size": 15, "color": MUTED}])
    buckets = answers_by_type(express)
    gap = 0.14
    width = (12.33 - 4 * gap) / 5
    x = 0.5
    for code, items in buckets.items():
        add_shape(s, x, 1.75, width, 4.55, WHITE, rounded=True, radius=0.08)
        add_shape(s, x, 1.75, width, 0.75, TYPE_COLORS[code], rounded=True, radius=0.08)
        add_shape(s, x, 2.25, width, 0.25, TYPE_COLORS[code])
        name = {"П": "природа", "Т": "техника", "Ч": "человек", "З": "знаки", "Х": "образ"}[code]
        write(s, x, 1.75, width, 0.75, [{"t": f"{code}  {name}", "size": 14, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, x + 0.08, 2.7, width - 0.16, 3.3, [{"t": item, "size": 18, "color": INK, "bold": True, "align": "center", "after": 5} for item in items])
        x += width + gap
    notes(s, "Дайте посчитать. Смотрите на 1–2 верхних столбца. Ноль — не запрет. Разница в 1 балл почти ничего не значит.")


def _market(prs):
    s = new_blank(prs)
    picture(s, "labor-market.jpg", 0, 0.85, 6.2, 5.9, darken=0.08)
    add_shape(s, 0, 0.85, 6.2, 0.12, INDIGO)
    eyebrow(s, "Рынок труда")
    title(s, "Спрос шире витрины", x=6.55, w=6.3)
    facts = [
        ("Не только громкие профессии", "Медицина, образование, инженерия, рабочие специальности, социальная сфера, данные, экология."),
        ("Ценен стык", "Люди и данные. Техника и умение объяснить. Забота и знание правил."),
        ("Умения живут дольше названий", "Учиться, проверять факты и договариваться нужны почти везде."),
    ]
    for i, (head, body) in enumerate(facts):
        y = 1.55 + i * 1.6
        add_shape(s, 6.55, y, 6.3, 1.45, WHITE, rounded=True, radius=0.1)
        write(s, 6.8, y + 0.2, 5.85, 1.1, [{"t": head, "size": 18, "color": INDIGO, "bold": True}, {"t": body, "size": 14, "color": INK, "before": 4}])
    notes(s, "Рассказ ведущего. Зарплаты не называйте. Нейросети — одной фразой: меняют задачи, не отменяют ответственность.")


def _myths(prs):
    s = new_blank(prs)
    eyebrow(s, "Частые заблуждения")
    title(s, "Три мифа и один факт")
    for i, (statement, kind, why) in enumerate(MYTHS):
        y = 1.45 + i * 1.3
        fill = TEAL if kind == "Факт" else CORAL
        add_shape(s, 0.5, y, 12.33, 1.15, WHITE, rounded=True, radius=0.1)
        add_shape(s, 0.5, y, 1.7, 1.15, fill)
        write(s, 0.5, y, 1.7, 1.15, [{"t": kind.upper(), "size": 15, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, 2.45, y + 0.15, 10.1, 0.9, [{"t": statement, "size": 17, "color": INK, "bold": True}, {"t": why, "size": 14, "color": MUTED, "before": 3}])
    notes(s, "Не голосование. Прочитайте как факты. Можно спросить «знакомо?» одним словом и идти дальше.")


def _paths(prs):
    s = new_blank(prs)
    picture(s, "edu-paths.jpg", 0, 0, W, 2.9, darken=0.25)
    add_shape(s, 0, 2.35, W, 0.55, CREAM)
    eyebrow(s, "Образовательные траектории", y=3.05)
    title(s, "Где вы сейчас — такая и задача", y=3.3, size=28)
    colors = [TEAL, AMBER_D, INDIGO]
    for i, (head, body) in enumerate(PATHS):
        x = 0.45 + i * 4.25
        add_shape(s, x, 4.15, 4.05, 2.55, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 4.15, 4.05, 0.14, colors[i])
        write(s, x + 0.22, 4.45, 3.6, 2.1, [{"t": head, "size": 20, "color": INK, "bold": True}, {"t": body, "size": 13, "color": MUTED, "before": 6}])
    notes(s, "Младшим: вуз называть не нужно. 9 классу: колледж — не утешительный приз. Старшим: решение с семьёй и по правилам года.")


def _skills(prs):
    s = new_blank(prs)
    eyebrow(s, "Навыки XXI века")
    title(s, "Что пригождается почти в любой работе", size=28)
    for i, (head, body) in enumerate(SKILLS):
        col = i % 4
        row = i // 4
        if i < 4:
            x = 0.45 + col * 3.2
            y = 1.5
            w = 3.05
        else:
            x = 0.45 + (i - 4) * 4.25
            y = 4.15
            w = 4.05
        add_shape(s, x, y, w, 2.35, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, y, 0.14, 2.35, TEAL if i % 2 == 0 else CORAL)
        write(s, x + 0.3, y + 0.25, w - 0.5, 1.9, [{"t": head, "size": 17, "color": NAVY, "bold": True}, {"t": body, "size": 13, "color": MUTED, "before": 8}])
    notes(s, "Не читайте все семь медленно. Назовите и задержитесь на этике работы с людьми.")


def _programs(prs):
    s = new_blank(prs)
    picture(s, "social-directions.jpg", 0, 0, W, 2.7, darken=0.3)
    write(s, 0.5, 0.85, 12, 0.35, [{"t": "ВЫСШЕЕ ОБРАЗОВАНИЕ  ·  ПРИМЕР, НЕ КАТАЛОГ", "size": 13, "color": AMBER, "bold": True}])
    write(s, 0.5, 1.25, 12, 0.7, [{"t": "Четыре двери, если близки люди и общество", "size": 28, "color": WHITE, "bold": True}])
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, p in enumerate(PROGRAMS):
        col = i % 2
        row = i // 2
        x = 0.45 + col * 6.45
        y = 3.0 + row * 2.05
        add_shape(s, x, y, 6.25, 1.9, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, y, 0.14, 1.9, accents[i])
        write(s, x + 0.35, y + 0.2, 5.7, 1.5, [
            {"t": f"{p['title']}   {p['code']}", "size": 18, "color": INK, "bold": True},
            {"t": p["question"], "size": 14, "color": accents[i], "italic": True, "before": 4},
            {"t": p["slide_jobs"], "size": 13, "color": MUTED, "before": 4},
        ])
    notes(s, "Рамка: это не список «выберите отсюда». Если тип техника/природа — ваши двери другие и не хуже. ЕГЭ — в тетради, на экране не диктуйте.")


def _support(prs):
    s = new_blank(prs)
    picture(s, "youth-support.jpg", 8.55, 0.9, 4.3, 5.7, darken=0.05)
    add_shape(s, 8.55, 0.9, 4.3, 0.12, AMBER_D)
    eyebrow(s, "Государственная поддержка")
    title(s, "Двери, которые уже открыты", size=28, w=7.8)
    items = SUPPORT[:6]
    for i, item in enumerate(items):
        y = 1.5 + i * 0.85
        add_shape(s, 0.5, y, 7.8, 0.75, WHITE, rounded=True, radius=0.1)
        write(s, 0.75, y, 7.35, 0.75, [{"t": item["title"], "size": 15, "color": NAVY, "bold": True}], anchor="middle")
    notes(s, "По одной фразе. Целевое — только с родителями, это обязательство. Адреса и осторожности — в тетради. «Движение Первых» и работа с 14 лет тоже там.")


def _case(prs):
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.12)
    add_shape(s, 0.45, 0.4, 12.4, 6.7, CREAM)
    write(s, 0.75, 0.6, 11, 0.3, [{"t": "ИНТЕРАКТИВ  ·  КЕЙС  ·  6 МИНУТ", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.75, 0.95, 11.5, 0.55, [{"t": CASE["title"], "size": 32, "color": NAVY, "bold": True}])
    write(s, 0.75, 1.6, 11.8, 0.95, [{"t": CASE["situation"], "size": 15, "color": INK}])
    write(s, 0.75, 2.65, 11.8, 0.4, [{"t": CASE["task"], "size": 17, "color": TEAL, "bold": True}])
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, (role, code, question) in enumerate(CASE["roles"]):
        x = 0.75 + (i % 4) * 3.0
        add_shape(s, x, 3.3, 2.85, 3.4, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 3.3, 2.85, 0.14, accents[i])
        write(s, x + 0.15, 3.55, 2.55, 3.0, [
            {"t": role, "size": 14, "color": NAVY, "bold": True},
            {"t": code, "size": 12, "color": accents[i], "before": 2},
            {"t": question, "size": 13, "color": INK, "before": 10},
        ])
    notes(s, "Один командный интерактив. Если ряды не двигаются — вдвоём или каждый сам. Вслух 2–3 ответа по полминуты.")


def _letter(prs, dark: set):
    s = new_blank(prs)
    picture(s, "future-letter.jpg", 0, 0, W, H, darken=0.08)
    add_shape(s, 0.7, 1.05, 11.9, 5.35, CREAM)
    add_shape(s, 0.7, 1.05, 11.9, 0.16, AMBER_D)
    write(s, 1.0, 1.4, 11, 0.3, [{"t": "ИНТЕРАКТИВ  ·  ПИСЬМО", "size": 13, "color": AMBER_D, "bold": True}])
    write(s, 1.0, 1.8, 11, 0.65, [{"t": "Письмо себе из будущего", "size": 34, "color": NAVY, "bold": True}])
    lines = [
        "Мне ____ лет. Я занимаюсь…",
        "Мне по-настоящему нравится…",
        "Из школьного пригодилось…",
        "Одна фраза себе сегодняшнему…",
    ]
    for i, line in enumerate(lines):
        y = 2.7 + i * 0.75
        add_shape(s, 1.0, y, 11.3, 0.65, WHITE, rounded=True, radius=0.12)
        write(s, 1.25, y, 10.8, 0.65, [{"t": line, "size": 18, "color": INK}], anchor="middle")
    notes(s, "3–4 минуты. Письмо не читаете и не собираете. Для 5–6 класса достаточно одной фразы или рисунка.")


def _close(prs, dark: set):
    s = new_blank(prs, NAVY)
    dark.add(len(prs.slides))
    picture(s, "hero-path.jpg", 0, 0, W, H, darken=0.4)
    add_shape(s, 0.7, 1.4, 11.9, 4.7, NAVY2)
    write(s, 1.0, 1.7, 11, 0.35, [{"t": "НА ВЫХОД", "size": 14, "color": AMBER, "bold": True}])
    write(s, 1.0, 2.15, 11, 0.7, [{"t": "Гипотеза. Один шаг. Тетрадь с собой.", "size": 30, "color": WHITE, "bold": True}])
    write(s, 1.0, 3.1, 11, 1.3, [{"t": "Мне интересно…  Я могу…  Я уже умею…  Другим нужно…\n\nМой шаг на этот месяц: …", "size": 20, "color": "D8E6EE"}])
    write(s, 1.0, 4.8, 11, 0.8, [{"t": "Вопрос домой: какую пробу я сделаю в этом месяце?", "size": 18, "color": AMBER, "bold": True}])
    notes(s, "Не открывайте новую тему. Поблагодарите.")


def _can_brief(prs):
    """Short non-workshop content slide reminding могу/умею difference before programs or after skills."""
    s = new_blank(prs)
    eyebrow(s, "Напоминание к формуле")
    title(s, "«Могу» и «умею» — не одно и то же")
    add_shape(s, 0.5, 1.6, 6.0, 4.7, SOFT_TEAL, rounded=True, radius=0.1)
    add_shape(s, 6.8, 1.6, 6.0, 4.7, SOFT_AMBER, rounded=True, radius=0.1)
    write(s, 0.8, 1.9, 5.4, 4.1, [
        {"t": "Могу", "size": 28, "color": TEAL, "bold": True},
        {"t": "Способность, которую вы в себе замечаете: объяснять, замечать детали, собирать людей, мастерить…", "size": 18, "color": INK, "before": 12},
        {"t": "Отметьте два варианта в тетради, когда будет минута.", "size": 15, "color": MUTED, "before": 16},
    ])
    write(s, 7.1, 1.9, 5.4, 4.1, [
        {"t": "Умею", "size": 28, "color": AMBER_D, "bold": True},
        {"t": "То, что уже получалось и это можно показать. Вас просили — выходило — довели до конца.", "size": 18, "color": INK, "before": 12},
        {"t": "«Умею лечить» в 7 классе обычно неправда. «Умею объяснить тему другу» — может быть правдой.", "size": 15, "color": MUTED, "before": 16},
    ])
    notes(s, "Не устраивайте опрос. Это пояснение к формуле. Запись — по желанию в тетради.")


def build_hour():
    prs = _prs("Путь в профессию · 60 минут")
    dark: set[int] = set()
    _hero(prs, 60, dark)
    _agenda(prs, 60)
    _section(prs, dark, "01", "Самопознание", "Кто вы\nпо интересу", "Интерес — к предмету труда, а не к красивому названию.", "self-discovery.jpg", CORAL)
    _content_self(prs)
    _formula(prs)
    _types(prs)
    _ddo(prs, express=False, dark=dark)
    _key(prs, express=False)
    _section(prs, dark, "02", "Рынок труда", "Что нужно\nмиру вокруг", "Спрос — информация для выбора, не рейтинг людей.", "labor-market.jpg", INDIGO)
    _market(prs)
    _myths(prs)
    _section(prs, dark, "03", "Траектории", "Путь длиннее\nэкзамена", "Сначала проба. Потом развилка. Потом дверь.", "edu-paths.jpg", TEAL)
    _paths(prs)
    _skills(prs)
    _can_brief(prs)
    _programs(prs)
    _section(prs, dark, "04", "План и поддержка", "От гипотезы\nк шагу", "Государственные возможности имеют имена и адреса.", "youth-support.jpg", AMBER)
    _support(prs)
    _case(prs)
    _letter(prs, dark)
    _close(prs, dark)
    return finish(prs, dark)


def build_short():
    prs = _prs("Путь в профессию · 20 минут")
    dark: set[int] = set()
    _hero(prs, 20, dark)
    # compact agenda + formula
    s = new_blank(prs)
    picture(s, "formula-visual.jpg", 7.9, 0.9, 4.9, 5.7, darken=0.1)
    eyebrow(s, "За 20 минут")
    title(s, "Формула и один шаг", size=30, w=7.2)
    write(s, 0.5, 1.4, 7.1, 0.8, [{"t": "Не выбираем судьбу. Успеем гипотезу интереса, формулу и шаг на месяц. Остальное — в тетради.", "size": 16, "color": MUTED}])
    for i, (name, color, text) in enumerate(FORMULA):
        y = 2.35 + i * 1.05
        add_shape(s, 0.5, y, 7.1, 0.92, color, rounded=True, radius=0.1)
        write(s, 0.75, y, 6.6, 0.92, [{"t": name, "size": 18, "color": WHITE if color != AMBER else NAVY, "bold": True}, {"t": text, "size": 12, "color": WHITE if color != AMBER else NAVY, "before": 2}], anchor="middle")
    notes(s, "1.5 минуты. Поясните формулу. Полную запись можно доделать дома.")

    _ddo(prs, express=True, dark=dark)
    _key(prs, express=True)
    _types(prs)

    s = new_blank(prs)
    picture(s, "labor-market.jpg", 0, 0, 6.4, H, darken=0.2)
    add_shape(s, 6.15, 0, 7.2, H, CREAM)
    write(s, 6.55, 0.7, 6.3, 0.35, [{"t": "РЫНОК И МАРШРУТ", "size": 13, "color": INDIGO, "bold": True}])
    write(s, 6.55, 1.15, 6.3, 1.0, [{"t": "Спрос шире модных названий", "size": 28, "color": INK, "bold": True}])
    for i, line in enumerate([
        "Нужны забота, инженеры, рабочие профессии, данные, земля.",
        "5–8 класс: пробы. 9 класс: школа или колледж. 10–11: поступление.",
        "Запишите одну строку «надо» в тетради.",
    ]):
        y = 2.5 + i * 1.3
        add_shape(s, 6.55, y, 6.2, 1.1, WHITE, rounded=True, radius=0.1)
        write(s, 6.8, y, 5.75, 1.1, [{"t": line, "size": 15, "color": INK}], anchor="middle")
    notes(s, "Короткий рассказ, без голосований.")

    _programs(prs)

    s = new_blank(prs)
    picture(s, "youth-support.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.55, 1.15, 12.2, 5.2, CREAM)
    write(s, 0.9, 1.45, 11, 0.35, [{"t": "ПОДДЕРЖКА", "size": 13, "color": AMBER_D, "bold": True}])
    write(s, 0.9, 1.9, 11, 0.6, [{"t": "Четыре имени, детали — в тетради", "size": 28, "color": NAVY, "bold": True}])
    names = ["«Билет в будущее»", "«Профессионалитет»", "Целевое обучение", "Госуслуги и «Движение Первых»"]
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, name in enumerate(names):
        x = 0.9 + (i % 2) * 5.8
        y = 2.8 + (i // 2) * 1.5
        add_shape(s, x, y, 5.5, 1.25, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, y, 0.16, 1.25, accents[i])
        write(s, x + 0.35, y, 5.0, 1.25, [{"t": name, "size": 18, "color": NAVY, "bold": True}], anchor="middle")
    notes(s, "Целевое — с родителями. Не читайте каталог.")

    # micro case as one interactive
    s = new_blank(prs)
    picture(s, "case-center.jpg", 0, 0, W, H, darken=0.1)
    add_shape(s, 0.5, 0.55, 12.3, 6.4, CREAM)
    write(s, 0.85, 0.8, 11, 0.3, [{"t": "ИНТЕРАКТИВ  ·  2 МИНУТЫ", "size": 13, "color": CORAL, "bold": True}])
    write(s, 0.85, 1.2, 11.5, 1.1, [{"t": "Одноклассник говорит:\n«Мне всё равно, кем быть»", "size": 30, "color": NAVY, "bold": True}])
    write(s, 0.85, 2.5, 11.5, 0.5, [{"t": "Выберите одну роль и напишите одно действие — не лекцию.", "size": 18, "color": INK}])
    accents = [TEAL, CORAL, INDIGO, AMBER_D]
    for i, (role, code, _) in enumerate(CASE["roles"]):
        x = 0.85 + i * 3.0
        add_shape(s, x, 3.3, 2.85, 3.2, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 3.3, 2.85, 0.14, accents[i])
        write(s, x + 0.15, 3.3, 2.55, 3.2, [{"t": role, "size": 15, "color": NAVY, "bold": True, "align": "center"}, {"t": code, "size": 12, "color": accents[i], "align": "center", "before": 8}], anchor="middle")
    notes(s, "Пишут все. Вслух максимум 3 человека.")

    _letter(prs, dark)
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
