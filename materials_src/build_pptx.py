"""Презентация 20 минут: компактный ход, стиль квиза, направления + опросник."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from content import DIRECTIONS_SLIDE, FORMULA, FORMULA_TRAPS, MINUTES, SLIDE_SIMPLE, TYPES_KID, VERSION
from ddo import answers_by_type, expected_sum, max_score
from pptx_kit import (
    AMBER_D,
    BURGUNDY,
    BURGUNDY_DEEP,
    CREAM,
    CREAM_SOFT,
    GOLD,
    GOLD_SOFT,
    INDIGO,
    INK,
    MUTED,
    SERIF,
    TEAL,
    TYPE_COLORS,
    WHITE,
    W,
    H,
    add_shape,
    chip,
    eyebrow,
    finish,
    frame,
    hairline,
    new_blank,
    notes,
    option_row,
    picture,
    title,
    write,
)


def _prs():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "Путь в профессию · 20 минут"
    prs.core_properties.author = "Комплект «Путь в профессию», фестиваль «Наука 0+»"
    prs.core_properties.subject = f"Комплект {VERSION}"
    prs.core_properties.language = "ru-RU"
    return prs


def _hero(prs, dark: set):
    s = new_blank(prs, BURGUNDY)
    dark.add(len(prs.slides))
    frame(s, GOLD, pt=1.5)
    picture(s, "hero-path.jpg", 0.55, 0.7, 5.1, 5.7, darken=0.08)
    write(s, 6.05, 1.35, 6.5, 0.35, [{"t": "ФЕСТИВАЛЬ  «НАУКА 0+»", "size": 14, "color": GOLD, "bold": True}])
    write(
        s,
        6.05,
        1.85,
        6.6,
        2.2,
        [{"t": "Путь в\nпрофессию", "size": 52, "color": WHITE, "bold": True, "spacing": 0.92, "font": SERIF}],
    )
    write(
        s,
        6.05,
        4.3,
        6.4,
        1.1,
        [{"t": "Формула выбора · типы Климова\nнаправления подготовки · опросник", "size": 18, "color": GOLD_SOFT}],
    )
    chip(s, 6.05, 5.7, 2.35, 0.48, f"{MINUTES} МИНУТ", GOLD, BURGUNDY_DEEP, size=13)
    chip(s, 8.55, 5.7, 2.5, 0.48, "5–11 КЛАСС", BURGUNDY_DEEP, GOLD_SOFT, size=13)
    notes(s, "Короткий старт: сегодня не назначаем профессию, берём инструменты.")


def _map(prs):
    s = new_blank(prs, CREAM)
    frame(s)
    hairline(s, 0.78)
    hairline(s, 6.85)
    eyebrow(s, "Карта встречи")
    title(s, "Что успеем за 20 минут")
    write(s, 0.55, 1.55, 12.2, 0.55, [{"t": SLIDE_SIMPLE["agenda"], "size": 16, "color": MUTED}])
    write(s, 0.55, 2.15, 12.2, 0.7, [{"t": SLIDE_SIMPLE["meaning"], "size": 17, "color": INK, "font": SERIF}])
    steps = [
        ("01", "Формула", "Хочу · могу · умею · надо"),
        ("02", "Типы", "Пять предметов труда"),
        ("03", "Направления", "Четыре пути в вузе"),
        ("04", "Опросник", "15 пар по Климову"),
    ]
    accents = [BURGUNDY, TEAL, INDIGO, AMBER_D]
    for i, ((n, head, body), accent) in enumerate(zip(steps, accents)):
        x = 0.55 + i * 3.15
        add_shape(s, x, 3.15, 3.0, 3.2, WHITE, line=GOLD_SOFT, rounded=True, radius=0.12, line_pt=1.5)
        write(s, x + 0.25, 3.4, 2.5, 0.55, [{"t": n, "size": 28, "color": accent, "bold": True, "font": SERIF}])
        write(s, x + 0.25, 4.15, 2.5, 0.55, [{"t": head, "size": 20, "color": INK, "bold": True, "font": SERIF}])
        write(s, x + 0.25, 4.85, 2.5, 1.0, [{"t": body, "size": 15, "color": MUTED}])
    notes(s, "1 минута. Дальше сразу формула на раздатке.")


def _formula(prs):
    s = new_blank(prs, CREAM)
    frame(s)
    hairline(s, 0.78)
    hairline(s, 6.85)
    eyebrow(s, "Интерактив · 4 минуты")
    title(s, "Хочу · Могу · Умею · Надо")
    write(s, 0.55, 1.55, 12.2, 0.45, [{"t": FORMULA_TRAPS, "size": 15, "color": MUTED}])
    gap = 0.2
    width = (12.2 - 3 * gap) / 4
    x = 0.55
    for name, color, text in FORMULA:
        add_shape(s, x, 2.2, width, 4.2, color, rounded=True, radius=0.1)
        write(
            s,
            x + 0.2,
            2.45,
            width - 0.4,
            3.7,
            [
                {"t": name, "size": 24, "color": WHITE, "bold": True, "font": SERIF},
                {"t": text, "size": 14, "color": GOLD_SOFT, "before": 12},
                {"t": "Запишите\nв раздатку", "size": 14, "color": WHITE, "bold": True, "before": 22},
            ],
        )
        x += width + gap
    notes(s, "4 минуты тишины. Заполняют графы в раздатке.")


def _types(prs):
    s = new_blank(prs, CREAM)
    frame(s)
    hairline(s, 0.78)
    hairline(s, 6.85)
    eyebrow(s, "Смысл · предмет труда")
    title(s, "Пять типов по Климову")
    write(s, 0.55, 1.55, 12.2, 0.45, [{"t": SLIDE_SIMPLE["types"], "size": 16, "color": MUTED}])
    for i, (code, name, obj, examples) in enumerate(TYPES_KID):
        x = 0.45 + i * 2.55
        add_shape(s, x, 2.2, 2.4, 4.2, WHITE, line=GOLD_SOFT, rounded=True, radius=0.1, line_pt=1.4)
        add_shape(s, x, 2.2, 2.4, 0.85, TYPE_COLORS[code], rounded=True, radius=0.1)
        add_shape(s, x, 2.75, 2.4, 0.3, TYPE_COLORS[code])
        write(
            s,
            x,
            2.2,
            2.4,
            0.85,
            [{"t": f"{code}\n{name}", "size": 16, "color": WHITE, "bold": True, "align": "center", "font": SERIF}],
            anchor="middle",
        )
        write(
            s,
            x + 0.15,
            3.25,
            2.1,
            2.8,
            [
                {"t": obj, "size": 14, "color": INK, "align": "center"},
                {"t": examples, "size": 12, "color": MUTED, "before": 14, "align": "center"},
            ],
        )
    notes(s, "2 минуты. Свяжите с формулой: предмет труда ≠ название должности.")


def _directions(prs):
    s = new_blank(prs, CREAM)
    frame(s)
    hairline(s, 0.78)
    hairline(s, 6.85)
    picture(s, "social-directions.jpg", 0.55, 1.7, 4.5, 4.7, darken=0.05)
    eyebrow(s, "Направления подготовки")
    title(s, "Четыре пути, если близки люди", w=7.5, x=5.35)
    write(s, 5.35, 1.55, 7.3, 0.55, [{"t": SLIDE_SIMPLE["directions"], "size": 14, "color": MUTED}])
    accents = [TEAL, BURGUNDY, INDIGO, AMBER_D]
    for i, d in enumerate(DIRECTIONS_SLIDE):
        y = 2.2 + i * 1.05
        add_shape(s, 5.35, y, 7.35, 0.95, WHITE, line=GOLD_SOFT, rounded=True, radius=0.12, line_pt=1.3)
        add_shape(s, 5.35, y, 0.14, 0.95, accents[i])
        write(
            s,
            5.7,
            y + 0.1,
            6.8,
            0.75,
            [
                {"t": f"{d['title']}  ·  {d['code']}", "size": 15, "color": INK, "bold": True, "font": SERIF},
                {"t": d["ask"], "size": 13, "color": accents[i], "italic": True, "before": 2, "font": SERIF},
            ],
        )
    notes(s, "3 минуты. Назовите вопрос каждого направления. Не читайте длинно.")


def _test(prs):
    s = new_blank(prs, CREAM)
    frame(s)
    hairline(s, 0.78)
    hairline(s, 6.85)
    picture(s, "self-discovery.jpg", 0.55, 1.55, 5.0, 4.9, darken=0.04)
    eyebrow(s, "Опросник · 7 минут")
    title(s, "15 пар по Климову", x=5.9, w=6.8)
    write(
        s,
        5.9,
        1.6,
        6.8,
        1.2,
        [{"t": SLIDE_SIMPLE["test"], "size": 16, "color": INK, "font": SERIF}],
    )
    write(
        s,
        5.9,
        2.95,
        6.8,
        0.4,
        [{"t": "В каждой паре выберите одно: А или Б", "size": 18, "color": BURGUNDY, "italic": True, "font": SERIF}],
    )
    option_row(s, 5.9, 3.55, 6.8, 0.85, "А", "Первый вариант из пары", BURGUNDY)
    option_row(s, 5.9, 4.55, 6.8, 0.85, "Б", "Второй вариант из пары", TEAL)
    write(
        s,
        5.9,
        5.65,
        6.8,
        0.55,
        [{"t": "Работайте в раздатке. Ключ пока не открывайте.", "size": 14, "color": MUTED}],
    )
    notes(s, "7 минут. Тишина. Вопросы 1–15 на одной странице раздатки.")


def _key(prs):
    s = new_blank(prs, CREAM)
    frame(s)
    hairline(s, 0.78)
    hairline(s, 6.85)
    eyebrow(s, "Считаем")
    title(s, "Ключ опросника")
    mx, sm = max_score(False), expected_sum(False)
    write(
        s,
        0.55,
        1.55,
        12.2,
        0.4,
        [{"t": f"Совпадение = 1 балл. Максимум в столбце — {mx}. Сумма — {sm}. Смотрите на 1–2 верхних.", "size": 15, "color": MUTED}],
    )
    buckets = answers_by_type(False)
    gap = 0.16
    width = (12.2 - 4 * gap) / 5
    x = 0.55
    names = {"П": "природа", "Т": "техника", "Ч": "человек", "З": "знаки", "Х": "образ"}
    for code, items in buckets.items():
        add_shape(s, x, 2.15, width, 4.35, WHITE, line=GOLD_SOFT, rounded=True, radius=0.1, line_pt=1.3)
        add_shape(s, x, 2.15, width, 0.75, TYPE_COLORS[code], rounded=True, radius=0.1)
        add_shape(s, x, 2.65, width, 0.25, TYPE_COLORS[code])
        write(
            s,
            x,
            2.15,
            width,
            0.75,
            [{"t": f"{code}  {names[code]}", "size": 13, "color": WHITE, "bold": True, "align": "center"}],
            anchor="middle",
        )
        write(
            s,
            x + 0.08,
            3.15,
            width - 0.16,
            3.1,
            [{"t": item, "size": 15, "color": INK, "bold": True, "align": "center", "after": 3} for item in items],
        )
        x += width + gap
    notes(s, "3 минуты. Ноль — не запрет. Свяжите верхний тип с формулой и направлениями.")


def _close(prs, dark: set):
    s = new_blank(prs, BURGUNDY)
    dark.add(len(prs.slides))
    frame(s, GOLD, pt=1.5)
    write(s, 0.7, 1.5, 12, 0.35, [{"t": "НА ВЫХОД", "size": 14, "color": GOLD, "bold": True}])
    write(
        s,
        0.7,
        2.0,
        12,
        1.0,
        [{"t": "Гипотеза · формула · один шаг", "size": 36, "color": WHITE, "bold": True, "font": SERIF}],
    )
    write(
        s,
        0.7,
        3.3,
        12,
        2.2,
        [
            {
                "t": "1) Четыре клетки: хочу — могу — умею — надо\n"
                "2) Ведущий тип по опроснику\n"
                "3) Если близки люди — вспомните четыре направления\n"
                "4) Мой шаг на месяц: …",
                "size": 20,
                "color": GOLD_SOFT,
            }
        ],
    )
    write(s, 0.7, 5.7, 12, 0.5, [{"t": SLIDE_SIMPLE["close"], "size": 16, "color": WHITE}])
    notes(s, "1 минута. Спасибо.")


def build():
    prs = _prs()
    dark: set[int] = set()
    _hero(prs, dark)
    _map(prs)
    _formula(prs)
    _types(prs)
    _directions(prs)
    _test(prs)
    _key(prs)
    _close(prs, dark)
    return finish(prs, dark)


def main():
    prs = build()
    path = "/workspace/materials/Путь_в_профессию_презентация_20_минут.pptx"
    prs.save(path)
    print(f"slides: {len(prs.slides)}")
    print(path)
    # 60-минутный файл больше не основной: оставляем ту же короткую версию как копию-алиас смысла
    alias = "/workspace/materials/Путь_в_профессию_презентация_60_минут.pptx"
    prs.save(alias)
    print(f"alias: {alias}")


if __name__ == "__main__":
    main()
