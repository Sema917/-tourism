"""Презентация 20 минут: вопросы опросника на слайдах, типы после подсчёта."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from content import (
    DIRECTIONS_SLIDE,
    FORMULA,
    FORMULA_TRAPS,
    MINUTES,
    SLIDE_SIMPLE,
    STICKERS,
    TYPE_BORDERS,
    TYPES,
    VERSION,
)
from ddo import QUESTIONS, answers_by_type, expected_sum, max_score
from pptx_kit import (
    AMBER,
    CORAL,
    INK,
    MUTED,
    SKY,
    TEAL,
    TYPE_COLORS,
    WHITE,
    W,
    H,
    accent_bar,
    add_shape,
    chip,
    eyebrow,
    finish,
    new_blank,
    notes,
    option_row,
    picture,
    sticker,
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
    s = new_blank(prs, INK)
    dark.add(len(prs.slides))
    add_shape(s, 0, 0, 0.18, H, TEAL)
    picture(s, "hero-path.jpg", 7.2, 0, 6.15, H, darken=0.18)
    add_shape(s, 7.2, 0, 0.12, H, CORAL)
    write(s, 0.55, 1.4, 6.3, 0.35, [{"t": "ФЕСТИВАЛЬ  «НАУКА 0+»", "size": 14, "color": TEAL, "bold": True}])
    write(
        s,
        0.55,
        1.95,
        6.4,
        2.0,
        [{"t": "Путь в\nпрофессию", "size": 52, "color": WHITE, "bold": True, "spacing": 0.92}],
    )
    write(
        s,
        0.55,
        4.2,
        6.2,
        1.0,
        [{"t": "Формула выбора, направления подготовки\nи опросник Климова", "size": 18, "color": "B8C4CE"}],
    )
    chip(s, 0.55, 5.55, 2.35, 0.48, f"{MINUTES} МИНУТ", TEAL)
    chip(s, 3.05, 5.55, 2.5, 0.48, "5–11 КЛАСС", "243444", "D5DEE6")
    notes(s, "Короткий старт. Профессию сегодня не назначаем.")


def _map(prs):
    s = new_blank(prs)
    accent_bar(s)
    eyebrow(s, "Карта встречи")
    title(s, "Что успеем за 20 минут")
    write(s, 0.5, 1.45, 12.3, 0.55, [{"t": SLIDE_SIMPLE["agenda"], "size": 16, "color": MUTED}])
    write(s, 0.5, 2.1, 12.3, 0.7, [{"t": SLIDE_SIMPLE["meaning"], "size": 18, "color": INK}])
    steps = [
        ("01", "Формула", "Хочу · могу · умею · надо", TEAL),
        ("02", "Направления", "Четыре пути в вузе", SKY),
        ("03", "Опросник", "15 пар на слайдах", CORAL),
        ("04", "Типы", "Подробный разбор после ключа", AMBER),
    ]
    for i, (n, head, body, color) in enumerate(steps):
        x = 0.5 + i * 3.15
        add_shape(s, x, 3.15, 3.0, 3.25, WHITE, rounded=True, radius=0.12)
        add_shape(s, x, 3.15, 3.0, 0.12, color)
        write(s, x + 0.25, 3.5, 2.5, 0.55, [{"t": n, "size": 28, "color": color, "bold": True}])
        write(s, x + 0.25, 4.25, 2.5, 0.5, [{"t": head, "size": 20, "color": INK, "bold": True}])
        write(s, x + 0.25, 4.95, 2.5, 1.0, [{"t": body, "size": 15, "color": MUTED}])
    notes(s, "1 минута. Типы — после опроса, не до.")


def _formula(prs):
    s = new_blank(prs)
    accent_bar(s, CORAL)
    eyebrow(s, "Интерактив · 4 минуты", CORAL)
    title(s, "Хочу · Могу · Умею · Надо")
    write(s, 0.5, 1.45, 12.3, 0.5, [{"t": FORMULA_TRAPS, "size": 16, "color": MUTED}])
    gap = 0.2
    width = (12.3 - 3 * gap) / 4
    x = 0.5
    for name, color, text in FORMULA:
        add_shape(s, x, 2.2, width, 4.25, color, rounded=True, radius=0.1)
        write(
            s,
            x + 0.22,
            2.5,
            width - 0.44,
            3.7,
            [
                {"t": name, "size": 24, "color": WHITE, "bold": True},
                {"t": text, "size": 15, "color": "F3F7F8", "before": 14},
                {"t": "Запишите в раздатку", "size": 14, "color": WHITE, "bold": True, "before": 24},
            ],
        )
        x += width + gap
    notes(s, "4 минуты. Заполняют графы в раздатке.")


def _directions_overview(prs):
    s = new_blank(prs)
    accent_bar(s, SKY)
    picture(s, "social-directions.jpg", 0.45, 1.55, 4.4, 5.0, darken=0.04)
    add_shape(s, 0.45, 1.55, 4.4, 0.1, SKY)
    eyebrow(s, "Направления подготовки", SKY)
    title(s, "Четыре пути, если близки люди", x=5.15, w=7.7)
    write(s, 5.15, 1.4, 7.7, 0.55, [{"t": SLIDE_SIMPLE["directions"], "size": 15, "color": MUTED}])
    accents = [TEAL, CORAL, SKY, AMBER]
    for i, d in enumerate(DIRECTIONS_SLIDE):
        y = 2.1 + i * 1.1
        add_shape(s, 5.15, y, 7.7, 1.0, WHITE, rounded=True, radius=0.12)
        add_shape(s, 5.15, y, 0.12, 1.0, accents[i])
        write(
            s,
            5.5,
            y + 0.12,
            7.15,
            0.8,
            [
                {"t": f"{d['title']}  ·  {d['code']}", "size": 16, "color": INK, "bold": True},
                {"t": d["ask"], "size": 14, "color": accents[i], "italic": True, "before": 3},
            ],
        )
    notes(s, "Короткий обзор. Дальше — подробные карточки.")


def _direction_detail(prs, items, accents, heading):
    s = new_blank(prs)
    accent_bar(s, accents[0])
    eyebrow(s, "Направления подготовки · подробно", accents[0])
    title(s, heading, size=28)
    for i, d in enumerate(items):
        col = i % 2
        x = 0.45 + col * 6.4
        y = 1.45
        color = accents[i]
        add_shape(s, x, y, 6.2, 5.2, WHITE, rounded=True, radius=0.12)
        add_shape(s, x, y, 6.2, 0.12, color)
        blocks = [
            {"t": f"{d['title']}  ·  {d['code']}", "size": 18, "color": INK, "bold": True},
            {"t": d["ask"], "size": 15, "color": color, "italic": True, "before": 6},
            {"t": d["lead"], "size": 14, "color": INK, "before": 10},
            {"t": "Чему учат", "size": 13, "color": color, "bold": True, "before": 12},
            {"t": d["do"], "size": 13, "color": INK, "before": 2},
            {"t": "Что изучают", "size": 13, "color": color, "bold": True, "before": 10},
            {"t": d["study"], "size": 13, "color": INK, "before": 2},
            {"t": "Куда идут", "size": 13, "color": color, "bold": True, "before": 10},
            {"t": d["work"], "size": 13, "color": MUTED, "before": 2},
            {"t": "Кому близко", "size": 13, "color": color, "bold": True, "before": 10},
            {"t": d["fit"] + " " + d["type_hint"], "size": 13, "color": INK, "before": 2},
        ]
        if d.get("law"):
            blocks.append({"t": d["law"], "size": 12, "color": MUTED, "italic": True, "before": 10})
        write(s, x + 0.3, y + 0.3, 5.6, 4.7, blocks)
    notes(s, "1–1,5 минуты на пару направлений. Не торопитесь.")


def _directions(prs):
    _directions_overview(prs)
    _direction_detail(
        prs,
        DIRECTIONS_SLIDE[:2],
        [TEAL, CORAL],
        "Социология и социальная работа",
    )
    _direction_detail(
        prs,
        DIRECTIONS_SLIDE[2:],
        [SKY, AMBER],
        "Политология и работа с молодёжью",
    )


def _test_intro(prs):
    s = new_blank(prs)
    accent_bar(s, CORAL)
    picture(s, "self-discovery.jpg", 0.45, 1.4, 5.1, 5.15, darken=0.03)
    add_shape(s, 0.45, 1.4, 5.1, 0.1, CORAL)
    eyebrow(s, "Опросник Климова · 15 пар", CORAL)
    title(s, "Дальше — вопросы на слайдах", x=5.85, w=6.9)
    write(s, 5.85, 1.5, 6.9, 1.15, [{"t": SLIDE_SIMPLE["test"], "size": 17, "color": INK}])
    write(
        s,
        5.85,
        2.85,
        6.9,
        0.4,
        [{"t": "Отмечайте ответ в раздатке. Ключ и типы — после всех пар.", "size": 15, "color": MUTED}],
    )
    option_row(s, 5.85, 3.45, 6.9, 0.9, "А", "Первый вариант из пары", TEAL)
    option_row(s, 5.85, 4.5, 6.9, 0.9, "Б", "Второй вариант из пары", CORAL)
    sticker(s, STICKERS[4], 11.55, 5.55, size=1.35, rotate=-12)
    sticker(s, STICKERS[0], 10.15, 5.7, size=1.15, rotate=8)
    notes(s, "Дальше листайте 15 вопросов. Темп ровный, без обсуждения.")


def _question(prs, q: dict, total: int = 15):
    s = new_blank(prs)
    accent_bar(s, TEAL)
    eyebrow(s, f"Опросник  ·  вопрос {q['n']} из {total}", TEAL)
    title(s, f"Пара {q['n']}", size=28)
    write(
        s,
        0.5,
        1.4,
        10.5,
        0.5,
        [{"t": "Выберите только одно занятие — то, что ближе сейчас.", "size": 16, "color": MUTED}],
    )
    option_row(s, 0.5, 2.15, 12.3, 1.65, "А", q["a"], TEAL)
    option_row(s, 0.5, 4.05, 12.3, 1.65, "Б", q["b"], CORAL)
    write(
        s,
        0.5,
        6.0,
        10.2,
        0.4,
        [{"t": "Отметьте А или Б в раздатке. Не спорьте с собой долго.", "size": 14, "color": MUTED}],
    )
    # стикеры в стиле мессенджера (оригинальные, не копипаст паков ВК)
    main = STICKERS[(q["n"] - 1) % len(STICKERS)]
    side = STICKERS[(q["n"]) % len(STICKERS)]
    positions = [
        (11.55, 0.35, 1.4, -14),
        (11.7, 5.55, 1.35, 10),
        (0.35, 5.55, 1.25, -8),
    ]
    px, py, sz, rot = positions[(q["n"] - 1) % len(positions)]
    sticker(s, main, px, py, size=sz, rotate=rot)
    # второй стикер на части слайдов
    if q["n"] % 2 == 0:
        sticker(s, side, 11.45, 1.85, size=1.05, rotate=16)
    notes(s, f"Вопрос {q['n']}. Тишина. 20–25 секунд.")


def _key(prs):
    s = new_blank(prs)
    accent_bar(s, AMBER)
    eyebrow(s, "Считаем", AMBER)
    title(s, "Ключ опросника")
    mx, sm = max_score(False), expected_sum(False)
    write(
        s,
        0.5,
        1.4,
        12.3,
        0.45,
        [{"t": f"Совпадение = 1 балл. Максимум в столбце — {mx}. Сумма — {sm}. Смотрите на 1–2 верхних столбца.", "size": 15, "color": MUTED}],
    )
    buckets = answers_by_type(False)
    gap = 0.16
    width = (12.3 - 4 * gap) / 5
    x = 0.5
    names = {"П": "природа", "Т": "техника", "Ч": "человек", "З": "знаки", "Х": "образ"}
    for code, items in buckets.items():
        add_shape(s, x, 2.05, width, 4.5, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 2.05, width, 0.8, TYPE_COLORS[code], rounded=True, radius=0.1)
        add_shape(s, x, 2.6, width, 0.25, TYPE_COLORS[code])
        write(
            s,
            x,
            2.05,
            width,
            0.8,
            [{"t": f"{code}\n{names[code]}", "size": 14, "color": WHITE, "bold": True, "align": "center"}],
            anchor="middle",
        )
        write(
            s,
            x + 0.1,
            3.15,
            width - 0.2,
            3.15,
            [{"t": item, "size": 15, "color": INK, "bold": True, "align": "center", "after": 3} for item in items],
        )
        x += width + gap
    notes(s, "3 минуты. Дальше — подробный разбор типов.")


def _types_detail(prs):
    s = new_blank(prs)
    accent_bar(s, TEAL)
    eyebrow(s, "После опроса · типы профессий")
    title(s, "Пять типов по предмету труда", size=30)
    write(s, 0.5, 1.35, 12.3, 0.55, [{"t": SLIDE_SIMPLE["types_lead"], "size": 15, "color": MUTED}])
    for i, t in enumerate(TYPES):
        x = 0.4 + i * 2.55
        add_shape(s, x, 2.05, 2.45, 4.55, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 2.05, 2.45, 0.85, TYPE_COLORS[t["code"]], rounded=True, radius=0.1)
        add_shape(s, x, 2.65, 2.45, 0.25, TYPE_COLORS[t["code"]])
        write(
            s,
            x + 0.1,
            2.1,
            2.25,
            0.75,
            [{"t": f"{t['code']}  {t['name'].split('—')[-1].strip().title()}", "size": 13, "color": WHITE, "bold": True, "align": "center"}],
            anchor="middle",
        )
        write(
            s,
            x + 0.12,
            3.1,
            2.2,
            3.3,
            [
                {"t": "Предмет", "size": 11, "color": TYPE_COLORS[t["code"]], "bold": True},
                {"t": t["object"], "size": 11, "color": INK, "before": 2},
                {"t": "Ближе тем, кому нравится", "size": 11, "color": TYPE_COLORS[t["code"]], "bold": True, "before": 8},
                {"t": t["likes"], "size": 11, "color": INK, "before": 2},
                {"t": "Примеры", "size": 11, "color": TYPE_COLORS[t["code"]], "bold": True, "before": 8},
                {"t": t["examples"], "size": 11, "color": MUTED, "before": 2},
            ],
        )
    notes(s, "2 минуты. Не читайте всё подряд — укажите на верхний столбец участника.")


def _types_notes(prs):
    s = new_blank(prs)
    accent_bar(s, SKY)
    eyebrow(s, "Как читать результат", SKY)
    title(s, "Границы типов и связь с направлениями")
    write(s, 0.5, 1.45, 12.3, 1.1, [{"t": TYPE_BORDERS, "size": 17, "color": INK}])
    cards = [
        ("П · природа", TYPES[0]["note"], TYPE_COLORS["П"]),
        ("Т · техника", TYPES[1]["note"], TYPE_COLORS["Т"]),
        ("Ч · человек", TYPES[2]["note"], TYPE_COLORS["Ч"]),
        ("З · знаки", TYPES[3]["note"], TYPE_COLORS["З"]),
        ("Х · образ", TYPES[4]["note"], TYPE_COLORS["Х"]),
    ]
    for i, (head, body, color) in enumerate(cards):
        y = 2.75 + i * 0.75
        add_shape(s, 0.5, y, 12.3, 0.68, WHITE, rounded=True, radius=0.1)
        add_shape(s, 0.5, y, 0.12, 0.68, color)
        write(
            s,
            0.85,
            y,
            11.7,
            0.68,
            [{"t": head, "size": 14, "color": color, "bold": True}, {"t": "  ·  " + body, "size": 13, "color": INK}],
            anchor="middle",
        )
    notes(s, "1–2 минуты. Свяжите тип «человек/знаки» с четырьмя направлениями.")


def _close(prs, dark: set):
    s = new_blank(prs, INK)
    dark.add(len(prs.slides))
    add_shape(s, 0, 0, 0.18, H, TEAL)
    write(s, 0.7, 1.5, 12, 0.35, [{"t": "НА ВЫХОД", "size": 14, "color": TEAL, "bold": True}])
    write(
        s,
        0.7,
        2.05,
        12,
        0.9,
        [{"t": "Гипотеза · формула · один шаг", "size": 34, "color": WHITE, "bold": True}],
    )
    write(
        s,
        0.7,
        3.2,
        12,
        2.3,
        [
            {
                "t": "1) Четыре клетки: хочу — могу — умею — надо\n"
                "2) Ведущий тип по опроснику и его смысл\n"
                "3) Если близки люди — вспомните четыре направления\n"
                "4) Мой шаг на месяц: …",
                "size": 20,
                "color": "C5D0D8",
            }
        ],
    )
    write(s, 0.7, 5.7, 12, 0.5, [{"t": SLIDE_SIMPLE["close"], "size": 16, "color": WHITE}])
    notes(s, "Финал.")


def build():
    prs = _prs()
    dark: set[int] = set()
    _hero(prs, dark)
    _map(prs)
    _formula(prs)
    _directions(prs)
    _test_intro(prs)
    for q in QUESTIONS:
        _question(prs, q)
    _key(prs)
    _types_detail(prs)
    _types_notes(prs)
    _close(prs, dark)
    return finish(prs, dark)


def main():
    prs = build()
    path = "/workspace/materials/Путь_в_профессию_презентация_20_минут.pptx"
    prs.save(path)
    print(f"slides: {len(prs.slides)}")
    print(path)
    alias = "/workspace/materials/Путь_в_профессию_презентация_60_минут.pptx"
    prs.save(alias)
    print(f"alias: {alias}")


if __name__ == "__main__":
    main()
