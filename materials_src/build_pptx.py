"""Презентация под дословный сценарий: единый дизайн, мало текста, воздух."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from content import DIRECTIONS_CARD, FORMULA, MINUTES, SLIDE_SIMPLE, STICKERS, VERSION
from ddo import QUESTIONS, answers_by_type, expected_sum, max_score
from pptx_kit import (
    AMBER,
    CORAL,
    INK,
    LINE,
    MUTED,
    SKY,
    SOFT,
    TEAL,
    TYPE_COLORS,
    WHITE,
    W,
    H,
    add_shape,
    chip,
    eyebrow,
    finish,
    new_blank,
    notes,
    option_row,
    picture,
    rail,
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
    picture(s, "hero-path.jpg", 6.9, 0, 6.45, H, darken=0.22)
    add_shape(s, 0, 0, 7.05, H, INK)
    add_shape(s, 6.95, 0, 0.12, H, TEAL)
    write(s, 0.65, 1.55, 5.9, 0.35, [{"t": "ФЕСТИВАЛЬ  «НАУКА 0+»", "size": 13, "color": TEAL, "bold": True}])
    write(s, 0.65, 2.1, 5.9, 1.9, [{"t": "Путь в\nпрофессию", "size": 50, "color": WHITE, "bold": True, "spacing": 0.92}])
    write(
        s,
        0.65,
        4.25,
        5.8,
        0.8,
        [{"t": "Как думать о будущей работе\nбез громких названий", "size": 18, "color": "B7C3CE"}],
    )
    chip(s, 0.65, 5.5, 2.2, 0.46, f"{MINUTES} МИНУТ", TEAL)
    chip(s, 3.0, 5.5, 2.35, 0.46, "5–11 КЛАСС", "243444", "D7E0E8")
    notes(s, "Всем привет… Сегодня не угадываем должность.")


def _map(prs):
    s = new_blank(prs)
    rail(s, TEAL)
    picture(s, "slide-notebook.jpg", 8.55, 1.35, 4.25, 5.15, darken=0.02)
    add_shape(s, 8.55, 1.35, 4.25, 0.1, TEAL)
    eyebrow(s, "Карта встречи")
    title(s, "Что сделаем за 20 минут", w=7.7)
    write(s, 0.55, 1.5, 7.6, 0.7, [{"t": SLIDE_SIMPLE["meaning"], "size": 17, "color": MUTED}])
    steps = [
        ("1", "Четыре клетки", "хочу · могу · умею · надо", TEAL),
        ("2", "Направления", "если интересны люди", SKY),
        ("3", "Опрос", "15 пар на слайдах", CORAL),
        ("4", "Баллы и шаг", "что значат типы", AMBER),
    ]
    for i, (n, head, body, color) in enumerate(steps):
        y = 2.45 + i * 1.0
        add_shape(s, 0.55, y, 7.6, 0.88, WHITE, rounded=True, radius=0.12)
        add_shape(s, 0.55, y, 0.12, 0.88, color)
        write(s, 0.9, y, 0.7, 0.88, [{"t": n, "size": 24, "color": color, "bold": True}], anchor="middle")
        write(
            s,
            1.7,
            y,
            6.1,
            0.88,
            [{"t": head, "size": 18, "color": INK, "bold": True}, {"t": "  ·  " + body, "size": 16, "color": MUTED}],
            anchor="middle",
        )
    notes(s, "Лист ваш. Сдавать не нужно.")


def _formula(prs):
    s = new_blank(prs)
    rail(s, CORAL)
    eyebrow(s, "Четыре клетки · 4 минуты", CORAL)
    title(s, "Хочу · Могу · Умею · Надо")
    write(
        s,
        0.55,
        1.45,
        12.2,
        0.45,
        [{"t": "Пишите занятие, не должность. Клетки должны хоть немного пересекаться.", "size": 16, "color": MUTED}],
    )
    gap = 0.22
    width = (12.2 - 3 * gap) / 4
    x = 0.55
    for name, color, text in FORMULA:
        add_shape(s, x, 2.15, width, 4.25, color, rounded=True, radius=0.1)
        write(
            s,
            x + 0.22,
            2.45,
            width - 0.44,
            3.7,
            [
                {"t": name, "size": 26, "color": WHITE, "bold": True},
                {"t": text, "size": 15, "color": "F4F8F9", "before": 16, "spacing": 1.12},
                {"t": "В раздатку →", "size": 14, "color": WHITE, "bold": True, "before": 28},
            ],
        )
        x += width + gap
    notes(s, "Откройте верх листа… Четыре минуты.")


def _directions_overview(prs):
    s = new_blank(prs)
    rail(s, SKY)
    picture(s, "slide-people.jpg", 0.5, 1.55, 4.6, 5.0, darken=0.03)
    add_shape(s, 0.5, 1.55, 4.6, 0.1, SKY)
    eyebrow(s, "Направления учёбы", SKY)
    title(s, "Если интересны люди", x=5.4, w=7.4)
    write(s, 5.4, 1.45, 7.4, 0.5, [{"t": SLIDE_SIMPLE["directions"], "size": 16, "color": MUTED}])
    accents = [TEAL, CORAL, SKY, AMBER]
    for i, d in enumerate(DIRECTIONS_CARD):
        y = 2.15 + i * 1.1
        add_shape(s, 5.4, y, 7.4, 0.98, WHITE, rounded=True, radius=0.12)
        add_shape(s, 5.4, y, 0.12, 0.98, accents[i])
        write(
            s,
            5.75,
            y + 0.14,
            6.85,
            0.75,
            [
                {"t": d["title"], "size": 17, "color": INK, "bold": True},
                {"t": d["ask"], "size": 15, "color": accents[i], "italic": True, "before": 4},
            ],
        )
    notes(s, "Четыре дороги… Коды запоминать не надо.")


def _directions_pair(prs, pair, accents, heading, images):
    s = new_blank(prs)
    rail(s, accents[0])
    eyebrow(s, "Направления · коротко", accents[0])
    title(s, heading, size=30)
    for i, d in enumerate(pair):
        x = 0.5 + i * 6.4
        add_shape(s, x, 1.5, 6.15, 5.15, WHITE, rounded=True, radius=0.12)
        picture(s, images[i], x + 0.2, 1.7, 5.75, 1.55, darken=0.04)
        add_shape(s, x + 0.2, 1.7, 5.75, 0.1, accents[i])
        write(s, x + 0.35, 3.4, 5.45, 0.4, [{"t": d["title"], "size": 20, "color": INK, "bold": True}])
        write(s, x + 0.35, 3.85, 5.45, 0.55, [{"t": d["ask"], "size": 14, "color": accents[i], "italic": True}])
        for j, line in enumerate(d["lines"]):
            y = 4.5 + j * 0.65
            write(s, x + 0.35, y, 5.45, 0.55, [{"t": "·  " + line, "size": 14, "color": INK}], anchor="middle")
    notes(s, "Подробности — устно по сценарию. На слайде только опоры.")


def _test_intro(prs):
    s = new_blank(prs)
    rail(s, CORAL)
    picture(s, "self-discovery.jpg", 0.5, 1.5, 5.0, 5.05, darken=0.02)
    add_shape(s, 0.5, 1.5, 5.0, 0.1, CORAL)
    eyebrow(s, "Опрос · 15 пар", CORAL)
    title(s, "Дальше вопросы на экране", x=5.85, w=6.9)
    write(s, 5.85, 1.5, 6.9, 1.0, [{"t": SLIDE_SIMPLE["test"], "size": 18, "color": INK}])
    rules = [
        "Отмечайте А или Б в раздатке",
        "Долго не думайте",
        "Ключ откроем после всех пар",
    ]
    for i, rule in enumerate(rules):
        y = 2.8 + i * 0.95
        add_shape(s, 5.85, y, 6.9, 0.82, WHITE, rounded=True, radius=0.12)
        add_shape(s, 5.85, y, 0.12, 0.82, [TEAL, CORAL, AMBER][i])
        write(s, 6.2, y, 6.3, 0.82, [{"t": rule, "size": 17, "color": INK}], anchor="middle")
    sticker(s, STICKERS[4], 11.7, 5.75, size=1.15, rotate=-10)
    notes(s, "Готовы? Тишина. Поехали.")


def _question(prs, q: dict, total: int = 15):
    s = new_blank(prs)
    rail(s, TEAL)
    eyebrow(s, f"Опрос  ·  вопрос {q['n']} из {total}", TEAL)
    title(s, f"Пара {q['n']}", size=30)
    write(s, 0.55, 1.4, 10.4, 0.4, [{"t": "Выберите одно занятие — то, что ближе сейчас.", "size": 16, "color": MUTED}])
    option_row(s, 0.55, 2.1, 12.2, 1.75, "А", q["a"], TEAL)
    option_row(s, 0.55, 4.15, 12.2, 1.75, "Б", q["b"], CORAL)
    write(s, 0.55, 6.15, 10.5, 0.35, [{"t": "Отметьте ответ в раздатке.", "size": 14, "color": MUTED}])
    sticker(s, STICKERS[(q["n"] - 1) % len(STICKERS)], 11.75, 0.35, size=1.15, rotate=-8 if q["n"] % 2 else 10)
    notes(s, f"Пара {q['n']}. 20–25 секунд. Тишина.")


def _key(prs):
    s = new_blank(prs)
    rail(s, AMBER)
    eyebrow(s, "Считаем", AMBER)
    title(s, "Ключ опроса")
    write(
        s,
        0.55,
        1.4,
        12.2,
        0.4,
        [{"t": f"Совпадение = 1 балл. В столбце максимум {max_score(False)}. Сумма — {expected_sum(False)}.", "size": 15, "color": MUTED}],
    )
    buckets = answers_by_type(False)
    names = {"П": "природа", "Т": "техника", "Ч": "человек", "З": "знаки", "Х": "образ"}
    gap = 0.18
    width = (12.2 - 4 * gap) / 5
    x = 0.55
    for code, items in buckets.items():
        add_shape(s, x, 2.0, width, 4.5, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 2.0, width, 0.85, TYPE_COLORS[code], rounded=True, radius=0.1)
        add_shape(s, x, 2.6, width, 0.25, TYPE_COLORS[code])
        write(
            s,
            x,
            2.0,
            width,
            0.85,
            [{"t": f"{code}\n{names[code]}", "size": 14, "color": WHITE, "bold": True, "align": "center"}],
            anchor="middle",
        )
        write(
            s,
            x + 0.08,
            3.15,
            width - 0.16,
            3.1,
            [{"t": item, "size": 15, "color": INK, "bold": True, "align": "center", "after": 4} for item in items],
        )
        x += width + gap
    notes(s, "Смотрите на один-два верхних столбца.")


def _types(prs):
    s = new_blank(prs)
    rail(s, TEAL)
    eyebrow(s, "После опроса")
    title(s, "Пять типов — коротко")
    write(s, 0.55, 1.4, 12.2, 0.4, [{"t": SLIDE_SIMPLE["types_lead"], "size": 16, "color": MUTED}])
    short = [
        ("П", "Природа", "Живое: ухаживать, наблюдать, выращивать", "ветеринар · эколог · биолог"),
        ("Т", "Техника", "Вещи и машины: собирать, чинить, понимать", "инженер · строитель · механик"),
        ("Ч", "Человек", "Люди: помочь, объяснить, договориться", "врач · учитель · соцработа"),
        ("З", "Знаки", "Тексты, цифры, схемы, правила", "аналитик · социолог · программист"),
        ("Х", "Образ", "Сделать вид, звук, оформление", "дизайнер · актёр · музыкант"),
    ]
    for i, (code, name, body, examples) in enumerate(short):
        x = 0.45 + i * 2.55
        add_shape(s, x, 2.05, 2.4, 4.4, WHITE, rounded=True, radius=0.1)
        add_shape(s, x, 2.05, 2.4, 1.0, TYPE_COLORS[code], rounded=True, radius=0.1)
        add_shape(s, x, 2.8, 2.4, 0.25, TYPE_COLORS[code])
        write(
            s,
            x,
            2.05,
            2.4,
            1.0,
            [{"t": f"{code}\n{name}", "size": 18, "color": WHITE, "bold": True, "align": "center"}],
            anchor="middle",
        )
        write(
            s,
            x + 0.18,
            3.25,
            2.05,
            2.9,
            [
                {"t": body, "size": 15, "color": INK, "align": "center"},
                {"t": examples, "size": 13, "color": MUTED, "align": "center", "before": 18},
            ],
        )
    notes(s, "Не приговор… Природа — живое…")


def _types_link(prs):
    s = new_blank(prs)
    rail(s, SKY)
    picture(s, "slide-types.jpg", 8.6, 1.5, 4.2, 5.0, darken=0.02)
    add_shape(s, 8.6, 1.5, 4.2, 0.1, SKY)
    eyebrow(s, "Как читать результат", SKY)
    title(s, "Границы типов мягкие", w=7.8)
    points = [
        "Врач рядом и с человеком, и с природой",
        "Архитектор — между образом и техникой",
        "Социолог — между людьми и знаками",
        "Если сверху «человек» или «знаки» — вспомните четыре направления",
        "Техника или природа — другие двери, и они не хуже",
    ]
    for i, line in enumerate(points):
        y = 1.55 + i * 0.95
        add_shape(s, 0.55, y, 7.7, 0.82, WHITE, rounded=True, radius=0.12)
        add_shape(s, 0.55, y, 0.12, 0.82, [TEAL, CORAL, SKY, AMBER, TYPE_COLORS["П"]][i])
        write(s, 0.9, y, 7.1, 0.82, [{"t": line, "size": 16, "color": INK}], anchor="middle")
    notes(s, "Границы размыты…")


def _close(prs, dark: set):
    s = new_blank(prs, INK)
    dark.add(len(prs.slides))
    add_shape(s, 0, 0, 0.12, H, TEAL)
    write(s, 0.7, 1.6, 12, 0.35, [{"t": "НА ВЫХОД", "size": 13, "color": TEAL, "bold": True}])
    write(s, 0.7, 2.15, 12, 0.8, [{"t": "Что унести с собой", "size": 36, "color": WHITE, "bold": True}])
    items = [
        "Четыре клетки с листа",
        "Один-два типа по опросу",
        "Один маленький шаг на месяц",
    ]
    for i, item in enumerate(items):
        y = 3.3 + i * 0.85
        add_shape(s, 0.7, y, 11.9, 0.7, "243444", rounded=True, radius=0.12)
        write(s, 1.0, y, 11.3, 0.7, [{"t": f"{i+1}.  {item}", "size": 20, "color": WHITE}], anchor="middle")
    write(s, 0.7, 6.1, 12, 0.4, [{"t": SLIDE_SIMPLE["close"], "size": 15, "color": "A8B6C2"}])
    notes(s, "Запишите шаг… Спасибо.")


def build():
    prs = _prs()
    dark: set[int] = set()
    _hero(prs, dark)
    _map(prs)
    _formula(prs)
    _directions_overview(prs)
    _directions_pair(
        prs,
        DIRECTIONS_CARD[:2],
        [TEAL, CORAL],
        "Социология и социальная работа",
        ["slide-people.jpg", "youth-support.jpg"],
    )
    _directions_pair(
        prs,
        DIRECTIONS_CARD[2:],
        [SKY, AMBER],
        "Политология и работа с молодёжью",
        ["labor-market.jpg", "social-directions.jpg"],
    )
    _test_intro(prs)
    for q in QUESTIONS:
        _question(prs, q)
    _key(prs)
    _types(prs)
    _types_link(prs)
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
