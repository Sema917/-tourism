"""Две презентации: 60 минут и 20 минут."""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches

from content import (
    ADMISSION_COMMON,
    CAN_OPTIONS,
    CASE,
    EVENT,
    FORMULA,
    FORMULA_TRAPS,
    LAB_FACT,
    MYTHS,
    PATHS,
    PROGRAMS,
    SKILLS,
    SUPPORT,
    TYPE_BORDERS,
    TYPES,
    VERSION,
)
from ddo import answers_by_type, expected_sum, max_score
from pptx_kit import (
    AMBER,
    AMBER_BG,
    CREAM,
    INDIGO,
    INK,
    MUTED,
    NAVY,
    NAVY_DEEP,
    SOFT,
    TEAL,
    TYPE_COLORS,
    WHITE,
    add_shape,
    chip,
    eyebrow,
    footer,
    new_slide,
    notes,
    title,
    write,
)

W = 13.333
H = 7.5


def _prs():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = EVENT
    prs.core_properties.author = "Комплект «Путь в профессию», фестиваль «Наука 0+»"
    prs.core_properties.subject = f"Профориентационная встреча. Комплект {VERSION}"
    prs.core_properties.language = "ru-RU"
    prs.core_properties.category = "Наука 0+"
    return prs


def _finish(prs, dark_indexes):
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, start=1):
        footer(slide, i, total, dark=(i in dark_indexes))
    return prs


def _formula_row(slide, y=4.55, h=1.35):
    gap = 0.16
    width = (12.33 - 3 * gap) / 4
    x = 0.5
    for name, color, text in FORMULA:
        add_shape(slide, x, y, width, h, color, rounded=True, radius=0.08)
        write(
            slide,
            x + 0.16,
            y + 0.12,
            width - 0.32,
            h - 0.2,
            [
                {"t": name, "size": 20, "color": WHITE if color != AMBER else NAVY_DEEP, "bold": True},
                {"t": text, "size": 13, "color": WHITE if color != AMBER else NAVY_DEEP, "before": 4, "spacing": 0.95},
            ],
        )
        x += width + gap


def _section(prs, number, kicker, heading, sentence, chip_text, chip_fill):
    slide = new_slide(prs, dark=True)
    add_shape(slide, 0, 0, 0.22, H, chip_fill)
    write(slide, 0.7, 1.55, 4, 0.4, [{"t": "ТЕМА", "size": 14, "color": "8FB4C4", "bold": True}])
    write(slide, 0.62, 1.9, 5, 1.5, [{"t": number, "size": 84, "color": WHITE, "bold": True}])
    write(slide, 6.1, 2.15, 6.5, 0.35, [{"t": kicker.upper(), "size": 14, "color": chip_fill, "bold": True}])
    write(slide, 6.1, 2.55, 6.5, 1.5, [{"t": heading, "size": 36, "color": WHITE, "bold": True, "spacing": 0.95}])
    write(slide, 6.1, 4.25, 6.4, 1.1, [{"t": sentence, "size": 18, "color": "E4EEF2", "spacing": 1.0}])
    chip(slide, 6.1, 5.55, 2.3, 0.46, chip_text, chip_fill, NAVY_DEEP if chip_fill == AMBER_BG else WHITE, size=16)
    return slide


def _key_columns(slide, express, y=1.85):
    buckets = answers_by_type(express)
    gap = 0.14
    width = (12.33 - 4 * gap) / 5
    count = len(next(iter(buckets.values())))
    card_h = 1.2 + count * 0.46
    x = 0.5
    for code in buckets:
        color = TYPE_COLORS[code]
        add_shape(slide, x, y, width, card_h, WHITE, rounded=True, radius=0.06)
        add_shape(slide, x, y, width, 0.82, color, rounded=True, radius=0.06)
        add_shape(slide, x, y + 0.5, width, 0.32, color)
        name = {"П": "природа", "Т": "техника", "Ч": "человек", "З": "знаки", "Х": "образ"}[code]
        write(
            slide,
            x + 0.08,
            y + 0.06,
            width - 0.16,
            0.72,
            [
                {"t": code, "size": 22, "color": WHITE, "bold": True, "align": "center"},
                {"t": name, "size": 13, "color": WHITE, "align": "center", "before": 0},
            ],
        )
        write(
            slide,
            x + 0.08,
            y + 1.05,
            width - 0.16,
            3.1,
            [{"t": item, "size": 20, "color": INK, "bold": True, "align": "center", "after": 6} for item in buckets[code]],
        )
        x += width + gap


def build_hour():
    prs = _prs()
    prs.core_properties.title = "Путь в профессию · 60 минут"
    dark = set()

    # 1 Title
    s = new_slide(prs, dark=True)
    dark.add(1)
    write(s, 0.7, 1.15, 10, 0.32, [{"t": "ФЕСТИВАЛЬ  «НАУКА 0+»", "size": 15, "color": "7DCEC4", "bold": True}])
    write(s, 0.7, 1.65, 12, 1.3, [{"t": "Путь в профессию", "size": 60, "color": WHITE, "bold": True}])
    write(
        s,
        0.7,
        3.15,
        10,
        0.7,
        [{"t": "Как выбрать дело по формуле, а не по чужому ожиданию", "size": 22, "color": "D5E4EA"}],
    )
    chip(s, 0.7, 4.15, 2.15, 0.46, "60 МИНУТ", TEAL, WHITE, size=15)
    chip(s, 3.0, 4.15, 2.5, 0.46, "5–11 КЛАСС", "245066", WHITE, size=15)
    _formula_pills(s, 0.7, 5.15)
    notes(
        s,
        "0:00. Поздоровайтесь. Скажите, что это исследование, а не вердикт. "
        "К концу часа — формула из четырёх слов и один шаг на месяц, не «единственная профессия».",
    )

    # 2 Rules
    s = new_slide(prs)
    eyebrow(s, "Договор на этот час")
    title(s, "Четыре правила, без которых тест вреден")
    rules = [
        ("01", "Это гипотеза", "Опросник не ставит оценку и не назначает профессию."),
        ("02", "Баллы не собираем", "Вслух — только по собственному желанию. По рядам не идём."),
        ("03", "Тетрадь ваша", "Её не сдают. Имя можно не писать."),
        ("04", "Молчать можно", "Ответить — тоже можно. Оба выбора взрослые."),
    ]
    for i, (n, head, body) in enumerate(rules):
        x = 0.5 + (i % 4) * 3.2
        add_shape(s, x, 1.7, 3.02, 3.55, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.2, 1.95, 2.6, 0.5, [{"t": n, "size": 22, "color": TEAL, "bold": True}])
        write(s, x + 0.2, 2.6, 2.6, 1.1, [{"t": head, "size": 22, "color": INK, "bold": True, "spacing": 0.95}])
        write(s, x + 0.2, 3.85, 2.6, 1.15, [{"t": body, "size": 15, "color": MUTED, "spacing": 0.98}])
    notes(s, "0:01. Проговорите все четыре правила. Особенно первое и второе. Не пропускайте даже при нехватке времени.")

    # 3 Poll
    s = new_slide(prs)
    eyebrow(s, "Короткая разведка")
    title(s, "Поднимите руку, если…")
    polls = [
        "уже всерьёз думали о профессии",
        "слышали «выбери нормальную профессию»",
        "умеете что-то, чему не учат на обычном уроке",
        "хотите работу, где помогают людям",
    ]
    for i, line in enumerate(polls):
        y = 1.65 + i * 1.15
        add_shape(s, 0.5, y, 12.33, 1.0, WHITE, rounded=True, radius=0.1)
        add_shape(s, 0.5, y, 0.18, 1.0, TEAL if i % 2 == 0 else INDIGO)
        write(s, 1.0, y, 11.4, 1.0, [{"t": line, "size": 24, "color": INK, "bold": True}], anchor="middle")
    notes(s, "0:03. По одной фразе, руки опускать. Финал: «В зале нет правильной руки». 90 секунд, не больше.")

    # 4 Roadmap
    s = new_slide(prs)
    eyebrow(s, "Карта часа")
    title(s, "Четыре темы — четыре части одного выбора")
    themes = [
        ("01", "Самопознание", "Что вам интересно и как это увидеть методом, а не ярлыком.", "Хочу"),
        ("02", "Рынок труда", "Что нужно людям — без моды и без запугивания.", "Надо"),
        ("03", "Учёба и навыки", "Какими путями идут и что пригождается почти в любой работе.", "Могу"),
        ("04", "План и поддержка", "Один шаг на месяц и двери, которые уже открыты.", "Умею"),
    ]
    colors = ["E15A3A", "3A46A0", "0F7A6C", "A86B12"]
    for i, (n, head, body, word) in enumerate(themes):
        y = 1.55 + i * 1.28
        add_shape(s, 0.5, y, 12.33, 1.14, WHITE, rounded=True, radius=0.08)
        add_shape(s, 0.5, y, 1.3, 1.14, colors[i])
        write(s, 0.5, y, 1.3, 1.14, [{"t": n, "size": 22, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, 2.05, y + 0.14, 8.3, 0.86, [
            {"t": head, "size": 22, "color": INK, "bold": True},
            {"t": body, "size": 15, "color": MUTED, "before": 2},
        ])
        write(s, 10.5, y, 2.1, 1.14, [{"t": word, "size": 16, "color": colors[i], "bold": True, "align": "center"}], anchor="middle")
    notes(s, "0:05. Один проход по карте. Не комментируйте каждое слово. Скажите, что всё сложится в формулу.")

    # 5 Formula
    s = new_slide(prs)
    eyebrow(s, "Сквозной интерактив")
    title(s, "Формула профессии", size=34)
    write(s, 0.5, 1.25, 12.3, 0.4, [{"t": "Хочу  ·  Могу  ·  Умею  ·  Надо", "size": 18, "color": MUTED, "bold": True}])
    _formula_row(s, 1.8, 3.15)
    add_shape(s, 0.5, 5.2, 12.33, 1.55, NAVY, rounded=True, radius=0.08)
    write(
        s,
        0.75,
        5.38,
        11.85,
        1.25,
        [{"t": FORMULA_TRAPS, "size": 18, "color": WHITE, "spacing": 1.0}],
        anchor="middle",
    )
    notes(
        s,
        "0:06–0:08. Разведите четыре слова. Ловушки проговорите вслух: только хочу — мечта, только надо — чужой выбор. "
        "Попросите открыть в тетради блок «Формула».",
    )

    # 6 Section want
    s = _section(
        prs,
        "01",
        "Самопознание и профессиональные интересы",
        "Сначала предмет\nтруда, не должность",
        "Метод Климова спрашивает не «кем быть», а «с чем вам интереснее иметь дело».",
        "ХОЧУ",
        "E15A3A",
    )
    dark.add(6)
    notes(s, "10 секунд на вдох. Не задерживайтесь. Сразу к заданию в парах.")

    # 7 Pair
    s = new_slide(prs)
    eyebrow(s, "Задание  ·  2 минуты", CORAL := "E15A3A")
    title(s, "Какое дело заставляет забыть о времени?")
    write(
        s,
        0.5,
        1.4,
        12.3,
        0.7,
        [{"t": "Не должность. Занятие: чинить, объяснять, считать, ухаживать, спорить по правилам, собирать людей, рисовать.", "size": 18, "color": INK}],
    )
    cards = [
        ("1", "Сосед", "Две минуты вдвоём. Кто не хочет говорить — пишет молча. Это тоже ответ."),
        ("2", "Тетрадь", "Одна строка в клетку «Хочу». Своими словами, без красивой профессии."),
        ("3", "Стоп", "Должность «юрист» или «блогер» годится, только если за ней видно дело."),
    ]
    for i, (n, head, body) in enumerate(cards):
        x = 0.5 + i * 4.2
        add_shape(s, x, 2.4, 3.98, 3.3, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.25, 2.65, 3.4, 0.5, [{"t": n, "size": 28, "color": "E15A3A", "bold": True}])
        write(s, x + 0.25, 3.3, 3.4, 0.5, [{"t": head, "size": 24, "color": INK, "bold": True}])
        write(s, x + 0.25, 4.0, 3.4, 1.4, [{"t": body, "size": 16, "color": MUTED, "spacing": 1.0}])
    notes(s, "0:08–0:11. Засеките 2 минуты. По рядам не собирайте ответы. Напомните про блок «Формула» в тетради.")

    # 8 DDO intro
    s = new_slide(prs)
    eyebrow(s, "Метод  ·  Е. А. Климов")
    title(s, "Одинаковые вопросы. Общий ключ.")
    points = [
        ("15 пар", "В каждой одно занятие из двух. Сегодня делаем все, не только блок «экспресс»."),
        ("6 минут", "Первый отклик. Представьте, что научиться можно всему."),
        ("Не диагноз", "Ключ переводит выбор на язык типов. Он не знает «лучшую» профессию."),
        ("Тишина", "Непонятное слово — тихая рука. Ведущий подойдёт, на зал не объясняет."),
    ]
    for i, (head, body) in enumerate(points):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 6.4
        y = 1.6 + row * 2.35
        add_shape(s, x, y, 6.15, 2.15, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.28, y + 0.28, 5.6, 0.5, [{"t": head, "size": 26, "color": NAVY, "bold": True}])
        write(s, x + 0.28, y + 0.95, 5.6, 0.9, [{"t": body, "size": 16, "color": INK}])
    notes(
        s,
        "0:11–0:13. Назовите метод: дифференциально-диагностический опросник, предмет труда. "
        "Пять типов назовёте после, не до. Попросите не открывать страницу ключа.",
    )

    # 9 Timer
    s = new_slide(prs, dark=True)
    dark.add(9)
    write(s, 0.7, 1.3, 12, 0.4, [{"t": "ОПРОСНИК ДДО  ·  ВОПРОСЫ 1–15", "size": 16, "color": "7DCEC4", "bold": True}])
    write(s, 0.7, 1.9, 12, 1.4, [{"t": "6:00", "size": 92, "color": WHITE, "bold": True}])
    write(s, 0.7, 3.7, 11, 1.2, [{"t": "Отметьте А или Б.\nОдин выбор в паре. Без ключа.", "size": 28, "color": "E4EEF2", "spacing": 1.0}])
    write(s, 0.7, 5.5, 11, 0.8, [{"t": "За минуту до конца: поставьте выбор даже там, где сомневаетесь.", "size": 18, "color": "F3C15B"}])
    notes(s, "0:13–0:19. Молчите. Ходите по рядам. Подсказывайте значения слов, не выборы. На 5-й минуте — фраза про сомнения.")

    # 10 Key
    s = new_slide(prs)
    eyebrow(s, "Считаем вместе")
    title(s, "Ключ сокращённого варианта", size=30)
    write(
        s,
        0.5,
        1.25,
        12.3,
        0.4,
        [{"t": f"Один балл за каждое совпадение. Максимум в столбце — {max_score(False)}. Сумма пяти чисел — {expected_sum(False)}.", "size": 16, "color": MUTED}],
    )
    _key_columns(s, express=False, y=1.75)
    add_shape(s, 0.5, 6.15, 12.33, 0.78, NAVY, rounded=True, radius=0.1)
    write(
        s,
        0.75,
        6.15,
        11.85,
        0.78,
        [{"t": "Разница в 1 балл — не вывод. Ноль — не запрет. Если сумма не равна 15, ищите пропуск.", "size": 16, "color": WHITE}],
        anchor="middle",
    )
    notes(s, "0:19. Дайте время перенести. Сумма 15 — самопроверка. Не комментируйте чужие столбцы. Кто сбился — помогите индивидуально. Буква Б заглавная, чтобы зал не прочитал её как цифру 6.")

    # 11 Read
    s = new_slide(prs)
    eyebrow(s, "Как читать числа")
    title(s, "Балл — не судьба")
    bands = [
        ("0–1", "Интерес в этих вопросах почти не проявился."),
        ("2–3", "Интерес есть, но он не главный."),
        ("4–5", "Интерес выраженный."),
        ("6", "Интерес выражен очень ярко."),
    ]
    for i, (n, body) in enumerate(bands):
        y = 1.5 + i * 1.05
        add_shape(s, 0.5, y, 7.3, 0.92, WHITE, rounded=True, radius=0.1)
        write(s, 0.7, y, 1.5, 0.92, [{"t": n, "size": 24, "color": TEAL, "bold": True}], anchor="middle")
        write(s, 2.3, y, 5.2, 0.92, [{"t": body, "size": 16, "color": INK}], anchor="middle")
    add_shape(s, 8.05, 1.5, 4.78, 4.2, NAVY, rounded=True, radius=0.08)
    write(
        s,
        8.3,
        1.75,
        4.3,
        3.8,
        [
            {"t": "Три оговорки", "size": 18, "color": "7DCEC4", "bold": True},
            {"t": "Разница в 1 балл почти ничего не значит.", "size": 16, "color": WHITE, "before": 12},
            {"t": "Ноль не значит «вам нельзя».", "size": 16, "color": WHITE, "before": 10},
            {"t": "Ровный профиль — повод пробовать разное, а не поломка.", "size": 16, "color": WHITE, "before": 10},
            {"t": "Шкала учебная, для этого комплекта из 15 пар.", "size": 16, "color": "F3C15B", "before": 10},
        ],
    )
    notes(
        s,
        "Смотрите на 1–2 верхних столбца. Запишите их в «Хочу». Шкалу не выдавайте за клиническую норму полного ДДО. "
        "Опросник не измеряет талант, доброту, здоровье и рынок труда.",
    )

    # 12 Types
    s = new_slide(prs)
    eyebrow(s, "Пять типов по предмету труда")
    title(s, "С чем вам интереснее иметь дело", size=30)
    for i, t in enumerate(TYPES):
        y = 1.42 + i * 1.02
        add_shape(s, 0.5, y, 12.33, 0.94, WHITE, rounded=True, radius=0.08)
        add_shape(s, 0.5, y, 0.7, 0.94, TYPE_COLORS[t["code"]])
        write(s, 0.5, y, 0.7, 0.94, [{"t": t["code"], "size": 18, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(
            s,
            1.4,
            y + 0.08,
            11.2,
            0.8,
            [
                {"t": t["name"], "size": 16, "color": INK, "bold": True},
                {"t": t["slide"], "size": 14, "color": MUTED, "before": 1},
            ],
        )
    notes(s, "0:23–0:25. Коротко по пяти строкам. Обязательно скажите факт про лабораторию и больного: место не равно предмету труда. " + LAB_FACT + " " + TYPE_BORDERS)

    # 13 Section market
    s = _section(
        prs,
        "02",
        "Современный рынок труда и востребованные профессии",
        "«Надо» — это\nне приказ моды",
        "Спрос показывает, где людям нужны руки и головы. Он не ранжирует человеческое достоинство.",
        "НАДО",
        INDIGO,
    )
    dark.add(13)
    notes(s, "Короткий переход. Не начинайте с зарплат: в комплекте их нет специально.")

    # 14 Market
    s = new_slide(prs)
    eyebrow(s, "Три спокойных факта")
    title(s, "Карта спроса шире витрины", size=32)
    facts = [
        ("01", "Не только громкие профессии", "Медицина, образование, инженерия, рабочие специальности, социальная сфера, данные, экология и агро — это тоже спрос."),
        ("02", "Ценен стык", "Люди и данные. Техника и умение объяснить её человеку. Забота и знание правил."),
        ("03", "Должность меняется быстрее умений", "Учиться, проверять факты и договариваться остаются, даже если название работы сменится."),
    ]
    for i, (n, head, body) in enumerate(facts):
        y = 1.5 + i * 1.45
        add_shape(s, 0.5, y, 12.33, 1.32, WHITE, rounded=True, radius=0.08)
        write(s, 0.75, y, 1.1, 1.32, [{"t": n, "size": 22, "color": INDIGO, "bold": True}], anchor="middle")
        write(s, 2.0, y + 0.16, 10.4, 1.05, [
            {"t": head, "size": 20, "color": INK, "bold": True},
            {"t": body, "size": 15, "color": MUTED, "before": 2},
        ])
    notes(
        s,
        "0:26. Если спросят про нейросети: они меняют задачи внутри профессий, не отменяют ответственность за человека и умение поставить вопрос. "
        "Востребованность — информация, не рейтинг людей. Не называйте зарплаты.",
    )

    # 15 Myths Q
    s = new_slide(prs)
    eyebrow(s, "Голосование  ·  миф или факт")
    title(s, "Сначала зал. Потом ответы.")
    for i, (statement, _, _) in enumerate(MYTHS):
        y = 1.5 + i * 1.25
        add_shape(s, 0.5, y, 12.33, 1.1, WHITE, rounded=True, radius=0.1)
        write(s, 0.8, y, 1.0, 1.1, [{"t": f"0{i+1}", "size": 20, "color": "E15A3A", "bold": True}], anchor="middle")
        write(s, 1.9, y, 10.5, 1.1, [{"t": statement, "size": 22, "color": INK, "bold": True}], anchor="middle")
    notes(s, "По одной фразе. Руки: миф / факт. Не спорьте с залом до следующего слайда. Уложитесь в 3 минуты вместе с ответами.")

    # 16 Myths A
    s = new_slide(prs)
    eyebrow(s, "Разбор")
    title(s, "Три мифа и один рабочий факт", size=32)
    for i, (statement, kind, why) in enumerate(MYTHS):
        y = 1.48 + i * 1.28
        fill = "1F7A45" if kind == "Факт" else "E15A3A"
        add_shape(s, 0.5, y, 12.33, 1.16, WHITE, rounded=True, radius=0.08)
        add_shape(s, 0.5, y, 1.55, 1.16, fill)
        write(s, 0.5, y, 1.55, 1.16, [{"t": kind, "size": 16, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, 2.25, y + 0.1, 10.3, 0.98, [
            {"t": statement, "size": 16, "color": INK, "bold": True},
            {"t": why, "size": 14, "color": MUTED, "before": 2},
        ])
    notes(s, "Четвёртое — факт про профпробы. Это мост к траекториям. Не растягивайте.")

    # 17 Need write
    s = new_slide(prs)
    eyebrow(s, "Запись  ·  45 секунд", INDIGO)
    title(s, "Одна честная строка в «Надо»")
    add_shape(s, 0.5, 1.6, 12.33, 4.3, WHITE, rounded=True, radius=0.06)
    write(s, 0.9, 1.95, 11.5, 1.5, [{"t": "Кому и в чём может пригодиться то,\nчто вам самим не противно делать?", "size": 30, "color": INK, "bold": True, "spacing": 0.98}])
    write(s, 0.9, 3.7, 11.5, 1.6, [
        {"t": "Не лозунг «хочу приносить пользу». Можно про свой город.", "size": 20, "color": MUTED},
        {"t": "Если семья уже выбрала за вас — запишите это отдельной строкой как ожидание, и всё равно заполните «хочу».", "size": 20, "color": INK, "before": 10},
    ])
    notes(s, "0:31. 45 секунд тишины. Не собирайте строки. Фраза про семью — если видите, что зал про это, или скажите её превентивно одной фразой.")

    # 18 Section paths
    s = _section(
        prs,
        "03",
        "Образовательные траектории и навыки XXI века",
        "Путь длиннее,\nчем один экзамен",
        "Сначала проба и навык. Потом развилка. Потом конкретная дверь — если она ваша.",
        "МОГУ",
        TEAL,
    )
    dark.add(18)
    notes(s, "Переход 10 секунд.")

    # 19 Paths
    s = new_slide(prs)
    eyebrow(s, "Маршрут без клятвы")
    title(s, "Где вы сейчас — такая и задача", size=32)
    for i, (head, body) in enumerate(PATHS):
        x = 0.5 + i * 4.2
        add_shape(s, x, 1.6, 4.0, 4.55, WHITE, rounded=True, radius=0.08)
        add_shape(s, x, 1.6, 4.0, 0.18, [TEAL, AMBER_BG, INDIGO][i])
        write(s, x + 0.25, 2.0, 3.5, 1.1, [{"t": head, "size": 24, "color": INK, "bold": True}])
        write(s, x + 0.25, 3.2, 3.5, 2.5, [{"t": body, "size": 16, "color": MUTED, "spacing": 1.02}])
    notes(
        s,
        "0:33. Младшим: вуз называть не нужно. Девятому классу отдельно проговорите, что колледж — не утешительный приз. "
        "Старшим: решение с семьёй и по правилам этого года. «Профессионалитет» назовите.",
    )

    # 20 Skills
    s = new_slide(prs)
    eyebrow(s, "Навыки, которые переживают смену должности")
    title(s, "Семь опор почти любой работы", size=32)
    for i, (head, body) in enumerate(SKILLS):
        col = i % 4
        row = i // 4
        x = 0.5 + col * 3.2
        y = 1.55 + row * 2.5
        w = 3.05 if i < 4 else 4.0
        if i >= 4:
            x = 0.5 + (i - 4) * 4.2
            w = 4.0
        add_shape(s, x, y, w, 2.3, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.2, y + 0.25, w - 0.4, 0.7, [{"t": head, "size": 18, "color": NAVY, "bold": True}])
        write(s, x + 0.2, y + 1.05, w - 0.4, 1.0, [{"t": body, "size": 14, "color": MUTED}])
    notes(s, "Не читайте все семь медленно. Назовите список и задержитесь на этике: помощь и власть без границ опасны. Остальное — в тетради, галочки «хочу тренировать».")

    # 21 Can / skill
    s = new_slide(prs)
    eyebrow(s, "Задание  ·  90 секунд")
    title(s, "Могу — наблюдение. Умею — доказательство.")
    add_shape(s, 0.5, 1.55, 6.05, 5.05, WHITE, rounded=True, radius=0.08)
    add_shape(s, 0.5, 1.55, 6.05, 0.7, TEAL)
    write(s, 0.75, 1.55, 5.6, 0.7, [{"t": "Могу  ·  отметьте два", "size": 20, "color": WHITE, "bold": True}], anchor="middle")
    for i, opt in enumerate(CAN_OPTIONS):
        write(s, 0.85, 2.5 + i * 0.6, 5.4, 0.55, [{"t": opt, "size": 22, "color": INK}])
    add_shape(s, 6.75, 1.55, 6.08, 5.05, WHITE, rounded=True, radius=0.08)
    add_shape(s, 6.75, 1.55, 6.08, 0.7, AMBER_BG)
    write(s, 7.0, 1.55, 5.6, 0.7, [{"t": "Умею  ·  два доказательства", "size": 20, "color": NAVY_DEEP, "bold": True}], anchor="middle")
    write(
        s,
        7.05,
        2.55,
        5.5,
        3.6,
        [
            {"t": "«Я уже умею…»", "size": 22, "color": INK, "bold": True},
            {"t": "И ещё одно умение.", "size": 22, "color": INK, "bold": True, "before": 8},
            {"t": "«Умею лечить» в 7 классе обычно неправда. «Умею объяснить тему другу» — может быть правдой.", "size": 16, "color": MUTED, "before": 18},
        ],
    )
    notes(s, "0:37. Верните зал на страницу формулы. Полторы минуты. Не превращайте в конкурс достижений.")

    # 22 Programs
    s = new_slide(prs)
    eyebrow(s, "Высшее образование  ·  пример, не каталог")
    title(s, "Четыре двери, если близки люди и общество", size=28)
    write(s, 0.5, 1.2, 12.3, 0.38, [{"t": "Техника и природа ничуть не хуже. Эти направления — понятный путь для типов «человек» и «знаки».", "size": 15, "color": MUTED}])
    for i, p in enumerate(PROGRAMS):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 6.4
        y = 1.7 + row * 2.5
        add_shape(s, x, y, 6.2, 2.35, WHITE, rounded=True, radius=0.08)
        add_shape(s, x, y, 0.12, 2.35, [TEAL, "C24B35", INDIGO, "E15A3A"][i])
        write(s, x + 0.32, y + 0.14, 5.7, 0.55, [
            {"t": p["title"], "size": 20, "color": INK, "bold": True, "runs": [
                {"t": p["title"], "size": 20, "color": INK, "bold": True},
                {"t": "   " + p["code"], "size": 14, "color": MUTED, "bold": False},
            ]},
        ])
        write(s, x + 0.32, y + 0.75, 5.7, 1.4, [
            {"t": p["slide_body"], "size": 15, "color": INK},
            {"t": "Кому: " + p["slide_fit"], "size": 14, "color": MUTED, "before": 4},
            {"t": "Дальше: " + p["slide_jobs"], "size": 14, "color": MUTED, "before": 2},
        ])
    notes(
        s,
        "0:38. Одна фраза-рамка: это не список «выберите отсюда». По карточке — вопрос профессии, не лекция. "
        "Коды нужны старшим для поиска. Младшим достаточно сути. "
        + " ".join(f"{p['title']}: {p['question']}" for p in PROGRAMS),
    )

    # 23 Diff + admission
    s = new_slide(prs)
    eyebrow(s, "Чем они отличаются")
    title(s, "Один интерес — четыре разных вопроса", size=30)
    for i, p in enumerate(PROGRAMS):
        y = 1.4 + i * 0.85
        add_shape(s, 0.5, y, 12.33, 0.75, WHITE, rounded=True, radius=0.1)
        write(s, 0.7, y, 4.6, 0.75, [{"t": p["title"], "size": 16, "color": INK, "bold": True}], anchor="middle")
        write(s, 5.3, y, 7.2, 0.75, [{"t": p["question"], "size": 16, "color": INDIGO}], anchor="middle")
    add_shape(s, 0.5, 4.9, 12.33, 1.85, NAVY, rounded=True, radius=0.08)
    write(
        s,
        0.75,
        5.05,
        11.85,
        1.55,
        [
            {"t": "Поступление — ориентир, не шпаргалка на год вперёд", "size": 16, "color": "7DCEC4", "bold": True},
            {"t": "Бакалавриат, очно обычно 4 года. Предметы ЕГЭ назначает вуз. В тетради — типичный набор и прямая оговорка: правила приёма важнее этой презентации. Есть бюджет, плата, целевые места и путь через колледж.", "size": 15, "color": WHITE, "before": 6},
        ],
    )
    notes(s, "Не читайте перечень ЕГЭ с экрана: его нет на слайде специально, он в тетради, чтобы не врать точным набором. Скажите оговорку про социального педагога: это не то же самое, что специалист по социальной работе. Политолог — не автоматически политик. Организатор работы с молодёжью: возраст 14–35 по закону о молодёжной политике. " + ADMISSION_COMMON)

    # 24 Case intro
    s = new_slide(prs)
    eyebrow(s, "Кейс  ·  6 минут на команду")
    title(s, "Городская «Точка выбора»", size=34)
    write(s, 0.5, 1.35, 12.3, 1.15, [{"t": CASE["situation"], "size": 18, "color": INK, "spacing": 1.0}])
    write(s, 0.5, 2.6, 12.3, 0.45, [{"t": CASE["task"], "size": 20, "color": NAVY, "bold": True}])
    for i, crit in enumerate(CASE["criteria"]):
        y = 3.3 + i * 0.95
        add_shape(s, 0.5, y, 12.33, 0.82, WHITE, rounded=True, radius=0.1)
        write(s, 0.75, y, 1.1, 0.82, [{"t": f"0{i+1}", "size": 18, "color": "E15A3A", "bold": True}], anchor="middle")
        write(s, 1.9, y, 10.5, 0.82, [{"t": crit, "size": 20, "color": INK}], anchor="middle")
    notes(s, "0:40. Если зал — ряды, а не группы: каждый берёт одну роль и пишет сам, потом двое соседей складывают лучшее. План Б не прячьте.")

    # 25 Roles
    s = new_slide(prs)
    eyebrow(s, "Оставьте слайд на время работы")
    title(s, "Четыре роли, одно действие", size=32)
    for i, (role, code, question) in enumerate(CASE["roles"]):
        col = i % 2
        row = i // 2
        x = 0.5 + col * 6.4
        y = 1.5 + row * 2.5
        add_shape(s, x, y, 6.2, 2.32, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.28, y + 0.2, 5.65, 0.7, [
            {"t": role, "size": 18, "color": INK, "bold": True},
            {"t": code, "size": 13, "color": MUTED, "before": 1},
        ])
        write(s, x + 0.28, y + 1.15, 5.65, 0.9, [{"t": question, "size": 16, "color": INDIGO}])
    notes(
        s,
        "6 минут. Запись в тетради: название, для кого, какая клетка формулы, первый шаг. "
        "Потом 2–3 команды по 30 секунд. Вопрос: какую клетку закрыли и кого не заметят? "
        "Если застряли, подскажите направление, не готовый ответ: социологу нужно мнение многих; "
        "социальному работнику — безопасность разговора; политологу — проверка фактов; организатору — живое дело со сроком. "
        "Готовые примеры есть только в сценарии, на экран их не выводите.",
    )

    # 26 Section support
    s = _section(
        prs,
        "04",
        "Планирование карьеры и государственная поддержка",
        "От гипотезы\nк одному шагу",
        "Поддержка — не лозунг. Это конкретные двери. Подробности лежат в тетради и открываются со взрослым.",
        "НАДО",
        AMBER_BG,
    )
    dark.add(26)
    notes(s, "Переход. Скажите, что сейчас будет короткий план, а не каталог всех льгот страны.")

    # 27 Steps
    s = new_slide(prs)
    eyebrow(s, "Личный план")
    title(s, "Три шага, которые помещаются в месяц")
    steps = [
        ("1", "Гипотеза", "Ведущий тип и клетка «хочу». Это версия, её можно пересмотреть."),
        ("2", "Проба", "Одно живое дело на месяц. Список проб — в тетради. Не пять дел сразу."),
        ("3", "Маршрут", "Какие предметы держать в фокусе, с кем поговорить, какой официальный сервис открыть."),
    ]
    for i, (n, head, body) in enumerate(steps):
        x = 0.5 + i * 4.2
        add_shape(s, x, 1.7, 4.0, 4.4, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.25, 2.0, 3.5, 0.8, [{"t": n, "size": 40, "color": TEAL, "bold": True}])
        write(s, x + 0.25, 3.0, 3.5, 0.6, [{"t": head, "size": 26, "color": INK, "bold": True}])
        write(s, x + 0.25, 3.8, 3.5, 1.8, [{"t": body, "size": 16, "color": MUTED}])
    notes(s, "0:48. 90 секунд. Пробу не назначайте детям сами: они выбирают из списка в тетради.")

    # 28 Support
    s = new_slide(prs)
    eyebrow(s, "Государственная поддержка молодёжи")
    title(s, "Шесть дверей на экране. Ещё две — в тетради.", size=28)
    briefs = [
        "Диагностики, профпробы, экскурсии. Спросите педагога в школе.",
        "Курс для 6–11 классов. Узнайте, идёт ли он у вас и какие отрасли важны региону.",
        "Колледжи вместе с работодателями: больше практики и не одна квалификация.",
        "Договор и работа после выпуска. Читать только вместе с родителями.",
        "Заявление в вуз онлайн. Лимиты своего года проверьте заново, они в тетради.",
        "Профильные классы и кружки — способ проверить «умею» до поступления.",
    ]
    shown = SUPPORT[:6]
    for i, item in enumerate(shown):
        col = i % 3
        row = i // 3
        x = 0.5 + col * 4.2
        y = 1.5 + row * 2.55
        add_shape(s, x, y, 4.02, 2.4, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.2, y + 0.18, 3.62, 0.85, [{"t": item["title"], "size": 16, "color": NAVY, "bold": True, "spacing": 0.95}])
        write(s, x + 0.2, y + 1.1, 3.62, 1.1, [{"t": briefs[i], "size": 13, "color": MUTED, "spacing": 0.98}])
    notes(
        s,
        "0:50–0:53. По одной фразе на дверь, не читайте абзацы. Обязательно: целевое обучение читают с родителями; "
        "это обязательство, не подарок. «Билет в будущее» — bvbinfo.ru. Поступление — Госуслуги. "
        "Про «Движение Первых» и работу с 14 лет скажите, что они тоже в тетради: на слайде шесть самых ходовых. "
        "Не обещайте, что диагностика с портала автоматически зачислит в вуз.",
    )

    # 29 Letter
    s = new_slide(prs)
    eyebrow(s, "Письмо себе из будущего  ·  4 минуты")
    title(s, "Напишите от своего будущего «я»", size=32)
    lines = [
        "Мне ____ лет. Я учусь или работаю вот чем: …",
        "Мне по-настоящему нравится: …",
        "Из школьных умений пригодилось: …",
        "Одна фраза себе сегодняшнему: …",
    ]
    for i, line in enumerate(lines):
        y = 1.55 + i * 1.15
        add_shape(s, 0.5, y, 12.33, 1.02, WHITE, rounded=True, radius=0.1)
        write(s, 0.85, y, 11.7, 1.02, [{"t": line, "size": 22, "color": INK}], anchor="middle")
    notes(s, "0:53–0:58. Можно короче. Можно рисунком, если так честнее для 5–6 класса. Вслух — один доброволец и одна фраза, либо никто. Письмо не читаете вы. Полный бланк — в тетради.")

    # 30 Close
    s = new_slide(prs, dark=True)
    dark.add(30)
    write(s, 0.7, 0.45, 12, 0.3, [{"t": "СБОРКА", "size": 14, "color": "7DCEC4", "bold": True}])
    write(s, 0.7, 0.85, 12, 0.7, [{"t": "Соберите формулу одной строкой", "size": 30, "color": WHITE, "bold": True}])
    stems = ["Мне интересно…", "У меня получается…", "Я уже умею…", "Другим людям нужно…"]
    for i, stem in enumerate(stems):
        x = 0.7 + (i % 4) * 3.1
        add_shape(s, x, 1.8, 2.95, 1.15, "1B3E58", rounded=True, radius=0.1)
        write(s, x + 0.12, 1.8, 2.7, 1.15, [{"t": stem, "size": 16, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
    add_shape(s, 0.7, 3.2, 11.9, 1.15, TEAL, rounded=True, radius=0.08)
    write(s, 0.95, 3.2, 11.4, 1.15, [{"t": "Мой шаг на этот месяц: …", "size": 26, "color": WHITE, "bold": True}], anchor="middle")
    write(s, 0.7, 4.6, 12, 1.5, [
        {"t": "Вопрос домой", "size": 14, "color": "F3C15B", "bold": True},
        {"t": "Какую пробу я сделаю в этом месяце?", "size": 24, "color": WHITE, "bold": True, "before": 4},
        {"t": "Тетрадь заберите. Тест за вас не решает. Спасибо.", "size": 16, "color": "C5D0D6", "before": 8},
    ])
    notes(s, "0:58–1:00. Шёпотом соседу или строкой в тетради. Поблагодарите. Не добавляйте новую тему. Если осталось 30 секунд — только вопрос домой.")

    return _finish(prs, dark)


def build_short():
    prs = _prs()
    prs.core_properties.title = "Путь в профессию · 20 минут"
    dark = {1, 4, 12}

    s = new_slide(prs, dark=True)
    write(s, 0.7, 1.05, 12, 0.3, [{"t": "ФЕСТИВАЛЬ  «НАУКА 0+»", "size": 15, "color": "7DCEC4", "bold": True}])
    write(s, 0.7, 1.5, 12, 1.2, [{"t": "Путь в профессию", "size": 54, "color": WHITE, "bold": True}])
    write(s, 0.7, 2.9, 11, 0.8, [{"t": "Экспресс-встреча: гипотеза, формула и один шаг.\nНе притворяемся, что за 20 минут выбирают судьбу.", "size": 20, "color": "D5E4EA"}])
    chip(s, 0.7, 4.15, 2.35, 0.48, "20 МИНУТ", AMBER_BG, NAVY_DEEP, size=15)
    chip(s, 3.2, 4.15, 2.5, 0.48, "5–11 КЛАСС", "245066", WHITE, size=15)
    _formula_pills(s, 0.7, 5.15)
    notes(s, "0:00. Сразу рамка: профессию за 20 минут не выбираем. Тетрадь унесёт глубину. Формула — хочу, могу, умею, надо. Тест — гипотеза.")

    s = new_slide(prs)
    eyebrow(s, "Вся встреча держится на этом")
    title(s, "Формула и одно правило", size=32)
    _formula_row(s, 1.45, 2.55)
    add_shape(s, 0.5, 4.25, 12.33, 2.4, NAVY, rounded=True, radius=0.08)
    write(
        s,
        0.8,
        4.45,
        11.75,
        2.05,
        [
            {"t": "Правило", "size": 14, "color": "7DCEC4", "bold": True},
            {"t": "Опросник — гипотеза, не приговор и не оценка. Баллы вслух не собираем. Тетрадь не сдают.", "size": 20, "color": WHITE, "before": 6},
            {"t": FORMULA_TRAPS, "size": 16, "color": "E4EEF2", "before": 8},
        ],
    )
    notes(s, "0:01. Не расшифровывайте каждое слово долго: определения уже на карточках. Правило проговорите целиком.")

    s = new_slide(prs)
    eyebrow(s, "ДДО Е. А. Климова  ·  экспресс")
    title(s, "Только вопросы 1–10, помета «экспресс»")
    bits = [
        ("5 минут", "Первый отклик. В паре один вариант."),
        ("Сумма 10", "Если ответили на все. Иначе ищите пропуск."),
        ("Максимум 4", "Разница в один балл — шум, не характер."),
        ("Ключ потом", "Он не подсказывает «правильный» выбор."),
    ]
    for i, (head, body) in enumerate(bits):
        x = 0.5 + (i % 4) * 3.2
        add_shape(s, x, 1.6, 3.05, 2.7, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.18, 1.85, 2.7, 0.9, [{"t": head, "size": 22, "color": NAVY, "bold": True}])
        write(s, x + 0.18, 2.9, 2.7, 1.1, [{"t": body, "size": 16, "color": MUTED}])
    add_shape(s, 0.5, 4.55, 12.33, 2.1, "E7F3F1", rounded=True, radius=0.08)
    write(
        s,
        0.8,
        4.8,
        11.8,
        1.65,
        [
            {"t": "Дома можно дорешать вопросы 11–15 и пересчитать по полному ключу: максимум станет 6, сумма — 15.", "size": 18, "color": INK},
            {"t": "Представьте, что научиться можно всему. Непонятное слово — тихая рука.", "size": 18, "color": INK, "before": 8},
        ],
    )
    notes(s, "0:01:30. Покажите в тетради помету «экспресс» у вопросов 1–10. Попросите не открывать ключ. Полную версию не делайте в зале: не успеете.")

    s = new_slide(prs, dark=True)
    write(s, 0.7, 1.45, 12, 0.35, [{"t": "ВОПРОСЫ 1–10", "size": 16, "color": "7DCEC4", "bold": True}])
    write(s, 0.7, 1.95, 12, 1.5, [{"t": "5:00", "size": 96, "color": WHITE, "bold": True}])
    write(s, 0.7, 4.0, 11, 1.2, [{"t": "Тишина. А или Б.\nБез разговоров и без ключа.", "size": 28, "color": "E4EEF2"}])
    notes(s, "0:02–0:07. Молчите. На четвёртой минуте: «где сомневаетесь — ставьте первый отклик».")

    s = new_slide(prs)
    eyebrow(s, "Экспресс-ключ  ·  только 1–10")
    title(s, "Сложите совпадения", size=32)
    write(s, 0.5, 1.22, 12.3, 0.35, [{"t": "Максимум в столбце — 4. Сумма — 10. Вопросы 11–15 сейчас не считаем.", "size": 16, "color": MUTED}])
    _key_columns(s, express=True, y=1.7)
    add_shape(s, 0.5, 5.35, 12.33, 1.35, NAVY, rounded=True, radius=0.1)
    write(
        s,
        0.75,
        5.5,
        11.85,
        1.05,
        [
            {"t": "Разница в 1 балл — шум. Ноль — не запрет. Сумма должна быть 10.", "size": 18, "color": WHITE, "bold": True},
            {"t": "Вопросы 11–15 можно досчитать дома: максимум станет 6.", "size": 16, "color": "D5E4EA", "before": 4},
        ],
        anchor="middle",
    )
    notes(s, "0:07. Дайте минуту. Не просите называть баллы. Напомните: верхний столбец — нынешний интерес. Буква Б заглавная специально.")

    s = new_slide(prs)
    eyebrow(s, "Полминуты на чтение")
    title(s, "Ваш верхний столбец — гипотеза", size=30)
    for i, t in enumerate(TYPES):
        y = 1.4 + i * 0.95
        add_shape(s, 0.5, y, 12.33, 0.86, WHITE, rounded=True, radius=0.08)
        add_shape(s, 0.5, y, 0.62, 0.86, TYPE_COLORS[t["code"]])
        write(s, 0.5, y, 0.62, 0.86, [{"t": t["code"], "size": 16, "color": WHITE, "bold": True, "align": "center"}], anchor="middle")
        write(s, 1.3, y, 7.2, 0.86, [{"t": t["name"], "size": 16, "color": INK, "bold": True}], anchor="middle")
        write(s, 8.3, y, 4.3, 0.86, [{"t": t["slide"], "size": 13, "color": MUTED}], anchor="middle")
    notes(s, "0:08–0:09. Одна фраза: ноль — не запрет, ровный профиль — не поломка. Портреты типов есть в тетради. Факт про лабораторию скажите только если останется 20 секунд.")

    s = new_slide(prs)
    eyebrow(s, "Надо и маршрут  ·  без каталога зарплат")
    title(s, "Спрос шире витрины", size=32)
    cols = [
        ("Надо", "Нужны забота, инженеры, рабочие профессии, данные, земля — не только модные названия.", INDIGO),
        ("Сейчас", "5–8 класс: пробы. 9 класс: школа или колледж. 10–11: маршрут поступления.", TEAL),
        ("Строка", "Запишите, кому может пригодиться дело, которое вам не противно.", "E15A3A"),
    ]
    for i, (head, body, color) in enumerate(cols):
        x = 0.5 + i * 4.2
        add_shape(s, x, 1.6, 4.0, 4.5, WHITE, rounded=True, radius=0.08)
        add_shape(s, x, 1.6, 4.0, 0.16, color)
        write(s, x + 0.25, 2.05, 3.5, 0.7, [{"t": head, "size": 26, "color": color, "bold": True}])
        write(s, x + 0.25, 3.0, 3.5, 2.6, [{"t": body, "size": 18, "color": INK, "spacing": 1.02}])
    notes(s, "0:09–0:11. Три колонки вслух по предложению. Не открывайте мифы: на них нет времени. Нейросети — только если спросят, одной фразой из часового сценария.")

    s = new_slide(prs)
    eyebrow(s, "Пример пути, не весь мир профессий")
    title(s, "Четыре направления высшего образования", size=28)
    for i, p in enumerate(PROGRAMS):
        y = 1.4 + i * 1.2
        add_shape(s, 0.5, y, 12.33, 1.08, WHITE, rounded=True, radius=0.08)
        write(s, 0.75, y, 4.5, 1.08, [
            {"t": p["title"], "size": 18, "color": INK, "bold": True},
            {"t": p["code"], "size": 13, "color": MUTED, "before": 0},
        ], anchor="middle")
        write(s, 5.3, y, 7.2, 1.08, [{"t": p["question"], "size": 18, "color": INDIGO}], anchor="middle")
    notes(
        s,
        "0:11–0:12. Прочитайте четыре вопроса. Добавьте: бакалавриат обычно 4 года, предметы ЕГЭ — в тетради и в правилах вуза. "
        "Социальный педагог не равен специалисту по социальной работе. Политолог не равен политику. "
        "Если тип зала — техника или природа, скажите прямо: ваши двери другие и они не хуже.",
    )

    s = new_slide(prs)
    eyebrow(s, "60 секунд в тетради")
    title(s, "Два «могу». Одно «умею».")
    add_shape(s, 0.5, 1.55, 7.5, 5.0, WHITE, rounded=True, radius=0.08)
    write(s, 0.8, 1.8, 7.0, 0.5, [{"t": "Отметьте два", "size": 20, "color": TEAL, "bold": True}])
    write(s, 0.8, 2.5, 7.0, 3.5, [{"t": "  ·  ".join(CAN_OPTIONS), "size": 22, "color": INK, "spacing": 1.15}])
    add_shape(s, 8.2, 1.55, 4.63, 5.0, NAVY, rounded=True, radius=0.08)
    write(
        s,
        8.45,
        1.9,
        4.2,
        4.3,
        [
            {"t": "Умею", "size": 20, "color": "F3C15B", "bold": True},
            {"t": "Одно умение, которое можно доказать.", "size": 20, "color": WHITE, "before": 12},
            {"t": "Не мечта. Факт: получалось и это кто-то видел.", "size": 16, "color": "D5E4EA", "before": 14},
        ],
    )
    notes(s, "0:12–0:13:30. Тишина. Не проверяйте вслух.")

    s = new_slide(prs)
    eyebrow(s, "Микрокейс  ·  2 минуты")
    title(s, "«Мне всё равно, кем быть»", size=32)
    write(s, 0.5, 1.35, 12.3, 0.8, [{"t": "Так говорит одноклассник. Выберите одну роль и напишите одно действие. Не лекцию.", "size": 20, "color": INK}])
    for i, (role, code, question) in enumerate(CASE["roles"]):
        x = 0.5 + i * 3.2
        add_shape(s, x, 2.4, 3.05, 3.55, WHITE, rounded=True, radius=0.08)
        write(s, x + 0.16, 2.6, 2.75, 1.3, [{"t": role, "size": 16, "color": INK, "bold": True, "spacing": 0.95}])
        write(s, x + 0.16, 4.0, 2.75, 1.6, [{"t": question, "size": 14, "color": INDIGO, "spacing": 0.98}])
    notes(s, "0:13:30–0:16. Пишут все. Вслух — три человека по одному предложению, только добровольцы. Если тишина, зачитайте свой запасной ход из сценария и идите дальше. Не устраивайте защиту проектов.")

    s = new_slide(prs)
    eyebrow(s, "Письмо  ·  первые строки, остальное дома")
    title(s, "Себе из будущего — три начала", size=32)
    for i, line in enumerate(["Мне ____ лет, и я занимаюсь…", "Из школьного у меня осталось…", "Тебе сегодняшнему я скажу…"]):
        y = 1.6 + i * 1.45
        add_shape(s, 0.5, y, 12.33, 1.28, WHITE, rounded=True, radius=0.1)
        write(s, 0.85, y, 11.7, 1.28, [{"t": line, "size": 26, "color": INK, "bold": True}], anchor="middle")
    notes(s, "0:16–0:18. Две минуты. Письмо не собираете. Полный бланк в тетради. Для 5 класса достаточно одной фразы или рисунка.")

    s = new_slide(prs, dark=True)
    write(s, 0.65, 0.35, 12, 0.3, [{"t": "ЗАБРАТЬ С СОБОЙ", "size": 14, "color": "7DCEC4", "bold": True}])
    write(s, 0.65, 0.7, 12, 0.6, [{"t": "Четыре опоры и одна строка", "size": 30, "color": WHITE, "bold": True}])
    doors = [
        ("Билет в будущее", "Диагностики и профпробы. Спросите педагога."),
        ("Профессионалитет", "Колледж вместе с работодателем."),
        ("Целевое обучение", "Договор. Только вместе с родителями."),
        ("Госуслуги и Первые", "Поступление онлайн. Проекты — в Движении Первых."),
    ]
    for i, (head, body) in enumerate(doors):
        x = 0.65 + (i % 4) * 3.15
        add_shape(s, x, 1.55, 3.0, 2.15, "1B3E58", rounded=True, radius=0.08)
        write(s, x + 0.14, 1.7, 2.72, 1.85, [
            {"t": head, "size": 15, "color": "F3C15B", "bold": True},
            {"t": body, "size": 13, "color": WHITE, "before": 6},
        ])
    write(s, 0.65, 3.95, 12, 1.3, [
        {"t": "Мне интересно…  Я могу…  Я уже умею…  Другим нужно…", "size": 20, "color": WHITE, "bold": True},
        {"t": "Шаг на месяц запишите в тетради.", "size": 18, "color": "D5E4EA", "before": 6},
    ])
    write(s, 0.65, 5.5, 12, 1.2, [
        {"t": "Вопрос домой: какую пробу я сделаю в этом месяце?", "size": 22, "color": WHITE, "bold": True},
        {"t": "Остальное — в тетради. Спасибо.", "size": 16, "color": "C5D0D6", "before": 6},
    ])
    notes(
        s,
        "0:18–0:20. Четыре имени поддержки произнесите. Целевое — с оговоркой про родителей. "
        "Не начинайте новую лекцию. Вопрос домой — последняя фраза.",
    )

    return _finish(prs, dark)


def _formula_pills(slide, x, y):
    colors = ["E15A3A", TEAL, AMBER_BG, INDIGO]
    names = ["Хочу", "Могу", "Умею", "Надо"]
    fg = [WHITE, WHITE, NAVY_DEEP, WHITE]
    for i, name in enumerate(names):
        chip(slide, x + i * 1.7, y, 1.55, 0.48, name, colors[i], fg[i], size=16)


def main():
    hour = build_hour()
    short = build_short()
    hour.save("/workspace/materials/Путь_в_профессию_презентация_60_минут.pptx")
    short.save("/workspace/materials/Путь_в_профессию_презентация_20_минут.pptx")
    print(f"hour slides: {len(hour.slides)}")
    print(f"short slides: {len(short.slides)}")


if __name__ == "__main__":
    main()
