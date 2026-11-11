"""Рабочая тетрадь участника. Одна на обе версии встречи."""

from __future__ import annotations

from content import (
    ADMISSION_COMMON,
    CAN_OPTIONS,
    CASE,
    FAMILY_QUESTIONS,
    FORMULA,
    FORMULA_TRAPS,
    LAB_FACT,
    LETTER_PROMPTS,
    MONTH_TRIALS,
    NOT_MEASURED,
    PATHS,
    PRO_QUESTIONS,
    PROFESSION_FORMULA,
    PROGRAMS,
    SKILLS,
    SOURCES,
    SUPPORT,
    TYPE_BORDERS,
    TYPES,
    VERSION,
)
from ddo import QUESTIONS, TYPE_ORDER, TYPE_SHORT, answers_by_type, expected_sum, max_score
from docx_kit import (
    INK,
    MUTED,
    NAVY,
    TEAL,
    add_text,
    _bottom_border,
    banner,
    bullet,
    callout,
    cell_text,
    content_width_cm,
    h1,
    h2,
    lines,
    new_doc,
    p,
    page_break,
    prevent_row_split,
    rich,
    set_cell_margins,
    set_row_height,
    set_table_borders,
    set_table_widths,
    set_valign,
    shade_cell,
    spacer,
    spacing,
)

W = content_width_cm()
FORMULA_COLORS = {
    "Хочу": ("FDE8E2", "E15A3A"),
    "Могу": ("E5F4F1", "0F7A6C"),
    "Умею": ("FFF4D8", "A86B12"),
    "Надо": ("E8EAF8", "3A46A0"),
}
TYPE_FILL = {"П": "1F7A45", "Т": "3E5870", "Ч": "C24B35", "З": "2A6FBA", "Х": "7A4C96"}


def build():
    doc = new_doc(f"Наука 0+   ·   Путь в профессию   ·   тетрадь участника   ·   комплект {VERSION}")
    doc.core_properties.title = "Путь в профессию. Рабочая тетрадь участника"
    doc.core_properties.subject = f"Фестиваль «Наука 0+». Комплект {VERSION}"
    _cover(doc)
    page_break(doc)
    _case(doc)
    page_break(doc)
    _formula(doc)
    page_break(doc)
    _profession_formula_page(doc)
    page_break(doc)
    _questions(doc)
    page_break(doc)
    _score_sheet(doc)
    page_break(doc)
    _key(doc)
    page_break(doc)
    _types(doc)
    page_break(doc)
    _route(doc)
    page_break(doc)
    _programs(doc)
    page_break(doc)
    _support(doc)
    _letter(doc)
    page_break(doc)
    _home(doc)
    path = "/workspace/materials/Путь_в_профессию_раздаточный_материал.docx"
    doc.save(path)
    return path


def _cover(doc):
    banner(doc, "ФЕСТИВАЛЬ  «НАУКА 0+»", NAVY, "7DCEC4", 12)
    p(doc, "Рабочая тетрадь участника", size=14, color=TEAL, bold=True, before=8, after=2)
    p(doc, "Путь в профессию", size=32, color=NAVY, bold=True, before=0, after=4, line=1.0)
    p(doc, "Формула  Хочу  ·  Могу  ·  Умею  ·  Надо", size=14, color=INK, before=0, after=8)
    p(
        doc,
        "Для 5–11 классов. Встреча на 60 минут, экспресс на 20 минут или работа дома. "
        "Тетрадь не сдают и не собирают. Имя писать не обязательно.",
        size=11,
        after=8,
    )

    table = doc.add_table(rows=1, cols=4)
    set_table_widths(table, [W / 4] * 4)
    set_table_borders(table, "FFFFFF", "8", "single")
    for i, (name, _, text) in enumerate(FORMULA):
        cell = table.cell(0, i)
        fill, bar = FORMULA_COLORS[name]
        shade_cell(cell, fill)
        set_cell_margins(cell, 80, 80, 80, 80)
        set_valign(cell, "top")
        cell_text(
            cell,
            [
                {"t": name, "size": 13, "color": bar, "bold": True, "after": 2},
                {"t": text, "size": 9, "color": INK, "after": 0},
            ],
        )
    prevent_row_split(table.rows[0])
    spacer(doc, 8)

    callout(
        doc,
        "Как пользоваться",
        "Порядок: кейс про Леру → сбор фактов → «хочу — могу — умею — надо» → формула профессии по Климову "
        "(это не тест!) → тест ДДО → направления подготовки → решение кейса → письмо. "
        "На 60 минут — вопросы 1–15; на 20 минут — только «экспресс» 1–10.",
        "F7F4EE",
        NAVY,
    )
    callout(
        doc,
        "Это не приговор",
        "Сокращённый опросник Е. А. Климова показывает, с каким предметом труда вам сейчас интереснее иметь дело. "
        "Он не назначает профессию, не измеряет талант, доброту и здоровье и ничего не говорит о том, чего хочет страна. "
        "Результат — гипотеза. Её проверяют делом, а не спором о баллах.",
        "FDE8E2",
        "E15A3A",
    )

    meta = doc.add_table(rows=2, cols=2)
    set_table_widths(meta, [W / 2, W / 2])
    set_table_borders(meta, "D9D0C4", "4")
    labels = [("Дата", ""), ("Класс", ""), ("Город", ""), ("Имя — по желанию", "")]
    for i, (label, _) in enumerate(labels):
        cell = meta.cell(i // 2, i % 2)
        shade_cell(cell, "FFFFFF")
        set_cell_margins(cell, 80, 80, 100, 100)
        cell_text(cell, [{"t": label, "size": 10, "color": MUTED, "after": 0}, {"t": " ", "size": 14, "after": 0}])
    spacer(doc, 8)
    p(doc, f"Комплект {VERSION}. Печать: А4, можно чёрно-белую, двустороннюю, скрепить скобой.", size=9, color=MUTED, after=8)
    callout(
        doc,
        "Что унесёте с собой",
        "Формулу выбора, понимание формулы профессии (это не тест!), гипотезу по ДДО, "
        "ориентир по четырём направлениям, заполненную карточку для Леры, шаг на месяц и начатое письмо.",
        "E7F3F1",
        TEAL,
    )


def _formula(doc):
    h1(doc, "Формула «Хочу — Могу — Умею — Надо»")
    p(doc, FORMULA_TRAPS, size=11, after=3)
    p(doc, "Это формула выбора человека — второй интерактив. Заполните клетки себе и отдельно для Леры. Не путайте с формулой профессии и с тестом.", size=11, after=6)
    table = doc.add_table(rows=2, cols=2)
    set_table_widths(table, [W / 2, W / 2])
    set_table_borders(table, "FFFFFF", "12", "single")
    prompts = {
        "Хочу": "Какие занятия тянут, даже когда за них не ставят отметку? Не должность, а дело.\n\nВедущий тип по опросу: __________\nВторой тип: __________",
        "Могу": "К чему есть способность — не только отметка? Отметьте два или напишите своё.\n" + "\n".join(f"☐  {item}" for item in CAN_OPTIONS),
        "Умею": "Что уже получается так, что это можно показать? Вас просили, у вас выходило, дело довели до конца.\n\n1. ________________________________\n\n2. ________________________________",
        "Надо": "Кому и в чём может пригодиться то, что вам самим не противно? Если семья уже выбрала за вас, запишите это отдельно как ожидание, а не как своё «хочу».",
    }
    for i, (name, _, _) in enumerate(FORMULA):
        cell = table.cell(i // 2, i % 2)
        fill, bar = FORMULA_COLORS[name]
        shade_cell(cell, fill)
        set_cell_margins(cell, 90, 100, 110, 110)
        set_valign(cell, "top")
        set_row_height(table.rows[i // 2], 6.3)
        blocks = [{"t": name.upper(), "size": 14, "color": bar, "bold": True, "after": 3}]
        for chunk in prompts[name].split("\n"):
            blocks.append({"t": chunk if chunk else " ", "size": 10, "color": INK, "after": 1, "line": 1.02})
        blocks.append({"t": " ", "size": 12, "after": 0})
        cell_text(cell, blocks)
    for row in table.rows:
        prevent_row_split(row)
    spacer(doc, 6)
    h2(doc, "Сборка — одна строка про себя")
    for stem in (
        "Мне интересно",
        "У меня получается",
        "Я уже умею",
        "Другим людям нужно",
        "Мой шаг на этот месяц",
    ):
        lines(doc, stem, 1, size=11)
    h2(doc, "Отдельно — формула Леры")
    for stem in (
        "Лера хочет",
        "Лере даётся",
        "Лера уже умеет",
        "Лере важно, чтобы польза была",
    ):
        lines(doc, stem, 1, size=11)


def _questions(doc):
    h1(doc, "Тест ДДО — отдельно от формулы профессии")
    p(doc, "Дифференциально-диагностический опросник Е. А. Климова. Учебный сокращённый вариант.", size=11, color=TEAL, bold=True, after=3)
    p(
        doc,
        "Это не формула профессии и не «хочу — могу — умею — надо». Тест показывает только, к какому предмету труда вас сейчас тянет. "
        "Представьте, что после обучения вы сможете выполнять любую работу. Если бы пришлось выбрать только одно из двух занятий — какое ближе? "
        "В каждой паре отметьте один вариант, А или Б. Долго не взвешивайте: нужен первый отклик. Правильных ответов нет.",
        size=11,
        after=3,
    )
    p(
        doc,
        "Вопросы 1–10 помечены словом «экспресс»: их достаточно для встречи на 20 минут. На встрече в 60 минут отвечайте на все 15. Ключ не открывайте, пока ведущий не скажет «считаем» — или пока не закончите, если проходите дома.",
        size=11,
        after=4,
    )
    callout(
        doc,
        "Про пункт 10Б",
        "Это учебная адаптация классической формулировки про набор текстов на клавишных машинах. Предмет труда тот же: знаки, а не природа и не люди.",
        "FFF4D8",
        "A86B12",
    )
    for q in QUESTIONS:
        _question_block(doc, q)


def _question_block(doc, q):
    table = doc.add_table(rows=1, cols=1)
    set_table_widths(table, [W])
    set_table_borders(table, "E4DCD0", "4")
    cell = table.cell(0, 0)
    shade_cell(cell, "FFFFFF")
    set_cell_margins(cell, 60, 70, 100, 100)
    mark = "   ·   экспресс" if q["express"] else ""
    cell_text(
        cell,
        [
            {"t": f"{q['n']}{mark}", "size": 11, "color": TEAL if q["express"] else NAVY, "bold": True, "after": 2},
            {"t": "А.  " + q["a"], "size": 11, "color": INK, "after": 1},
            {"t": "Б.  " + q["b"], "size": 11, "color": INK, "after": 2},
            {"t": "Мой выбор:    ☐ А      ☐ Б", "size": 11, "color": NAVY, "bold": True, "after": 0},
        ],
    )
    prevent_row_split(table.rows[0])
    spacer(doc, 3)


def _score_sheet(doc):
    h1(doc, "Лист подсчёта")
    p(doc, "Сначала выпишите свой выбор. Баллы по типам ставьте только с ключом следующей страницы.", size=11, after=4)
    table = doc.add_table(rows=1 + len(QUESTIONS), cols=3)
    set_table_widths(table, [2.0, 3.2, W - 5.2])
    set_table_borders(table, "D9D0C4", "4")
    headers = ["№", "Помета", "Мой выбор: А или Б"]
    for i, text in enumerate(headers):
        cell = table.cell(0, i)
        shade_cell(cell, NAVY)
        set_cell_margins(cell, 50, 50, 70, 70)
        cell_text(cell, [{"t": text, "size": 10, "color": "FFFFFF", "bold": True, "after": 0}])
    for idx, q in enumerate(QUESTIONS, start=1):
        values = [str(q["n"]), "экспресс" if q["express"] else "полная версия", ""]
        for col, text in enumerate(values):
            cell = table.cell(idx, col)
            shade_cell(cell, "F7F4EE" if idx % 2 == 0 else "FFFFFF")
            set_cell_margins(cell, 40, 40, 70, 70)
            set_row_height(table.rows[idx], 0.72)
            cell_text(cell, [{"t": text, "size": 10, "color": INK, "bold": col == 0, "after": 0}])
    spacer(doc, 8)
    h2(doc, "Баллы по типам")
    p(doc, "Когда ключ открыт, посчитайте совпадения. Запишите число в каждый столбец.", size=11, after=4)
    scores = doc.add_table(rows=2, cols=5)
    set_table_widths(scores, [W / 5] * 5)
    set_table_borders(scores, "FFFFFF", "8", "single")
    for i, code in enumerate(TYPE_ORDER):
        top = scores.cell(0, i)
        bottom = scores.cell(1, i)
        shade_cell(top, TYPE_FILL[code])
        shade_cell(bottom, "F7F4EE")
        set_cell_margins(top, 60, 60, 60, 60)
        set_cell_margins(bottom, 80, 80, 60, 60)
        set_valign(top, "center")
        cell_text(top, [{"t": f"{code}  ·  {TYPE_SHORT[code]}", "size": 10, "color": "FFFFFF", "bold": True, "align": "center", "after": 0}])
        cell_text(bottom, [{"t": "баллов: ______", "size": 11, "color": INK, "align": "center", "after": 0}])
    spacer(doc, 6)
    callout(
        doc,
        "Самопроверка",
        f"Полная версия, вопросы 1–15: сумма пяти чисел равна {expected_sum(False)}, максимум в одном столбце — {max_score(False)}. "
        f"Экспресс, только 1–10: сумма равна {expected_sum(True)}, максимум — {max_score(True)}. Если не сошлось, пропущен вопрос или балл попал не в тот столбец.",
    )


def _key(doc):
    banner(doc, "СТОП.  Эту страницу открывают со словом «считаем»", "E15A3A", "FFFFFF", 14)
    h1(doc, "Ключ")
    p(doc, "За каждое совпадение — один балл в столбец типа. «3 Б» значит: в вопросе 3 вы выбрали вариант Б. Буква Б написана заглавной, чтобы её не перепутать с цифрой.", size=11, after=4)
    _key_table(doc, "Полная версия, вопросы 1–15", False)
    spacer(doc, 4)
    _key_table(doc, "Экспресс, только вопросы 1–10", True)
    h2(doc, "Учебная шкала этого комплекта")
    p(doc, "Это не нормы полного опросника Климова, где максимум был 8. Шкала пересчитана под наш сокращённый вариант. Разница в один балл почти ничего не значит.", size=11, after=3)
    bullet(doc, "Полная версия. 0–1 — интерес почти не проявился. 2–3 — есть, но не главный. 4–5 — выраженный. 6 — очень яркий.")
    bullet(doc, "Экспресс. 0–1 — слабо. 2 — заметно. 3–4 — сильно. Сетка грубее: дома лучше дорешать вопросы 11–15.")
    bullet(doc, "Смотрите на один-два верхних столбца. Ничья — честный результат, запишите оба типа.")
    bullet(doc, "Ноль не значит «вам нельзя». Ровный профиль — повод пробовать разное, а не признак поломки.")
    h2(doc, "Чего опросник не измеряет")
    for item in NOT_MEASURED:
        bullet(doc, item)
    p(doc, LAB_FACT, size=11, italic=True, before=4, after=2)
    p(doc, "Запишите ведущий и второй тип в клетку «Хочу» на странице формулы. Вслух баллы называть не нужно.", size=11, before=2)


def _key_table(doc, title, express):
    p(doc, title, size=12, color=NAVY, bold=True, before=2, after=3)
    buckets = answers_by_type(express)
    table = doc.add_table(rows=2, cols=5)
    set_table_widths(table, [W / 5] * 5)
    set_table_borders(table, "FFFFFF", "8", "single")
    for i, code in enumerate(TYPE_ORDER):
        top = table.cell(0, i)
        bottom = table.cell(1, i)
        shade_cell(top, TYPE_FILL[code])
        shade_cell(bottom, "F7F4EE")
        set_cell_margins(top, 50, 50, 50, 50)
        set_cell_margins(bottom, 70, 70, 60, 60)
        cell_text(top, [{"t": f"{code}  {TYPE_SHORT[code]}", "size": 9, "color": "FFFFFF", "bold": True, "align": "center", "after": 0}])
        cell_text(
            bottom,
            [{"t": item, "size": 11, "color": INK, "bold": True, "align": "center", "after": 0} for item in buckets[code]],
        )
    prevent_row_split(table.rows[0])
    prevent_row_split(table.rows[1])


def _types(doc):
    h1(doc, "Пять типов профессий")
    p(
        doc,
        "Это часть формулы профессии (предмет труда), а не результат теста. "
        "Климов делит профессии по предмету труда: с чем вы работаете. Не по зарплате и не по тому, как профессия звучит. "
        "Для Леры чаще звучит тип «человек», возможен стык со «знаками».",
        size=11,
        after=4,
    )
    for t in TYPES:
        table = doc.add_table(rows=1, cols=2)
        set_table_widths(table, [1.3, W - 1.3])
        set_table_borders(table, "FFFFFF", "8", "single")
        mark, body = table.rows[0].cells
        shade_cell(mark, TYPE_FILL[t["code"]])
        shade_cell(body, "F7F4EE")
        set_valign(mark, "center")
        set_cell_margins(mark, 60, 60, 40, 40)
        set_cell_margins(body, 70, 70, 100, 100)
        cell_text(mark, [{"t": t["code"], "size": 16, "color": "FFFFFF", "bold": True, "align": "center", "after": 0}])
        cell_text(
            body,
            [
                {"t": t["name"], "size": 12, "color": NAVY, "bold": True, "after": 1},
                {"t": "Предмет. " + t["object"], "size": 10.5, "after": 1},
                {"t": "Ближе тем, кому нравится " + t["likes"][0].lower() + t["likes"][1:], "size": 10.5, "after": 1},
                {"t": "Примеры. " + t["examples"], "size": 10.5, "after": 1},
                {"t": t["note"], "size": 10.5, "italic": True, "color": MUTED, "after": 0},
            ],
        )
        prevent_row_split(table.rows[0])
    spacer(doc, 4)
    callout(doc, "Границы", TYPE_BORDERS, "E8EAF8", "3A46A0")
    p(
        doc,
        "Социальная работа и организация работы с молодёжью ближе к типу «человек». Социология и политология стоят на стыке «человека» и «знаков»: там изучают людей и решения через данные, тексты и правила.",
        size=11,
        before=2,
    )


def _route(doc):
    h1(doc, "Рынок, маршрут, навыки")
    h2(doc, "Что полезно знать про «надо»")
    bullet(doc, "Стране нужны не только профессии, о которых громко говорят. Медицина, образование, инженерия, рабочие специальности, социальная сфера, данные, экология и агро — это тоже спрос.")
    bullet(doc, "Всё ценнее стык: люди и данные, техника и умение её объяснить, забота и знание правил.")
    bullet(doc, "Название должности меняется чаще, чем умение учиться, проверять факты и договариваться.")
    bullet(doc, "Нейросети меняют задачи внутри профессий. Они не отменяют ответственность за другого человека и не заменяют умение поставить вопрос.")
    bullet(doc, "Востребованность — информация для выбора, а не рейтинг человеческой ценности.")
    h2(doc, "Где вы сейчас — такая и задача")
    for title, text in PATHS:
        p(doc, title, size=12, color=NAVY, bold=True, before=4, after=1)
        p(doc, text, size=11, after=2)
    h2(doc, "Навыки, которые переживают смену работы")
    p(doc, "Отметьте, что хотите тренировать в ближайший месяц. Не всё сразу.", size=11, after=2)
    for title, text in SKILLS:
        bullet(doc, f"☐  {title}. {text}")
    callout(
        doc,
        "Могу и умею — не одно и то же",
        "«Могу» — способность, которую вы в себе замечаете. «Умею» — то, что уже получалось и это можно показать. "
        "«Умею лечить людей» в школьные годы обычно неправда. «Умею объяснить тему так, что другой понял» — может быть правдой.",
        "FFF4D8",
        "A86B12",
    )


def _programs(doc):
    h1(doc, "Четыре направления высшего образования")
    p(
        doc,
        "Обязательный блок для решения кейса про Леру. Это реальные пути в вузе, если близки люди и общество. "
        "Не каталог «правильных» профессий: если ведущий тип — техника или природа, двери другие, и они не хуже. "
        "Коды нужны, чтобы найти программу на сайте вуза и на Госуслугах. "
        "Для Леры выберите одно направление (или два, если спорно) и объясните почему.",
        size=11,
        after=4,
    )
    for program in PROGRAMS:
        table = doc.add_table(rows=1, cols=1)
        set_table_widths(table, [W])
        set_table_borders(table, "E4DCD0", "8")
        cell = table.cell(0, 0)
        shade_cell(cell, "FFFFFF")
        set_cell_margins(cell, 80, 90, 120, 120)
        blocks = [
            {"t": f"{program['code']}    {program['title']}", "size": 13, "color": NAVY, "bold": True, "after": 1},
            {"t": program["question"], "size": 11, "color": "3A46A0", "italic": True, "after": 3},
            {"t": program["lead"], "size": 10.5, "after": 2},
            {"t": "Чему учат. " + program["study"], "size": 10.5, "after": 2},
            {"t": "Кому близко. " + program["fit"], "size": 10.5, "after": 2},
            {"t": "Куда идут работать. " + program["jobs"], "size": 10.5, "after": 2},
            {"t": "Поступление. " + program["ege"], "size": 10.5, "after": 2},
            {"t": "Формула. " + program["formula"], "size": 10.5, "after": 0},
        ]
        if program.get("law"):
            blocks.append({"t": program["law"], "size": 10, "italic": True, "color": MUTED, "before": 3, "after": 0})
        cell_text(cell, blocks)
        prevent_row_split(table.rows[0])
        spacer(doc, 6)
    callout(doc, "Общее для всех четырёх", ADMISSION_COMMON, "E7F3F1", TEAL)


def _case(doc):
    h1(doc, f"Кейс «{CASE['title']}»")
    p(doc, f"{CASE['hero']}, {CASE['grade']}", size=12, color=TEAL, bold=True, after=2)
    p(doc, CASE["situation"], size=11, after=3)
    p(doc, CASE["task"], size=11, after=3)
    p(doc, "Критерии хорошего ответа", size=12, color=NAVY, bold=True, before=2, after=2)
    for item in CASE["criteria"]:
        bullet(doc, item)
    h2(doc, "Интерактив 1. Сбор фактов")
    p(doc, CASE["interactive1"]["prompt"], size=11, after=2)
    for q in CASE["interactive1"]["questions"]:
        lines(doc, q, 2)
    h2(doc, "Подсказки из истории Леры")
    for cell, text in CASE["facts"]:
        p(doc, f"{cell}. {text}", size=11, after=2)
    h2(doc, "Карточка решения (заполняем в конце)")
    for i, step in enumerate(CASE["solve_steps"], start=1):
        lines(doc, f"{i}. {step}", 2)
    callout(
        doc,
        "Не смешивайте инструменты",
        "«Хочу — могу — умею — надо» — про человека. Формула профессии — про устройство работы. "
        "Тест ДДО — про ваш интерес к типу дел. Направления подготовки — про конкретные пути в вузе.",
        "FFF4D8",
        "A86B12",
    )


def _profession_formula_page(doc):
    h1(doc, "Формула профессии по Климову")
    p(
        doc,
        "Это не тест и не формула «хочу — могу — умею — надо». Так описывают любую профессию: "
        "предмет труда, цели, средства и условия. По предмету труда Климов выделяет пять типов — "
        "природа, техника, человек, знаки, образ. Типы — часть формулы профессии; тест ДДО идёт отдельно.",
        size=11,
        after=4,
    )
    for name, text in PROFESSION_FORMULA:
        p(doc, name, size=12, color=NAVY, bold=True, before=4, after=1)
        p(doc, text, size=11, after=2)
        lines(doc, "Пример для профессии, о которой думаете", 1)
    h2(doc, "Для кейса про Леру")
    lines(doc, "Какой предмет труда у Леры звучит ближе и почему", 2)
    callout(
        doc,
        "Дальше — отдельно тест",
        "Тест ДДО не заменяет формулу профессии. Он только показывает, к какому предмету труда вас сейчас тянет.",
        "E7F3F1",
        TEAL,
    )


def _support(doc):
    h1(doc, "Поддержка и три шага")
    p(doc, "Государственные двери полезно знать по имени. Открывают их вместе со взрослым: часть шагов требует учётной записи, подписи родителей или договора.", size=11, after=4)
    for item in SUPPORT:
        h2(doc, item["title"])
        p(doc, item["text"], size=11, after=2)
        p(doc, "С чего начать. " + item["start"], size=11, after=1)
        p(doc, "Осторожно. " + item["caution"], size=10.5, italic=True, color=MUTED, after=2)
    h2(doc, "Если вы в 10–11 классе")
    bullet(doc, "Заявление на бакалавриат подают летом после 11 класса. В кампанию 2026/27 года суперсервис «Поступление в вуз онлайн» открылся 20 июня. Дату своего года смотрите заново.")
    bullet(doc, "Ориентир той кампании: не больше 5 вузов и 5 направлений подготовки в каждом. На целевую квоту — один конкурс и один заказчик среди всех вузов.")
    bullet(doc, "Нужна подтверждённая учётная запись на Госуслугах. Очная подача и почта остаются.")
    bullet(doc, "Олимпиады из перечня Минобрнауки могут давать льготу или дополнительные баллы. Это история нескольких лет, а не вечера накануне экзамена. Условия каждой олимпиады свои.")
    bullet(doc, "Целевой договор читают дома до подписания. Если срок работы и обязанности непонятны, не подписывают.")
    h2(doc, "Три шага, которые помещаются в месяц")
    bullet(doc, "Гипотеза. Ведущий тип и клетка «хочу». Это версия, её можно изменить.")
    bullet(doc, "Проба. Одно дело из списка ниже, не пять.")
    bullet(doc, "Маршрут. Какие предметы держать в фокусе, с кем поговорить, какой официальный сайт открыть.")
    h2(doc, "Проба на месяц — выберите одну")
    for item in MONTH_TRIALS:
        bullet(doc, "☐  " + item)
    choice = p(doc, "Проба, которую я выбираю", size=11, bold=True, color=NAVY, before=2, after=2)
    _bottom_border(choice)


def _letter(doc):
    h1(doc, "Письмо себе из будущего")
    p(
        doc,
        "Напишите от своего будущего «я». Можно пятью предложениями, можно короче, в 5–6 классе можно рисунком на полях. "
        "Ведущий письмо не читает и не собирает. Дома его можно убрать в конверт и открыть через год.",
        size=11,
        after=3,
    )
    p(doc, "Подсказки можно переформулировать так, как вы говорите о себе.", size=10.5, italic=True, color=MUTED, after=6)
    banner(doc, "Письмо, которое не нужно сдавать", "16324A", "FFFFFF", 13)
    for prompt in LETTER_PROMPTS:
        lines(doc, prompt, 2 if prompt.startswith("Одна фраза") or prompt.startswith("Я учусь") else 1)
    spacer(doc, 6)
    p(doc, "Свободные строки, если хочется дописать.", size=10, color=MUTED, after=2)
    for _ in range(4):
        lines(doc, " ", 1)


def _home(doc):
    h1(doc, "Дома")
    h2(doc, "Пять вопросов семье")
    p(doc, "Не «кем ты будешь?», а так:", size=11, after=2)
    for i, item in enumerate(FAMILY_QUESTIONS, start=1):
        bullet(doc, f"{i}. {item}")
    h2(doc, "Пять вопросов человеку этой профессии")
    for i, item in enumerate(PRO_QUESTIONS, start=1):
        bullet(doc, f"{i}. {item}")
    h2(doc, "Куда заглянуть")
    bullet(doc, "«Билет в будущее» — bvbinfo.ru")
    bullet(doc, "«Поступление в вуз онлайн» — gosuslugi.ru/vuzonline")
    bullet(doc, "«Поступление в СПО онлайн» — сервис на Госуслугах")
    bullet(doc, "«Работа России», целевое обучение и официальная занятость — trudvsem.ru")
    bullet(doc, "«Движение Первых» — pervye.ru  (до 14 лет регистрирует родитель)")
    h2(doc, "Откуда сведения")
    for item in SOURCES:
        bullet(doc, item)
    spacer(doc, 6)
    callout(
        doc,
        "Последняя строка",
        "Профессию сегодня назначать не нужно. Нужны гипотеза, одна проба и человек, с которым это можно обсудить без насмешки.",
        "E7F3F1",
        TEAL,
    )


if __name__ == "__main__":
    print(build())
