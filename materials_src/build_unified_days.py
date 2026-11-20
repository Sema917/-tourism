"""Приказы о днях единых действий: методические материалы.

Два проекта приказа ОГУ имени И.С. Тургенева. В исходных черновиках
приложение 1 было скопировано с Дня народного единства. Здесь у каждого
праздника свои рекомендации, каркас недели и форматы.

Пересобрать из корня репозитория:
    python3 materials_src/build_unified_days.py
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "materials"
FONT = "Times New Roman"
BLACK = RGBColor(0x00, 0x00, 0x00)
HEADER_FILL = "E8EEF2"
LINE = "8AA0B2"


def _rfonts(rPr) -> None:
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)


def set_run(run, size=13, bold=False, italic=False) -> None:
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = BLACK
    _rfonts(run._element.get_or_add_rPr())


def _paragraph(
    doc,
    text="",
    *,
    align="left",
    size=13,
    bold=False,
    italic=False,
    before=0,
    after=6,
    first=0,
    left=0,
    line=1.08,
    keep_next=False,
    page_break=False,
    runs=None,
):
    paragraph = doc.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if first:
        fmt.first_line_indent = Cm(first)
    if left:
        fmt.left_indent = Cm(left)
    fmt.keep_together = False
    fmt.keep_with_next = keep_next
    fmt.page_break_before = page_break
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    paragraph.alignment = align_map[align]
    if runs is None:
        runs = [(text, bold, italic)]
    for item in runs:
        value, is_bold, is_italic = item if len(item) == 3 else (item[0], item[1], False)
        run = paragraph.add_run(value)
        set_run(run, size=size, bold=is_bold, italic=is_italic)
    return paragraph


def body(doc, text, **kwargs):
    opts = dict(align="justify", first=1.25, before=0, after=6, line=1.08)
    opts.update(kwargs)
    return _paragraph(doc, text, **opts)


def heading(doc, text, *, page_break=False):
    return _paragraph(
        doc,
        text,
        align="center",
        bold=True,
        before=14,
        after=6,
        line=1.08,
        keep_next=True,
        page_break=page_break,
    )


def format_title(doc, text):
    return _paragraph(
        doc,
        text,
        align="center",
        bold=True,
        before=12,
        after=4,
        line=1.08,
        keep_next=True,
    )


def description(doc, text):
    return _paragraph(
        doc,
        align="justify",
        first=1.25,
        before=0,
        after=4,
        line=1.08,
        runs=[("Описание: ", True, False), (text, False, False)],
    )


def nomination(doc, number, text):
    return _paragraph(
        doc,
        align="justify",
        first=0,
        left=1.25,
        before=0,
        after=2,
        line=1.08,
        runs=[(f"{number}. ", True, False), (text, False, False)],
    )


def shade(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "1C2830")


def set_cell_margins(cell, dxa=70) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = tc_pr.find(qn("w:tcMar"))
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tc_pr.append(mar)
    for edge in ("top", "left", "bottom", "right"):
        node = mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            mar.append(node)
        node.set(qn("w:w"), str(dxa))
        node.set(qn("w:type"), "dxa")


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant = tr_pr.find(qn("w:cantSplit"))
    if cant is None:
        cant = OxmlElement("w:cantSplit")
        tr_pr.append(cant)


def set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = tr_pr.find(qn("w:tblHeader"))
    if header is None:
        header = OxmlElement("w:tblHeader")
        tr_pr.append(header)


def set_row_height(row, twips: int) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    height = tr_pr.find(qn("w:trHeight"))
    if height is None:
        height = OxmlElement("w:trHeight")
        tr_pr.append(height)
    height.set(qn("w:val"), str(twips))
    height.set(qn("w:hRule"), "atLeast")


def set_table_widths(table, widths_cm: list[float]) -> None:
    table.autofit = False
    table.allow_autofit = False
    total = int(sum(widths_cm) * 567)
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")

    indent = tbl_pr.find(qn("w:tblInd"))
    if indent is None:
        indent = OxmlElement("w:tblInd")
        tbl_pr.append(indent)
    indent.set(qn("w:w"), "0")
    indent.set(qn("w:type"), "dxa")

    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
    for child in list(grid):
        grid.remove(child)
    for width in widths_cm:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(width * 567)))
        grid.append(col)

    for row in table.rows:
        for cell, width in zip(row.cells, widths_cm):
            cell.width = Cm(width)
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(width * 567)))
            tc_w.set(qn("w:type"), "dxa")


def write_cell(cell, text, *, bold=False, size=11, align="left", fill=None) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1.0
    run = paragraph.add_run(text)
    set_run(run, size=size, bold=bold)
    set_cell_border(cell)
    set_cell_margins(cell)
    if fill:
        shade(cell, fill)


def add_table(doc, headers, rows, widths, *, header_size=10, body_size=11, empty_rows=0, empty_height=420):
    table = doc.add_table(rows=1 + len(rows) + empty_rows, cols=len(headers))
    table.style = "Table Grid"
    set_table_widths(table, widths)
    for cell, header in zip(table.rows[0].cells, headers):
        write_cell(cell, header, bold=True, size=header_size, align="center", fill=HEADER_FILL)
    set_repeat_header(table.rows[0])
    prevent_row_split(table.rows[0])
    set_row_height(table.rows[0], 500)
    for index, row_values in enumerate(rows, start=1):
        for cell, value in zip(table.rows[index].cells, row_values):
            write_cell(cell, value, size=body_size, align="left")
        prevent_row_split(table.rows[index])
    for index in range(1 + len(rows), 1 + len(rows) + empty_rows):
        for cell in table.rows[index].cells:
            write_cell(cell, "", size=body_size)
        set_row_height(table.rows[index], empty_height)
    return table


def signature(doc, left: str, right: str, *, before=12) -> None:
    table = doc.add_table(rows=1, cols=2)
    set_table_widths(table, [11.5, 5.0])
    left_cell, right_cell = table.rows[0].cells
    for cell in (left_cell, right_cell):
        set_cell_margins(cell, 40)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    lines = left.split("\n")
    first = left_cell.paragraphs[0]
    first.paragraph_format.space_before = Pt(before)
    first.paragraph_format.space_after = Pt(0)
    first.paragraph_format.line_spacing = 1.0
    run = first.add_run(lines[0])
    set_run(run, size=13)
    for line in lines[1:]:
        extra = left_cell.add_paragraph()
        extra.paragraph_format.space_before = Pt(0)
        extra.paragraph_format.space_after = Pt(0)
        extra.paragraph_format.line_spacing = 1.0
        run = extra.add_run(line)
        set_run(run, size=13)
    right_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right_cell.paragraphs[0].paragraph_format.space_before = Pt(before)
    right_cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    run = right_cell.paragraphs[0].add_run(right)
    set_run(run, size=13)
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        borders.append(element)
    tbl_pr.append(borders)


def setup_portrait(section) -> None:
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Twips(11906)
    section.page_height = Twips(16838)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(1.7)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.6)


def setup_landscape(section) -> None:
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Twips(16838)
    section.page_height = Twips(11906)
    section.left_margin = Cm(1.4)
    section.right_margin = Cm(1.4)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)


def configure(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(13)
    normal.font.color.rgb = BLACK
    _rfonts(normal.element.get_or_add_rPr())
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.line_spacing = 1.08
    section = doc.sections[0]
    setup_portrait(section)
    lang = normal.element.get_or_add_rPr().find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        normal.element.get_or_add_rPr().append(lang)
    lang.set(qn("w:val"), "ru-RU")
    lang.set(qn("w:eastAsia"), "ru-RU")


def appendix_caption(doc, number: str, *, page_break=True) -> None:
    _paragraph(
        doc,
        f"Приложение №{number} к приказу",
        align="right",
        before=0,
        after=0,
        page_break=page_break,
    )
    _paragraph(doc, "№ ______ от ______________", align="right", before=0, after=8)


def add_order(doc: Document, spec: dict) -> None:
    _paragraph(doc, "О проведении мероприятий, приуроченных", align="left", before=0, after=0)
    _paragraph(doc, spec["title_line"], align="left", before=0, after=10)
    body(doc, spec["purpose"], after=8)
    _paragraph(doc, "ПРИКАЗЫВАЮ:", bold=True, before=4, after=6)
    for item in spec["points"]:
        body(doc, item, after=4)
    signature(doc, "И.о. первого проректора\nпо образовательной деятельности", "О.В. Леонова", before=14)
    _paragraph(doc, "Согласовано:", before=10, after=2)
    signature(doc, "И.о. проректора по МП и СВР", "Е.В. Сенько", before=2)
    signature(doc, "Начальник управления по воспитательной работе", "Н.Н. Васильева", before=2)
    _paragraph(doc, "Проект приказа вносит:", before=10, after=0)
    _paragraph(doc, "Отдел организации ВР", before=0, after=0)
    signature(doc, "Начальник отдела", "А.Г. Семин", before=2)


def add_formats(doc: Document, spec: dict) -> None:
    heading(doc, "Рекомендованные форматы")
    body(doc, spec["formats_note"], after=2)
    formats = spec["formats"]
    for item in formats:
        format_title(doc, item["title"])
        description(doc, item["lead"])
        for number, nomination_text in enumerate(item.get("items", []), start=1):
            nomination(doc, number, nomination_text)
        if item.get("tail"):
            body(doc, item["tail"], before=2, after=4)


def add_appendix_1(doc: Document, spec: dict) -> None:
    appendix_caption(doc, "1")
    _paragraph(doc, spec["materials_title"], align="center", bold=True, before=6, after=2, keep_next=True)
    _paragraph(doc, spec["dates"], align="center", bold=True, before=0, after=0, keep_next=True)
    _paragraph(doc, spec["memorial"], align="center", bold=True, before=0, after=8, keep_next=True)
    for paragraph in spec["intro"]:
        body(doc, paragraph)
    heading(doc, "Рекомендации организаторам")
    for index, text in enumerate(spec["advice"], start=1):
        body(doc, f"{index}. {text}", after=3)
    heading(doc, "Примерное распределение недели")
    body(doc, spec["week_note"], after=4)
    add_table(
        doc,
        ["День", "Что проводить"],
        spec["week"],
        [4.0, 12.2],
        header_size=11,
        body_size=12,
    )
    add_formats(doc, spec)
    signature(doc, "Начальник отдела организации ВР", "А.Г. Семин", before=14)


def add_forms(doc: Document, spec: dict) -> None:
    section = doc.add_section()
    setup_landscape(section)
    appendix_caption(doc, "2", page_break=False)
    _paragraph(
        doc,
        "Сведения о запланированных мероприятиях",
        align="center",
        bold=True,
        before=2,
        after=6,
        keep_next=True,
    )
    plan_headers = [
        "Название мероприятия",
        "Краткое описание",
        "Формат проведения (очно / дистанционно / очно и дистанционно)",
        "Место проведения",
        "Дата",
        "Время",
        "Планируемое число участников",
        "Почётные гости (планируются / не планируются)",
        "ФИО ответственного",
        "Контактный телефон",
    ]
    add_table(
        doc,
        plan_headers,
        [],
        [3.2, 3.0, 2.8, 2.6, 2.2, 1.8, 2.4, 2.7, 2.9, 2.9],
        header_size=9,
        empty_rows=6,
        empty_height=480,
    )
    signature(doc, "Начальник отдела организации ВР", "А.Г. Семин", before=12)

    appendix_caption(doc, "3")
    _paragraph(
        doc,
        spec["report_title"],
        align="center",
        bold=True,
        before=2,
        after=6,
        keep_next=True,
    )
    add_table(
        doc,
        [
            "№",
            "Название мероприятия",
            "Дата проведения",
            "Число участников",
            "Из них иностранных обучающихся",
            "Ссылка на электронный ресурс",
        ],
        [],
        [1.4, 7.2, 3.4, 3.6, 4.4, 6.5],
        header_size=11,
        empty_rows=8,
        empty_height=420,
    )
    signature(doc, "Начальник отдела организации ВР", "А.Г. Семин", before=12)


def build(spec: dict, path: Path) -> None:
    doc = Document()
    configure(doc)
    core = doc.core_properties
    core.author = "Отдел организации воспитательной работы"
    core.title = spec["core_title"]
    core.category = "День единых действий"
    core.language = "ru-RU"
    add_order(doc, spec)
    add_appendix_1(doc, spec)
    add_forms(doc, spec)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


UNKNOWN = {
    "core_title": "День единых действий, посвящённый Дню Неизвестного Солдата",
    "title_line": "к празднованию Дня Неизвестного Солдата",
    "purpose": (
        "В целях укрепления уважения к истории страны, воинской доблести и подвигам защитников Отечества "
        "среди обучающихся, а также увековечивания памяти о российских и советских воинах, погибших "
        "в боевых действиях и пропавших без вести,"
    ),
    "points": [
        "1. Организовать и провести День единых действий, посвящённый Дню Неизвестного Солдата, с 30 ноября 2026 г. по 5 декабря 2026 г. для обучающихся ОГУ имени И.С. Тургенева.",
        "2. Утвердить методические материалы по проведению Дня единых действий (Приложение 1).",
        "3. Директорам филиалов, директору медицинского института, деканам факультетов проинформировать обучающихся о проведении мероприятий, организовать их участие, предоставить информацию о запланированных мероприятиях по форме (Приложение 2) в срок до 27 ноября 2026 г. и отчёт о проведённых мероприятиях в срок до 11 декабря 2026 г. по прилагаемой форме (Приложение 3) на адрес электронной почты student_life@oreluniver.ru.",
        "4. Управлению документооборота разместить приказ в локальной сети.",
        "5. Ответственность за проведение мероприятий возложить на начальника отдела организации ВР А.Г. Семина.",
        "6. Контроль за исполнением приказа возложить на проректора по молодёжной политике и социально-воспитательной работе Е.В. Сенько.",
    ],
    "materials_title": "Методические материалы по проведению Дня единых действий, посвящённого Дню Неизвестного Солдата",
    "dates": "30.11.2026 г. – 05.12.2026 г.",
    "memorial": "Памятная дата — 3 декабря",
    "intro": [
        "День единых действий проводится в память о российских и советских воинах, погибших в боях и пропавших без вести, чьи имена остались неизвестными. Памятная дата 3 декабря связана с 3 декабря 1966 года: прах неизвестного солдата перенесли из братской могилы на 41-м километре Ленинградского шоссе и захоронили в Александровском саду у Кремлёвской стены. Надпись на плите — «Имя твоё неизвестно, подвиг твой бессмертен» — задаёт тон всей недели. Это тихие памятные действия, а не праздничная программа.",
        "Орёл — город воинской славы, поэтому рядом с общероссийской датой стоит местная история. 5 августа 1943 года город был освобождён, и в тот же день Москва впервые дала салют — Орлу и Белгороду. Площадки недели: сквер Танкистов (братская могила, Вечный огонь, танк на постаменте) и музей-диорама «Орловская наступательная операция» на улице Нормандия-Неман. Филиалы не везут обучающихся в орловский сквер: те же форматы проводят у местного мемориала.",
    ],
    "advice": [
        "Брать имена, части и даты только из ОБД «Мемориал», портала «Память народа», книг памяти Орловской области и музейных подписей. Биографию без источника не зачитывать и не печатать. Текст, собранный без ссылки, к читке и к публикации не допускать.",
        "Держать тишину у мемориала. Минута молчания и возложение цветов — центр недели. Громкую музыку, флешмобы и конкурсы с призами на ту же площадку и в тот же час не ставить. На плиты и венки не вставать, съёмку для отчёта делать до или после минуты молчания.",
        "Разослать факультетам, медицинскому институту и филиалам один текст справки и каркас открытого разговора до 30 ноября. Иначе группы услышат разные версии одних и тех же событий. Карточки акции «Верни имя» собрать организатору до 27 ноября, к сроку подачи планов. Поиск судьбы не перекладывать на участника в день мероприятия.",
        "Считать участие добровольным. Приглашать обучающихся, в том числе иностранных, преподавателей, работников и студенческие семьи. Иностранным обучающимся в начале занятия пояснить слова «мемориал», «братская могила», «Вечный огонь», «подвиг». Семейную военную историю от них не требовать. Иностранных обучающихся считать отдельно: эта графа есть в отчёте.",
        "Не подменять возложение 3 декабря публикацией. Пост и ролик дополняют церемонию, но не заменяют её. В отчёт (Приложение 3) входят дата, число участников, число иностранных обучающихся и ссылка на публикацию.",
    ],
    "week_note": "Ниже — удобный ход недели. Дни можно сдвинуть внутри 30 ноября – 5 декабря, но вахту памяти оставить на 3 декабря. Если площадка в этот день занята, возложение переносят на ближайший день и не выносят за пределы недели.",
    "formats_note": "Форматы носят рекомендательный характер. Подразделение проводит вахту памяти 3 декабря и ещё хотя бы один формат, остальные подключает по силам и по площадке. Форматы можно совмещать: инсталляцию с читкой имён, открытый разговор с письмом в инсталляцию, маршрут с возложением.",
    "week": [
        ["30 ноября, понедельник", "Открытие инсталляции «Солдатский треугольник». На стенде уже лежат проверенные карточки акции «Верни имя»."],
        ["1–2 декабря", "Открытые разговоры «Солдат без имени» в академических группах и на кураторских часах, очно или дистанционно."],
        ["3 декабря, четверг", "Вахта памяти и возложение цветов. Основной день недели."],
        ["4 декабря", "Читка имён по карточкам акции и публикация материалов медиапроекта «Одна история»."],
        ["5 декабря, суббота", "Маршрут памяти, в том числе для студенческих семей. Филиалы проводят его у своего мемориала."],
    ],
    "formats": [
        {
            "title": "Формат 1. Вахта памяти и возложение цветов",
            "lead": (
                "Общее памятное действие у воинского захоронения. В Орле его проводят в сквере Танкистов. "
                "Филиалы выбирают мемориал, обелиск или братскую могилу своего города. Церемония короткая: "
                "построение, справка на три-четыре минуты, минута молчания, возложение цветов. В справке два факта "
                "из общего текста: Могила Неизвестного Солдата, 3 декабря 1966 года, и освобождение Орла 5 августа 1943 года. "
                "Длинных речей нет, громкую аппаратуру у огня не ставят. Строй у мемориала могут держать обучающиеся "
                "студенческих патриотических объединений. Участвуют обучающиеся, преподаватели, работники и студенческие семьи. "
                "Это главное действие недели."
            ),
        },
        {
            "title": "Формат 2. Акция «Верни имя»",
            "lead": (
                "Акция о том, что у неизвестного солдата было имя. Организатор заранее отбирает короткие карточки "
                "воинов, связанных с Орловской областью или с городом филиала: фамилия, часть, дата и место гибели "
                "или пропажи, ссылка на источник. Источники — ОБД «Мемориал», портал «Память народа», книги памяти. "
                "Обучающийся разбирает готовую карточку и может дополнить её только той же ссылкой, а затем называет "
                "имя на читке. Новые, ещё не проверенные сведения на читку и на стенд не выносят: их передают в музей "
                "или поисковое объединение. Читка проходит у инсталляции или после возложения. Акция учит работать "
                "с документом, а не с легендой."
            ),
        },
        {
            "title": "Формат 3. Открытый разговор «Солдат без имени»",
            "lead": (
                "Занятие на 45–60 минут в академической группе или на кураторском часе, очно или дистанционно. "
                "Содержание общее для всех подразделений: почему памятная дата приходится на 3 декабря; как появилась "
                "Могила Неизвестного Солдата; чем братская могила отличается от именного захоронения; как устроены "
                "донесение о безвозвратных потерях и извещение о гибели; что можно и чего нельзя делать с чужой "
                "биографией в интернете. Занятие заканчивается одним действием группы: письмо в инсталляцию, участие "
                "в читке имён или выход на возложение. Развлекательные конкурсы и шутки на военную тему в это занятие не включают."
            ),
        },
        {
            "title": "Формат 4. Маршрут памяти",
            "lead": (
                "Пешеходная экскурсия по местам воинской памяти. В Орле основа маршрута — сквер Танкистов и, если "
                "позволяет время, музей-диорама «Орловская наступательная операция» на улице Нормандия-Неман. "
                "На каждой точке ведущий говорит пять–семь минут по единому тексту: что здесь произошло, какие имена "
                "установлены, кто остался безымянным. Маршрут можно сократить до одной площадки. Для непогоды и для тех, "
                "кто не может выйти в город, готовят заочный вариант: фотографии тех же мест и короткий рассказ "
                "о Могиле Неизвестного Солдата. Филиалы строят маршрут от своего мемориала. Ведут экскурсию подготовленные "
                "студенты, текст заранее проверяет преподаватель истории. Участвуют обучающиеся, студенческие семьи и преподаватели. "
                "Удобный день — суббота, 5 декабря."
            ),
        },
        {
            "title": "Формат 5. Инсталляция «Солдатский треугольник»",
            "lead": (
                "Тихая выставка в холле учебного корпуса на всю неделю. На стенде — копии писем-треугольников, "
                "фотографии безымянных захоронений, карта с отметкой 41-го километра Ленинградского шоссе и Орловской "
                "наступательной операции, текст надписи «Имя твоё неизвестно, подвиг твой бессмертен». У каждого "
                "изображения есть подпись источника. Рядом стоит стол: участник пишет короткое письмо неизвестному "
                "солдату и оставляет его в общей композиции. Письма не зачитывают вслух и не оценивают. Рядом размещают "
                "стенд акции «Верни имя». Это не ярмарка и не фотозона: у инсталляции не включают музыку и не проводят шумные акции. "
                "Выставка принимает тех, кто не был на церемонии возложения."
            ),
        },
        {
            "title": "Формат 6. Медиапроект «Одна история»",
            "lead": (
                "Серия публикаций студенческих СМИ и официальных ресурсов университета. Правило одно: одна публикация — "
                "одна проверенная судьба — один источник. Судьбу берут из карточек акции «Верни имя» или из музейной "
                "подписи, новую биографию для ролика не собирают. Текст, карточка или ролик — не длиннее полутора минут. "
                "В кадре нет жестоких изображений, постановки боя и подробностей, которых нет в документе. В конце — "
                "ссылка на ОБД «Мемориал», «Память народа» или музейную подпись. Перед выходом материал смотрит "
                "ответственный, назначенный подразделением. Филиалы присылают публикации на тот же ресурс. Отобранные "
                "материалы оставляют на сайте университета подборкой, а не снимают на следующий день."
            ),
        },
    ],
    "report_title": "Форма отчёта о проведении Дня единых действий, посвящённого Дню Неизвестного Солдата",
}


HEROES = {
    "core_title": "День единых действий, посвящённый Дню Героев Отечества",
    "title_line": "к празднованию Дня Героев Отечества",
    "purpose": (
        "В целях формирования у обучающихся уважения к защитникам Отечества, гражданственности "
        "и чувства гордости за подвиги соотечественников,"
    ),
    "points": [
        "1. Организовать и провести День единых действий, посвящённый Дню Героев Отечества, с 7 декабря 2026 г. по 12 декабря 2026 г. для обучающихся ОГУ имени И.С. Тургенева.",
        "2. Утвердить методические материалы по проведению Дня единых действий (Приложение 1).",
        "3. Директорам филиалов, директору медицинского института, деканам факультетов проинформировать обучающихся о проведении мероприятий, организовать их участие, предоставить информацию о запланированных мероприятиях по форме (Приложение 2) в срок до 4 декабря 2026 г. и отчёт о проведённых мероприятиях в срок до 18 декабря 2026 г. по прилагаемой форме (Приложение 3) на адрес электронной почты student_life@oreluniver.ru.",
        "4. Управлению документооборота разместить приказ в локальной сети.",
        "5. Ответственность за проведение мероприятий возложить на начальника отдела организации ВР А.Г. Семина.",
        "6. Контроль за исполнением приказа возложить на проректора по молодёжной политике и социально-воспитательной работе Е.В. Сенько.",
    ],
    "materials_title": "Методические материалы по проведению Дня единых действий, посвящённого Дню Героев Отечества",
    "dates": "07.12.2026 г. – 12.12.2026 г.",
    "memorial": "Памятная дата — 9 декабря",
    "intro": [
        "День единых действий приурочен к 9 декабря — Дню Героев Отечества. В этот день чествуют Героев Советского Союза, Героев Российской Федерации, кавалеров ордена Святого Георгия и кавалеров ордена Славы. Памятная дата установлена Федеральным законом от 28 февраля 2007 г. № 22-ФЗ. Она связана с орденом Святого Георгия, который Екатерина II учредила в 1769 году: день ордена был положен на 26 ноября по старому стилю, а современная дата совпадает с днём памяти Георгия Победоносца. Девиз ордена — «За службу и храбрость».",
        "Мероприятия показывают, что у подвига есть имя: от георгиевских кавалеров и участников Отечественной войны 1812 года до защитников Орла в 1943 году и Героев Российской Федерации. Орёл — город воинской славы. В городе стоят памятник генералу А. П. Ермолову, участнику Отечественной войны 1812 года, сквер Танкистов и музей-диорама «Орловская наступательная операция». Биографии берут из наградных документов, музейных справок и официальных изданий. Живого Героя, ветерана или родственника приглашают только с его согласия и не просят рассказывать то, о чём человек говорить не хочет.",
    ],
    "advice": [
        "Называть награду точно. Георгиевский крест, орден Святого Георгия, орден Славы, звание Героя Советского Союза и звание Героя Российской Федерации — разные награды разных эпох. В справке для ведущего каждую называют отдельно и коротко поясняют, кому и за что её вручали.",
        "Не добавлять подвигу подробностей, которых нет в источнике. Портрет, карточку и рассказ без ссылки на наградной документ, музейную справку или официальное издание не вывешивать и не читать. Если проверенной биографии земляка нет, берут опубликованную биографию и прямо говорят, откуда она.",
        "Беречь живого человека. Вопросы гостю передают заранее. Не спрашивают о ходе операций, о погибших товарищах и о семье, если гость сам этого не открыл. Съёмку и публикацию включают только с его разрешения. Имя живущего человека и конкретной семьи в конкурсе и на галерее ставят только с их согласия.",
        "Развести память и игру. На станциях уместно разобрать степени ордена, прочитать девиз и сопоставить ленты. Не проводить костюмированный розыгрыш, угадывание по тяжёлым фотографиям и конкурсы, где проигрыш связан с темой подвига.",
        "Собрать неделю вокруг 9 декабря: в этот день — встреча или торжественное собрание. Выставку можно открыть раньше, читку рассказов — позже. Факультету, медицинскому институту и филиалу достаточно общей встречи или галереи и одного формата, где обучающийся сам работает с биографией.",
        "Филиал ставит в центр героев и памятные места своего города. Иностранным обучающимся поясняют, что означают награды. Личный семейный рассказ не является условием участия. Иностранных обучающихся считают отдельно для отчёта.",
    ],
    "week_note": "Ниже — удобный ход недели. Дни можно сдвинуть внутри 7–12 декабря, но встречу или торжественное собрание оставить на 9 декабря. Если площадка в этот день занята, встречу переносят на ближайший день и не выносят за пределы недели.",
    "formats_note": "Форматы носят рекомендательный характер. Подразделение проводит встречу или торжественное собрание 9 декабря и ещё хотя бы один формат, остальные подключает по силам и по площадке. Форматы можно совмещать: галерею со станциями, конкурс с вечером, городской маршрут с галереей в корпусе.",
    "week": [
        ["7 декабря, понедельник", "Открытие галереи «Имена Орловщины» и станций «За службу и храбрость»."],
        ["8 декабря", "Последний день приёма рассказов на конкурс «Мой герой». Подготовка вечера «Живое имя»."],
        ["9 декабря, среда", "Встреча «Герой рядом» или торжественное собрание. Основной день недели."],
        ["10–11 декабря", "Вечер «Живое имя» и чтение лучших рассказов конкурса."],
        ["12 декабря, суббота", "Маршрут «Герои города», в том числе для студенческих семей."],
    ],
    "formats": [
        {
            "title": "Формат 1. Встреча «Герой рядом»",
            "lead": (
                "Очный разговор с Героем Российской Федерации, кавалером государственной награды, ветераном "
                "боевых действий, членом семьи героя или сотрудником музея. Если гость приехать не может, проводят "
                "портретную лекцию по опубликованной биографии. Встреча идёт как беседа: короткое представление, "
                "рассказ гостя, вопросы. Вопросы собирают заранее и отдают гостю до встречи. Для филиала допустима "
                "видеосвязь, если гость на неё согласен. Встречу ставят на 9 декабря. Выставку, конкурс и вечер "
                "проводят в том же спокойном тоне."
            ),
        },
        {
            "title": "Формат 2. Галерея «Имена Орловщины»",
            "lead": (
                "Выставка портретов в холлах корпусов и на электронной витрине сайта, работает всю неделю. "
                "В галерее — люди разных эпох, связанные с Орлом и областью: участники Отечественной войны 1812 года, "
                "включая генерала А. П. Ермолова; Герои Советского Союза и полные кавалеры ордена Славы — уроженцы "
                "и защитники Орловщины; Герои Российской Федерации. Условие одно: биография подтверждена источником. "
                "У портрета указывают имя, годы жизни, точное название награды, пять–шесть строк о поступке и ссылку. "
                "Карточку без источника не печатают. Подбор согласуют с музеем-диорамой или с преподавателем истории. "
                "Обучающийся готовит карточку по выданному источнику, а не по случайной странице в интернете. "
                "Филиал собирает галерею своего города. Выставка остаётся открытой для тех, кто не попал на встречу 9 декабря."
            ),
        },
        {
            "title": "Формат 3. Конкурс рассказов «Мой герой»",
            "lead": (
                "Конкурс коротких рассказов среди обучающихся, преподавателей, работников и выпускников. "
                "Объём — до двух страниц. В тексте указано, откуда автор знает факты. Конкурс объявляют сразу "
                "после издания приказа, работы принимают до 8 декабря. Если приказ выйдет позже, срок приёма сдвигают, "
                "а чтение лучших работ оставляют внутри недели, 10 или 11 декабря. Тексты публикуют на сайте университета. "
                "Номинации:"
            ),
            "items": [
                "«Герой моей семьи» — о родственнике, который защищал Отечество. Для публикации нужно согласие семьи, если в тексте есть имена.",
                "«Имя на карте города» — о человеке, чьим именем названы улица, школа или мемориал. Факты сверяют с краеведческим источником.",
                "«За службу и храбрость» — о конкретном поступке и о награде, которой этот поступок отмечен. Название награды приводят точно.",
                "«Письмо герою» — обращение к герою прошлого или к живущему Герою. Живого адресата называют только с его согласия.",
            ],
            "tail": "Чтение лучших рассказов совмещают с вечером «Живое имя». Работы без указания источника к чтению и к публикации не допускают.",
        },
        {
            "title": "Формат 4. Станции «За службу и храбрость»",
            "lead": (
                "Открытые просветительские станции в людном холле корпуса. Запись и плата не нужны: участник подходит, "
                "слушает и разбирает лист. Станций три. Первая — орден Святого Георгия: четыре степени, девиз "
                "«За службу и храбрость», правило статута, по которому награду давали за личный поступок, а не за чин "
                "и прежние заслуги. Вторая — солдатские награды: Георгиевский крест, знак отличия для нижних чинов "
                "с 1807 года, и орден Славы трёх степеней, учреждённый в 1943 году. Здесь же показывают, что "
                "чёрно-оранжевая лента ордена Славы восходит к ленте ордена Святого Георгия. Третья — медаль "
                "«Золотая Звезда»: звание Героя Советского Союза с 1934 года и звание Героя Российской Федерации "
                "с 1992 года. На станциях ничего не продают и не разыгрывают. Участник уносит одностраничный лист "
                "с названиями наград, а не сувенир «под орден». Станции удобно открыть 7 декабря рядом с галереей портретов."
            ),
        },
        {
            "title": "Формат 5. Маршрут «Герои города»",
            "lead": (
                "Пешеходная экскурсия с короткими остановками. В Орле в маршрут входят памятник А. П. Ермолову, "
                "сквер Танкистов и, при запасе времени, музей-диорама «Орловская наступательная операция» на улице "
                "Нормандия-Неман. На остановке — пять–семь минут: кто этот человек или кто помянут на этом месте, "
                "какая награда или какое событие с ним связаны. Ведущие идут по одному проверенному тексту. "
                "Маршрут сокращают до одной точки, если группа небольшая или погода плохая. Запасной вариант — "
                "пройти галерею «Имена Орловщины» в корпусе как внутренний маршрут. Филиалы выстраивают путь "
                "по своим мемориалам и по местам, связанным с именами героев. Удобный день — суббота, 12 декабря. "
                "Участвуют обучающиеся, студенческие семьи и преподаватели."
            ),
        },
        {
            "title": "Формат 6. Вечер «Живое имя»",
            "lead": (
                "Литературно-документальный вечер: три-четыре минуты на одного человека. Обучающийся рассказывает "
                "о герое по письму, наградному листу или опубликованным воспоминаниям. Это не костюмированная картина "
                "и не угадывание сюжета. Костюм не обязателен. Текст заранее сверяют с источником, вымысел в биографию "
                "не добавляют. Зрителям выдают карточку награды, о которой идёт речь: название и одна фраза, за что её давали. "
                "На вечер переносят лучшие рассказы конкурса «Мой герой». Вечер проводят 10 или 11 декабря, когда встреча "
                "9 декабря уже прошла и у аудитории есть живое впечатление, с которым можно сверить документ. "
                "Материалы вечера публикуют вместе с рассказами конкурса."
            ),
        },
    ],
    "report_title": "Форма отчёта о проведении Дня единых действий, посвящённого Дню Героев Отечества",
}


def validate(path: Path, required: list[str], forbidden: list[str]) -> None:
    from docx import Document as Open

    doc = Open(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text += "\n" + cell.text
    missing = [item for item in required if item not in text]
    leaked = [item for item in forbidden if item in text]
    if missing or leaked:
        raise SystemExit(f"{path.name}: missing={missing} leaked={leaked}")
    if "Описание:" not in text or text.count("Формат ") < 6:
        raise SystemExit(f"{path.name}: formats incomplete")


def main() -> None:
    unknown_path = OUT / "Приказ_День_Неизвестного_Солдата.docx"
    heroes_path = OUT / "Приказ_День_Героев_Отечества.docx"
    build(UNKNOWN, unknown_path)
    build(HEROES, heroes_path)
    forbidden = [
        "народного единства",
        "Минин",
        "Пожарск",
        "мастер-класс",
        "кулебяк",
        "Живые картины",
        "защитникам Отечества среди",
        "советских войнах",
        "со 30 ноября",
    ]
    validate(
        unknown_path,
        [
            "3 декабря 1966",
            "41-м километре",
            "Имя твоё неизвестно, подвиг твой бессмертен",
            "сквере Танкистов",
            "Верни имя",
            "27 ноября 2026",
            "11 декабря 2026",
            "student_life@oreluniver.ru",
            "защитников Отечества",
            "советских воинах",
            "с 30 ноября 2026",
            "А.Г. Семин",
            "Е.В. Сенько",
        ],
        forbidden,
    )
    validate(
        heroes_path,
        [
            "28 февраля 2007 г. № 22-ФЗ",
            "26 ноября",
            "За службу и храбрость",
            "9 декабря",
            "Ермолов",
            "4 декабря 2026",
            "18 декабря 2026",
            "Мой герой",
            "student_life@oreluniver.ru",
            "А.Г. Семин",
        ],
        forbidden,
    )
    print(unknown_path)
    print(heroes_path)


if __name__ == "__main__":
    main()
