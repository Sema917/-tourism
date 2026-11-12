#!/usr/bin/env python3
"""Генерация графика работы лифтов (2/2 и 5/2) на 3 недели."""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT = Path(__file__).resolve().parent / "График_работы_лифты.docx"

# Период: 3 недели с понедельника 16.11.2026
START = date(2026, 11, 16)
DAYS_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
WORKERS = ("№1", "№2", "№3")

# Кто на 5/2 в каждую неделю (остальные двое — на 2/2)
WEEK_5_2 = {0: "№1", 1: "№2", 2: "№3"}

# Паттерны 2/2 (W/O) на 21 день — противофаза, непрерывный цикл
# №2 и №3 стартуют в противофазе; при смене состава дополняют текущего
PATTERN_2_2 = {
    # W2: WW OO WW O | (неделя 2 на 5/2) | WW OO WW O
    "№2": list("WWOOWWO") + [None] * 7 + list("WWOOWWO"),
    # W3: OO WW OO W | WO OW WO O | (неделя 3 на 5/2)
    "№3": list("OOWWOOW") + list("WOOWWOO") + [None] * 7,
    # W1: (неделя 1 на 5/2) | OW WO OW W | OO WW OO W
    "№1": [None] * 7 + list("OWWOOWW") + list("OOWWOOW"),
}


def set_cell_shading(cell, hex_color: str) -> None:
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def set_run_font(run, *, bold=False, size=10, color=None) -> None:
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Arial"
    r = run._element
    r_pr = r.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), "Arial")
    r_fonts.set(qn("w:hAnsi"), "Arial")
    r_fonts.set(qn("w:cs"), "Arial")
    if color is not None:
        run.font.color.rgb = color


def fill_cell(cell, text: str, *, bold=False, size=10, center=True, fill=None, color=None) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    set_run_font(run, bold=bold, size=size, color=color)
    if fill:
        set_cell_shading(cell, fill)
    # вертикальное выравнивание по центру
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    v_align = OxmlElement("w:vAlign")
    v_align.set(qn("w:val"), "center")
    tc_pr.append(v_align)


def assignment_for_day(day_index: int) -> tuple[str | None, str | None]:
    """Возвращает (лифт_2_2, лифт_5_2) — ФИО/номер или None."""
    week = day_index // 7
    dow = day_index % 7  # 0=Пн … 6=Вс
    on_5_2 = WEEK_5_2[week]

    # 5/2: Пн–Пт работа, Сб–Вс выходной
    lift_5_2 = on_5_2 if dow < 5 else None

    # 2/2: кто из пары работает
    on_2_2 = [w for w in WORKERS if w != on_5_2]
    lift_2_2 = None
    for w in on_2_2:
        pat = PATTERN_2_2[w][day_index]
        if pat == "W":
            lift_2_2 = w
            break
    return lift_2_2, lift_5_2


def worker_mode(day_index: int, worker: str) -> str:
    week = day_index // 7
    if WEEK_5_2[week] == worker:
        dow = day_index % 7
        return "Р" if dow < 5 else "В"
    pat = PATTERN_2_2[worker][day_index]
    if pat == "W":
        return "Р"
    if pat == "O":
        return "В"
    return "—"


def build() -> Path:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(29.7)  # A4 landscape
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("График работы лифтов на 3 недели")
    set_run_font(r, bold=True, size=16)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(
        f"Период: {START.strftime('%d.%m.%Y')} — "
        f"{(START + timedelta(days=20)).strftime('%d.%m.%Y')}  ·  "
        "Лифт А — график 2/2  ·  Лифт Б — график 5/2"
    )
    set_run_font(r, size=11)

    # —— Условия ——
    h = doc.add_paragraph()
    r = h.add_run("Условия (не нарушаются)")
    set_run_font(r, bold=True, size=12)

    rules = [
        "Два лифта: Лифт А — смена 2 рабочих / 2 выходных (2/2); Лифт Б — 5 рабочих / 2 выходных (5/2).",
        "Три работника: №1, №2, №3.",
        "За 3 недели каждый работник отрабатывает 2 недели по графику 2/2 и 1 неделю по графику 5/2.",
        "На каждом лифте каждый день должен быть сотрудник; если все назначенные имеют законный выходной — ячейка оставлена пустой для резолюции.",
    ]
    for t in rules:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(t)
        set_run_font(run, size=10)

    # —— Ротация по неделям ——
    h = doc.add_paragraph()
    r = h.add_run("Ротация по неделям")
    set_run_font(r, bold=True, size=12)

    rot = doc.add_table(rows=4, cols=4)
    rot.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Неделя", "Лифт Б (5/2)", "Лифт А (2/2)", "Примечание"]
    for i, text in enumerate(headers):
        fill_cell(rot.rows[0].cells[i], text, bold=True, fill="1F4E79", color=RGBColor(255, 255, 255))

    notes = [
        "№2 и №3 в противофазе 2/2 — полное покрытие Лифта А",
        "№1 и №3 в противофазе 2/2 — полное покрытие Лифта А",
        "№1 и №2 в противофазе 2/2 — полное покрытие Лифта А",
    ]
    for w in range(3):
        week_start = START + timedelta(days=w * 7)
        week_end = week_start + timedelta(days=6)
        on_b = WEEK_5_2[w]
        on_a = ", ".join(x for x in WORKERS if x != on_b)
        fill_cell(
            rot.rows[w + 1].cells[0],
            f"Неделя {w + 1}\n{week_start.strftime('%d.%m')}–{week_end.strftime('%d.%m')}",
            bold=True,
            fill="D6E3F0",
        )
        fill_cell(rot.rows[w + 1].cells[1], f"Работник {on_b}", fill="FFF2CC")
        fill_cell(rot.rows[w + 1].cells[2], f"Работники {on_a}", fill="E2EFDA")
        fill_cell(rot.rows[w + 1].cells[3], notes[w], size=9)

    doc.add_paragraph()

    # —— Основная таблица: кто на каком лифте каждый день ——
    h = doc.add_paragraph()
    r = h.add_run("Ежедневная расстановка по лифтам")
    set_run_font(r, bold=True, size=12)

    legend = doc.add_paragraph()
    r = legend.add_run(
        "В ячейке — номер работника. Пустая ячейка = законный выходной по графику 5/2 (Сб–Вс), "
        "оставлена для вашей резолюции."
    )
    set_run_font(r, size=9, color=RGBColor(80, 80, 80))

    main = doc.add_table(rows=5, cols=22)
    main.alignment = WD_TABLE_ALIGNMENT.CENTER
    main.autofit = True

    fill_cell(main.rows[0].cells[0], "", fill="1F4E79")
    fill_cell(main.rows[1].cells[0], "Дата", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    fill_cell(main.rows[2].cells[0], "День", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    fill_cell(main.rows[3].cells[0], "Лифт А\n(2/2)", bold=True, fill="548235", color=RGBColor(255, 255, 255), size=8)
    fill_cell(main.rows[4].cells[0], "Лифт Б\n(5/2)", bold=True, fill="BF8F00", color=RGBColor(255, 255, 255), size=8)

    # week headers spanning conceptually (written in first row)
    for w in range(3):
        # mark week label on first day of week column; others blank week header
        for d in range(7):
            col = 1 + w * 7 + d
            if d == 0:
                fill_cell(
                    main.rows[0].cells[col],
                    f"Неделя {w + 1}",
                    bold=True,
                    fill="2E75B6",
                    color=RGBColor(255, 255, 255),
                    size=8,
                )
            else:
                fill_cell(main.rows[0].cells[col], "", fill="2E75B6")

    for day in range(21):
        col = day + 1
        dt = START + timedelta(days=day)
        dow = day % 7
        weekend = dow >= 5
        date_fill = "F2F2F2" if weekend else "FFFFFF"
        day_fill = "D9D9D9" if weekend else "FFFFFF"

        fill_cell(main.rows[1].cells[col], dt.strftime("%d.%m"), size=8, fill=date_fill)
        fill_cell(main.rows[2].cells[col], DAYS_RU[dow], bold=True, size=8, fill=day_fill)

        a, b = assignment_for_day(day)
        fill_cell(
            main.rows[3].cells[col],
            a or "",
            bold=True,
            size=10,
            fill="C6EFCE" if a else "FFFFFF",
        )
        if b:
            fill_cell(main.rows[4].cells[col], b, bold=True, size=10, fill="FFE699")
        else:
            # пусто для резолюции
            fill_cell(main.rows[4].cells[col], "", size=10, fill="FCE4D6")

    # ширины
    for row in main.rows:
        row.cells[0].width = Cm(2.4)
        for c in range(1, 22):
            row.cells[c].width = Cm(1.15)

    doc.add_paragraph()

    # —— Табель по работникам ——
    h = doc.add_paragraph()
    r = h.add_run("Табель работников (Р — работа, В — выходной)")
    set_run_font(r, bold=True, size=12)

    tab = doc.add_table(rows=7, cols=22)
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER

    fill_cell(tab.rows[0].cells[0], "Работник", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    fill_cell(tab.rows[1].cells[0], "Режим недели", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)

    for day in range(21):
        col = day + 1
        dt = START + timedelta(days=day)
        fill_cell(tab.rows[0].cells[col], dt.strftime("%d.%m"), size=7, fill="D6E3F0")
        week = day // 7
        # show mode only on Monday of each week block conceptually per worker rows below
        fill_cell(tab.rows[1].cells[col], DAYS_RU[day % 7], size=7, fill="D6E3F0")

    # worker rows: 2,3,4 for status; we'll use rows 2-4 for Р/В and add mode info in first cells
    # Redesign: rows = header dates, then for each worker: mode row + status... 
    # Simpler: 1 header + 3 workers
    # Rebuild a cleaner timesheet table instead.
    # Actually current tab has 7 rows - let me use:
    # row0 dates, row1 dow, row2 W1, row3 W2, row4 W3, row5 empty, row6 summary
    # Clear and refill properly.

    # Re-do timesheet more cleanly as a separate table
    # Remove incomplete approach - create new table
    # We'll abandon `tab` content partially - better create fresh.

    # Fix: rewrite document section by building timesheet as new table after removing confusion.
    # Easiest: don't use `tab` above - I already created it. Let me fill rows 2-4 with workers.

    for wi, worker in enumerate(WORKERS):
        row = tab.rows[2 + wi]
        # first cell: worker + week modes
        modes = []
        for w in range(3):
            modes.append("5/2" if WEEK_5_2[w] == worker else "2/2")
        fill_cell(
            row.cells[0],
            f"{worker}\n({' / '.join(modes)})",
            bold=True,
            size=8,
            fill="D6E3F0",
        )
        for day in range(21):
            col = day + 1
            status = worker_mode(day, worker)
            if status == "Р":
                # where working?
                a, b = assignment_for_day(day)
                place = "А" if a == worker else "Б"
                fill_cell(row.cells[col], f"Р/{place}", bold=True, size=8, fill="C6EFCE")
            elif status == "В":
                fill_cell(row.cells[col], "В", size=8, fill="DDEBF7")
            else:
                fill_cell(row.cells[col], "—", size=8)

    # hide unused rows 5-6 by leaving empty or delete - python-docx can't easily delete;
    # fill with summary
    fill_cell(tab.rows[5].cells[0], "Итого Р", bold=True, size=8, fill="F2F2F2")
    for day in range(21):
        a, b = assignment_for_day(day)
        count = (1 if a else 0) + (1 if b else 0)
        fill_cell(tab.rows[5].cells[col := day + 1], str(count), size=8, fill="F2F2F2")

    fill_cell(tab.rows[6].cells[0], "Пустые\nслоты", bold=True, size=8, fill="FCE4D6")
    for day in range(21):
        a, b = assignment_for_day(day)
        empty = []
        if not a:
            empty.append("А")
        if not b:
            empty.append("Б")
        fill_cell(
            tab.rows[6].cells[day + 1],
            ",".join(empty) if empty else "—",
            size=7,
            fill="FCE4D6" if empty else "FFFFFF",
        )

    # remove the unused "Режим недели" row meaning - row1 is DOW which is fine.
    # Rename row1 first cell
    fill_cell(tab.rows[1].cells[0], "День", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)

    doc.add_paragraph()

    # —— Контроль выполнения условий ——
    h = doc.add_paragraph()
    r = h.add_run("Контроль условий")
    set_run_font(r, bold=True, size=12)

    # count work days per worker on each mode
    check = doc.add_table(rows=4, cols=5)
    check.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, t in enumerate(["Работник", "Недель 2/2", "Недель 5/2", "Смен на Лифте А", "Смен на Лифте Б"]):
        fill_cell(check.rows[0].cells[i], t, bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=9)

    for wi, worker in enumerate(WORKERS):
        weeks_22 = sum(1 for w in range(3) if WEEK_5_2[w] != worker)
        weeks_52 = sum(1 for w in range(3) if WEEK_5_2[w] == worker)
        shifts_a = sum(1 for d in range(21) if assignment_for_day(d)[0] == worker)
        shifts_b = sum(1 for d in range(21) if assignment_for_day(d)[1] == worker)
        fill_cell(check.rows[wi + 1].cells[0], f"Работник {worker}", bold=True, size=9)
        fill_cell(check.rows[wi + 1].cells[1], str(weeks_22), size=9, fill="E2EFDA")
        fill_cell(check.rows[wi + 1].cells[2], str(weeks_52), size=9, fill="FFF2CC")
        fill_cell(check.rows[wi + 1].cells[3], str(shifts_a), size=9)
        fill_cell(check.rows[wi + 1].cells[4], str(shifts_b), size=9)

    doc.add_paragraph()
    note = doc.add_paragraph()
    r = note.add_run("Примечания для резолюции")
    set_run_font(r, bold=True, size=12)

    bullets = [
        "Лифт А (2/2) закрыт все 21 день: двое сотрудников работают в противофазе.",
        "Лифт Б (5/2) закрыт в Пн–Пт; в Сб и Вс ячейки пустые (законный выходной по 5/2) — требуется ваша резолюция "
        "(подмена, дежурство или вывод лифта с обслуживания).",
        "Паттерн 2/2 непрерывный при переходе недель: WW–ОО–WW–ОО…",
        "Дата начала условная (пн 16.11.2026); при сдвиге старта сохраните ту же структуру недель и противофазу.",
    ]
    for t in bullets:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(t)
        set_run_font(run, size=10)

    doc.save(OUT)
    return OUT


def verify() -> None:
    """Проверка инвариантов."""
    # каждый день лифт А занят
    for d in range(21):
        a, b = assignment_for_day(d)
        assert a is not None, f"Лифт А пуст в день {d}"
        if d % 7 < 5:
            assert b is not None, f"Лифт Б пуст в будний день {d}"
        else:
            assert b is None, f"Лифт Б должен быть пуст в выходной 5/2, день {d}"

    # никто не работает на двух лифтах в один день
    for d in range(21):
        a, b = assignment_for_day(d)
        if a and b:
            assert a != b, f"Один человек на двух лифтах в день {d}"

    # недели 2/2 и 5/2
    for worker in WORKERS:
        w52 = sum(1 for w in range(3) if WEEK_5_2[w] == worker)
        w22 = 3 - w52
        assert w52 == 1 and w22 == 2, worker

    # в день работы по 2/2 человек действительно на лифте А
    for d in range(21):
        for worker in WORKERS:
            status = worker_mode(d, worker)
            a, b = assignment_for_day(d)
            if status == "Р":
                assert worker in (a, b), f"{worker} Р но не на лифте, день {d}"
            else:
                assert worker not in (a, b), f"{worker} выходной но на лифте, день {d}"


if __name__ == "__main__":
    verify()
    path = build()
    print(f"OK: {path}")
