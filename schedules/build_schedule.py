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

MAX_CONSEC_ANY = 5
MAX_CONSEC_WITH_2_2 = 3

# Кто на 5/2 в каждую неделю (остальные двое — на 2/2)
WEEK_5_2 = {0: "№1", 1: "№2", 2: "№3"}

# Паттерны 2/2 (W/O) на 21 день — противофаза, непрерывный цикл
PATTERN_2_2 = {
    "№2": list("WWOOWWO") + [None] * 7 + list("WWOOWWO"),
    "№3": list("OOWWOOW") + list("WOOWWOO") + [None] * 7,
    "№1": [None] * 7 + list("OWWOOWW") + list("OOWWOOW"),
}

# Подмены на Лифт Б в Сб/Вс (базовый 5/2 даёт выходной).
# Подобраны так, чтобы:
# - никто не работал >5 дней подряд;
# - при участии Лифта А (2/2) в серии — не >3 дней подряд (в т.ч. при переходе А↔Б).
WEEKEND_FILL_B = {
    5: "№3",   # Сб нед.1
    6: "№1",   # Вс нед.1
    12: "№3",  # Сб нед.2
    13: "№2",  # Вс нед.2
    19: "№1",  # Сб нед.3
    20: "№2",  # Вс нед.3
}


def set_cell_shading(cell, hex_color: str) -> None:
    tc = cell._tc
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
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    v_align = OxmlElement("w:vAlign")
    v_align.set(qn("w:val"), "center")
    tc_pr.append(v_align)


def base_assignment(day_index: int) -> tuple[str | None, str | None]:
    """Базовая расстановка без подмен выходных 5/2."""
    week = day_index // 7
    dow = day_index % 7
    on_5_2 = WEEK_5_2[week]
    lift_5_2 = on_5_2 if dow < 5 else None
    lift_2_2 = None
    for w in (x for x in WORKERS if x != on_5_2):
        if PATTERN_2_2[w][day_index] == "W":
            lift_2_2 = w
            break
    return lift_2_2, lift_5_2


def assignment_for_day(day_index: int) -> tuple[str | None, str | None]:
    """Возвращает (лифт_А_2/2, лифт_Б_5/2) с учётом подмен на выходных."""
    lift_a, lift_b = base_assignment(day_index)
    if lift_b is None and day_index in WEEKEND_FILL_B:
        lift_b = WEEKEND_FILL_B[day_index]
    return lift_a, lift_b


def place_of(day_index: int, worker: str) -> str | None:
    a, b = assignment_for_day(day_index)
    if a == worker:
        return "A"
    if b == worker:
        return "B"
    return None


def worker_status(day_index: int, worker: str) -> str:
    """Р / В с учётом подмен."""
    return "Р" if place_of(day_index, worker) else "В"


def consecutive_streaks(worker: str) -> list[dict]:
    """Серии рабочих дней: длина, есть ли день на Лифте А (2/2)."""
    streaks: list[dict] = []
    i = 0
    while i < 21:
        if not place_of(i, worker):
            i += 1
            continue
        start = i
        has_a = False
        while i < 21 and place_of(i, worker):
            if place_of(i, worker) == "A":
                has_a = True
            i += 1
        streaks.append({"start": start, "end": i - 1, "len": i - start, "has_a": has_a})
    return streaks


def build() -> Path:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(29.7)
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

    h = doc.add_paragraph()
    r = h.add_run("Условия (не нарушаются)")
    set_run_font(r, bold=True, size=12)

    rules = [
        "Два лифта: Лифт А — смена 2 рабочих / 2 выходных (2/2); Лифт Б — 5 рабочих / 2 выходных (5/2).",
        "Три работника: №1, №2, №3.",
        "За 3 недели каждый работник отрабатывает 2 недели по графику 2/2 и 1 неделю по графику 5/2.",
        "На каждом лифте каждый день есть сотрудник.",
        f"Никто не работает больше {MAX_CONSEC_ANY} дней подряд.",
        f"Если в серии рабочих дней есть смена на Лифте А (2/2) — серия не длиннее {MAX_CONSEC_WITH_2_2} дней "
        "(считается и при переходе между лифтами А↔Б).",
        "Сб–Вс на Лифте Б закрыты подменой (сотрудник с законного выходного), без нарушения лимитов подряд.",
    ]
    for t in rules:
        p = doc.add_paragraph(style="List Bullet")
        set_run_font(p.add_run(t), size=10)

    h = doc.add_paragraph()
    r = h.add_run("Ротация по неделям")
    set_run_font(r, bold=True, size=12)

    rot = doc.add_table(rows=4, cols=4)
    rot.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, text in enumerate(["Неделя", "Лифт Б (5/2)", "Лифт А (2/2)", "Подмена Б в Сб–Вс"]):
        fill_cell(rot.rows[0].cells[i], text, bold=True, fill="1F4E79", color=RGBColor(255, 255, 255))

    weekend_notes = [
        "Сб: №3 · Вс: №1",
        "Сб: №3 · Вс: №2",
        "Сб: №1 · Вс: №2",
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
        fill_cell(rot.rows[w + 1].cells[3], weekend_notes[w], size=9, fill="FCE4D6")

    doc.add_paragraph()

    h = doc.add_paragraph()
    r = h.add_run("Ежедневная расстановка по лифтам")
    set_run_font(r, bold=True, size=12)

    legend = doc.add_paragraph()
    set_run_font(
        legend.add_run(
            "В ячейке — номер работника. Оранжевая заливка на Лифте Б в Сб–Вс — подмена "
            "(сотрудник выходит с выходного, чтобы закрыть лифт)."
        ),
        size=9,
        color=RGBColor(80, 80, 80),
    )

    main = doc.add_table(rows=5, cols=22)
    main.alignment = WD_TABLE_ALIGNMENT.CENTER

    fill_cell(main.rows[0].cells[0], "", fill="1F4E79")
    fill_cell(main.rows[1].cells[0], "Дата", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    fill_cell(main.rows[2].cells[0], "День", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    fill_cell(main.rows[3].cells[0], "Лифт А\n(2/2)", bold=True, fill="548235", color=RGBColor(255, 255, 255), size=8)
    fill_cell(main.rows[4].cells[0], "Лифт Б\n(5/2)", bold=True, fill="BF8F00", color=RGBColor(255, 255, 255), size=8)

    for w in range(3):
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
        fill_cell(main.rows[1].cells[col], dt.strftime("%d.%m"), size=8, fill="F2F2F2" if weekend else "FFFFFF")
        fill_cell(main.rows[2].cells[col], DAYS_RU[dow], bold=True, size=8, fill="D9D9D9" if weekend else "FFFFFF")

        a, b = assignment_for_day(day)
        fill_cell(main.rows[3].cells[col], a or "", bold=True, size=10, fill="C6EFCE" if a else "FFFFFF")
        if b:
            # подмена на выходных — другой оттенок
            fill = "F4B183" if day in WEEKEND_FILL_B else "FFE699"
            fill_cell(main.rows[4].cells[col], b, bold=True, size=10, fill=fill)
        else:
            fill_cell(main.rows[4].cells[col], "", size=10, fill="FCE4D6")

    for row in main.rows:
        row.cells[0].width = Cm(2.4)
        for c in range(1, 22):
            row.cells[c].width = Cm(1.15)

    # объединение заголовков недель
    for start in (1, 8, 15):
        main.cell(0, start).merge(main.cell(0, start + 6))

    doc.add_paragraph()

    h = doc.add_paragraph()
    r = h.add_run("Табель работников (Р — работа, В — выходной)")
    set_run_font(r, bold=True, size=12)

    tab = doc.add_table(rows=6, cols=22)
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER

    fill_cell(tab.rows[0].cells[0], "Работник", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    fill_cell(tab.rows[1].cells[0], "День", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    for day in range(21):
        col = day + 1
        dt = START + timedelta(days=day)
        fill_cell(tab.rows[0].cells[col], dt.strftime("%d.%m"), size=7, fill="D6E3F0")
        fill_cell(tab.rows[1].cells[col], DAYS_RU[day % 7], size=7, fill="D6E3F0")

    for wi, worker in enumerate(WORKERS):
        row = tab.rows[2 + wi]
        modes = ["5/2" if WEEK_5_2[w] == worker else "2/2" for w in range(3)]
        fill_cell(row.cells[0], f"{worker}\n({' / '.join(modes)})", bold=True, size=8, fill="D6E3F0")
        for day in range(21):
            col = day + 1
            place = place_of(day, worker)
            if place:
                label = f"Р/{'А' if place == 'A' else 'Б'}"
                # подмена — оранжевый
                fill = "F4B183" if (place == "B" and day in WEEKEND_FILL_B) else "C6EFCE"
                fill_cell(row.cells[col], label, bold=True, size=8, fill=fill)
            else:
                fill_cell(row.cells[col], "В", size=8, fill="DDEBF7")

    fill_cell(tab.rows[5].cells[0], "Итого Р\nна лифтах", bold=True, size=8, fill="F2F2F2")
    for day in range(21):
        a, b = assignment_for_day(day)
        count = (1 if a else 0) + (1 if b else 0)
        fill_cell(tab.rows[5].cells[day + 1], str(count), size=8, fill="C6EFCE" if count == 2 else "FCE4D6")

    doc.add_paragraph()

    h = doc.add_paragraph()
    r = h.add_run("Контроль условий")
    set_run_font(r, bold=True, size=12)

    check = doc.add_table(rows=4, cols=6)
    check.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = [
        "Работник",
        "Недель 2/2",
        "Недель 5/2",
        "Смен А",
        "Смен Б",
        f"Макс. подряд\n(с А ≤{MAX_CONSEC_WITH_2_2} / всего ≤{MAX_CONSEC_ANY})",
    ]
    for i, t in enumerate(headers):
        fill_cell(check.rows[0].cells[i], t, bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)

    for wi, worker in enumerate(WORKERS):
        weeks_22 = sum(1 for w in range(3) if WEEK_5_2[w] != worker)
        weeks_52 = 3 - weeks_22
        shifts_a = sum(1 for d in range(21) if place_of(d, worker) == "A")
        shifts_b = sum(1 for d in range(21) if place_of(d, worker) == "B")
        streaks = consecutive_streaks(worker)
        max_any = max((s["len"] for s in streaks), default=0)
        max_with_a = max((s["len"] for s in streaks if s["has_a"]), default=0)
        fill_cell(check.rows[wi + 1].cells[0], f"Работник {worker}", bold=True, size=9)
        fill_cell(check.rows[wi + 1].cells[1], str(weeks_22), size=9, fill="E2EFDA")
        fill_cell(check.rows[wi + 1].cells[2], str(weeks_52), size=9, fill="FFF2CC")
        fill_cell(check.rows[wi + 1].cells[3], str(shifts_a), size=9)
        fill_cell(check.rows[wi + 1].cells[4], str(shifts_b), size=9)
        fill_cell(
            check.rows[wi + 1].cells[5],
            f"с А: {max_with_a} · всего: {max_any}",
            size=8,
            fill="C6EFCE",
        )

    doc.add_paragraph()
    note = doc.add_paragraph()
    set_run_font(note.add_run("Примечания"), bold=True, size=12)

    bullets = [
        "Лифт А закрыт все 21 день противофазой пары 2/2.",
        "Лифт Б в Пн–Пт — основной сотрудник недели 5/2; в Сб–Вс — подмена с выходного "
        "(основной 5/2 не может выйти: уже 5 дней подряд).",
        "Подмены: Сб нед.1 №3, Вс нед.1 №1; Сб нед.2 №3, Вс нед.2 №2; Сб нед.3 №1, Вс нед.3 №2.",
        "Все серии с участием Лифта А ≤ 3 дня; чистые серии 5/2 = ровно 5 дней (Пн–Пт).",
        "Дата начала условная (пн 16.11.2026); структура недель и противофаза сохраняются при сдвиге.",
    ]
    for t in bullets:
        p = doc.add_paragraph(style="List Bullet")
        set_run_font(p.add_run(t), size=10)

    doc.save(OUT)
    return OUT


def verify() -> None:
    # покрытие обоих лифтов каждый день
    for d in range(21):
        a, b = assignment_for_day(d)
        assert a is not None, f"Лифт А пуст в день {d}"
        assert b is not None, f"Лифт Б пуст в день {d}"
        assert a != b, f"Один человек на двух лифтах в день {d}"

    # недели 2/2 и 5/2 (по базовому назначению)
    for worker in WORKERS:
        w52 = sum(1 for w in range(3) if WEEK_5_2[w] == worker)
        assert w52 == 1 and (3 - w52) == 2, worker

    # подмена только в свой выходной (не занят на А)
    for d, w in WEEKEND_FILL_B.items():
        a, _ = base_assignment(d)
        assert a != w, f"Подмена {w} в день {d}, но уже на Лифте А"
        assert base_assignment(d)[1] is None, f"День {d} не выходной 5/2"

    # лимиты подряд
    for worker in WORKERS:
        for s in consecutive_streaks(worker):
            assert s["len"] <= MAX_CONSEC_ANY, f"{worker}: серия {s['len']} > {MAX_CONSEC_ANY} (дни {s['start']}-{s['end']})"
            if s["has_a"]:
                assert s["len"] <= MAX_CONSEC_WITH_2_2, (
                    f"{worker}: серия с Лифтом А длиной {s['len']} > {MAX_CONSEC_WITH_2_2} "
                    f"(дни {s['start']}-{s['end']})"
                )


if __name__ == "__main__":
    verify()
    path = build()
    print(f"OK: {path}")
    for worker in WORKERS:
        print(f"{worker}: {consecutive_streaks(worker)}")
