#!/usr/bin/env python3
"""Генерация графика работы лифтов (2/2 и 5/2) на 3 недели.

Цикл из 21 дня стыкуется сам с собой: после 3-й недели можно снова
начинать 1-ю без нарушения лимитов подряд.
"""

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

START = date(2026, 11, 16)
DAYS_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
WORKERS = ("№1", "№2", "№3")

MAX_CONSEC_ANY = 5
MAX_CONSEC_WITH_2_2 = 3

# Кто на 5/2 в каждую неделю (остальные двое — на 2/2)
WEEK_5_2 = {0: "№1", 1: "№2", 2: "№3"}

# Паттерны 2/2 на Лифте А (W/O), длина 7.
# Пара в противофазе: WOOWWOO ↔ OWWOOWW (классический 2/2 на 7-дневном окне).
# Перед выходом на свою неделю 5/2 у сотрудника Вс выходной — в т.ч. на стыке циклов.
PATTERN_2_2_WEEK = {
    # неделя 0: №2 / №3
    (0, "№2"): "WOOWWOO",
    (0, "№3"): "OWWOOWW",
    # неделя 1: №1 / №3
    (1, "№1"): "OWWOOWW",
    (1, "№3"): "WOOWWOO",
    # неделя 2: №1 / №2
    (2, "№1"): "WOOWWOO",
    (2, "№2"): "OWWOOWW",
}

# Подмены на Лифт Б в Сб/Вс (основной 5/2 в Сб не выходит — уже 5 дней подряд).
# Вс перед чьей-либо новой 5/2 этот человек свободен.
WEEKEND_FILL_B = {
    5: "№2",   # Сб нед.1
    6: "№1",   # Вс нед.1
    12: "№3",  # Сб нед.2
    13: "№2",  # Вс нед.2
    19: "№1",  # Сб нед.3
    20: "№3",  # Вс нед.3 — №1 свободен перед своей 5/2 в следующем цикле
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
    week = day_index // 7
    dow = day_index % 7
    on_5_2 = WEEK_5_2[week]
    lift_b = on_5_2 if dow < 5 else None
    lift_a = None
    for w in WORKERS:
        if w == on_5_2:
            continue
        if PATTERN_2_2_WEEK[(week, w)][dow] == "W":
            lift_a = w
            break
    return lift_a, lift_b


def assignment_for_day(day_index: int) -> tuple[str | None, str | None]:
    lift_a, lift_b = base_assignment(day_index % 21)
    if lift_b is None:
        lift_b = WEEKEND_FILL_B[day_index % 21]
    return lift_a, lift_b


def place_of(day_index: int, worker: str) -> str | None:
    a, b = assignment_for_day(day_index)
    if a == worker:
        return "A"
    if b == worker:
        return "B"
    return None


def circular_streaks(worker: str, cycles: int = 1) -> list[dict]:
    """Серии рабочих дней по кругу (cycles полных 21-дневных циклов)."""
    n = 21 * cycles
    work = [place_of(d, worker) is not None for d in range(n)]
    on_a = [place_of(d, worker) == "A" for d in range(n)]
    if all(work):
        return [{"start": 0, "end": n - 1, "len": n, "has_a": any(on_a)}]
    if not any(work):
        return []
    start = next(i for i in range(n) if not work[i])
    streaks: list[dict] = []
    i = 0
    while i < n:
        if not work[(start + i) % n]:
            i += 1
            continue
        length = 0
        has_a = False
        s0 = (start + i) % n
        while i < n and work[(start + i) % n]:
            if on_a[(start + i) % n]:
                has_a = True
            length += 1
            i += 1
        streaks.append({"start": s0, "end": (s0 + length - 1) % n, "len": length, "has_a": has_a})
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
    set_run_font(title.add_run("График работы лифтов на 3 недели (пролонгируемый цикл)"), bold=True, size=16)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(
        sub.add_run(
            f"Период-образец: {START.strftime('%d.%m.%Y')} — "
            f"{(START + timedelta(days=20)).strftime('%d.%m.%Y')}  ·  "
            "Лифт А — 2/2  ·  Лифт Б — 5/2  ·  далее цикл повторяется"
        ),
        size=11,
    )

    h = doc.add_paragraph()
    set_run_font(h.add_run("Условия (не нарушаются)"), bold=True, size=12)

    rules = [
        "Два лифта: Лифт А — 2/2; Лифт Б — 5/2.",
        "Три работника: №1, №2, №3.",
        "За каждые 3 недели каждый работник: 2 недели 2/2 и 1 неделю 5/2.",
        "На каждом лифте каждый день есть сотрудник.",
        f"Никто не работает больше {MAX_CONSEC_ANY} дней подряд.",
        f"Если в серии есть смена на Лифте А (2/2) — серия не длиннее {MAX_CONSEC_WITH_2_2} дней "
        "(в т.ч. при переходе А↔Б).",
        "Цикл из 21 дня стыкуется с следующим таким же циклом: перед выходом на свою неделю 5/2 "
        "у сотрудника всегда выходной в воскресенье (включая стык 3-я→1-я неделя).",
    ]
    for t in rules:
        p = doc.add_paragraph(style="List Bullet")
        set_run_font(p.add_run(t), size=10)

    h = doc.add_paragraph()
    set_run_font(h.add_run("Ротация по неделям"), bold=True, size=12)

    rot = doc.add_table(rows=4, cols=4)
    rot.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, text in enumerate(["Неделя", "Лифт Б (5/2)", "Лифт А (2/2)", "Подмена Б в Сб–Вс"]):
        fill_cell(rot.rows[0].cells[i], text, bold=True, fill="1F4E79", color=RGBColor(255, 255, 255))

    weekend_notes = [
        "Сб: №2 · Вс: №1",
        "Сб: №3 · Вс: №2",
        "Сб: №1 · Вс: №3 (№1 свободен перед следующей 5/2)",
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
    set_run_font(h.add_run("Ежедневная расстановка по лифтам"), bold=True, size=12)
    legend = doc.add_paragraph()
    set_run_font(
        legend.add_run(
            "Оранжевая заливка на Лифте Б в Сб–Вс — подмена с выходного. "
            "После Вс 3-й недели цикл можно начинать снова с Пн 1-й недели."
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
        fill = "F4B183" if day in WEEKEND_FILL_B else "FFE699"
        fill_cell(main.rows[4].cells[col], b or "", bold=True, size=10, fill=fill if b else "FCE4D6")

    for row in main.rows:
        row.cells[0].width = Cm(2.4)
        for c in range(1, 22):
            row.cells[c].width = Cm(1.15)

    for start in (1, 8, 15):
        main.cell(0, start).merge(main.cell(0, start + 6))

    doc.add_paragraph()

    # Стык циклов
    h = doc.add_paragraph()
    set_run_font(h.add_run("Стык циклов (3-я неделя → следующая 1-я)"), bold=True, size=12)
    seam = doc.add_paragraph()
    set_run_font(
        seam.add_run(
            "Вс 3-й недели: №1 выходной → Пн 1-й недели следующего цикла №1 выходит на 5/2. "
            "Разрыв в 1 день не даёт серии длиннее 5. Аналогично: перед 5/2 №2 свободен в Вс 1-й недели; "
            "перед 5/2 №3 — в Вс 2-й недели."
        ),
        size=10,
    )

    doc.add_paragraph()

    h = doc.add_paragraph()
    set_run_font(h.add_run("Табель работников (Р — работа, В — выходной)"), bold=True, size=12)

    tab = doc.add_table(rows=6, cols=22)
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    fill_cell(tab.rows[0].cells[0], "Работник", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    fill_cell(tab.rows[1].cells[0], "День", bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)
    for day in range(21):
        col = day + 1
        fill_cell(tab.rows[0].cells[col], (START + timedelta(days=day)).strftime("%d.%m"), size=7, fill="D6E3F0")
        fill_cell(tab.rows[1].cells[col], DAYS_RU[day % 7], size=7, fill="D6E3F0")

    for wi, worker in enumerate(WORKERS):
        row = tab.rows[2 + wi]
        modes = ["5/2" if WEEK_5_2[w] == worker else "2/2" for w in range(3)]
        fill_cell(row.cells[0], f"{worker}\n({' / '.join(modes)})", bold=True, size=8, fill="D6E3F0")
        for day in range(21):
            place = place_of(day, worker)
            if place:
                fill = "F4B183" if (place == "B" and day in WEEKEND_FILL_B) else "C6EFCE"
                fill_cell(row.cells[day + 1], f"Р/{'А' if place == 'A' else 'Б'}", bold=True, size=8, fill=fill)
            else:
                fill_cell(row.cells[day + 1], "В", size=8, fill="DDEBF7")

    fill_cell(tab.rows[5].cells[0], "Итого Р\nна лифтах", bold=True, size=8, fill="F2F2F2")
    for day in range(21):
        a, b = assignment_for_day(day)
        count = (1 if a else 0) + (1 if b else 0)
        fill_cell(tab.rows[5].cells[day + 1], str(count), size=8, fill="C6EFCE" if count == 2 else "FCE4D6")

    doc.add_paragraph()

    h = doc.add_paragraph()
    set_run_font(h.add_run("Контроль условий (в т.ч. на стыке циклов)"), bold=True, size=12)

    check = doc.add_table(rows=4, cols=6)
    check.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = [
        "Работник",
        "Недель 2/2",
        "Недель 5/2",
        "Смен А",
        "Смен Б",
        f"Макс. подряд за 2 цикла\n(с А ≤{MAX_CONSEC_WITH_2_2} / всего ≤{MAX_CONSEC_ANY})",
    ]
    for i, t in enumerate(headers):
        fill_cell(check.rows[0].cells[i], t, bold=True, fill="1F4E79", color=RGBColor(255, 255, 255), size=8)

    for wi, worker in enumerate(WORKERS):
        weeks_22 = sum(1 for w in range(3) if WEEK_5_2[w] != worker)
        shifts_a = sum(1 for d in range(21) if place_of(d, worker) == "A")
        shifts_b = sum(1 for d in range(21) if place_of(d, worker) == "B")
        streaks = circular_streaks(worker, cycles=2)
        max_any = max((s["len"] for s in streaks), default=0)
        max_with_a = max((s["len"] for s in streaks if s["has_a"]), default=0)
        fill_cell(check.rows[wi + 1].cells[0], f"Работник {worker}", bold=True, size=9)
        fill_cell(check.rows[wi + 1].cells[1], str(weeks_22), size=9, fill="E2EFDA")
        fill_cell(check.rows[wi + 1].cells[2], "1", size=9, fill="FFF2CC")
        fill_cell(check.rows[wi + 1].cells[3], str(shifts_a), size=9)
        fill_cell(check.rows[wi + 1].cells[4], str(shifts_b), size=9)
        fill_cell(check.rows[wi + 1].cells[5], f"с А: {max_with_a} · всего: {max_any}", size=8, fill="C6EFCE")

    doc.add_paragraph()
    note = doc.add_paragraph()
    set_run_font(note.add_run("Примечания"), bold=True, size=12)
    bullets = [
        "Лифт А: пара 2/2 в противофазе WOOWWOO / OWWOOWW.",
        "Лифт Б Пн–Пт — основной сотрудник недели 5/2; Сб–Вс — подмена.",
        "Пролонгация: копируйте блок из 3 недель подряд; лимиты подряд сохраняются на стыках.",
        "Дата начала условная (пн 16.11.2026); структура цикла не зависит от даты.",
    ]
    for t in bullets:
        p = doc.add_paragraph(style="List Bullet")
        set_run_font(p.add_run(t), size=10)

    doc.save(OUT)
    return OUT


def verify() -> None:
    for d in range(21):
        a, b = assignment_for_day(d)
        assert a and b and a != b, f"покрытие дня {d}: {a}, {b}"

    for worker in WORKERS:
        assert sum(1 for w in range(3) if WEEK_5_2[w] == worker) == 1

    for d, w in WEEKEND_FILL_B.items():
        assert base_assignment(d)[0] != w, f"подмена {w} занята на А в день {d}"
        assert base_assignment(d)[1] is None

    # перед каждой 5/2 — полный выходной в воскресенье
    for w in range(3):
        who = WEEK_5_2[w]
        sun = (w * 7 - 1) % 21
        assert place_of(sun, who) is None, f"{who} должен быть свободен в день {sun} перед 5/2"

    # лимиты на 1 и 2 циклах (стык)
    for cycles in (1, 2):
        for worker in WORKERS:
            for s in circular_streaks(worker, cycles=cycles):
                assert s["len"] <= MAX_CONSEC_ANY, f"{worker} cycles={cycles}: {s}"
                if s["has_a"]:
                    assert s["len"] <= MAX_CONSEC_WITH_2_2, f"{worker} cycles={cycles} with A: {s}"


if __name__ == "__main__":
    verify()
    path = build()
    print(f"OK: {path}")
    for worker in WORKERS:
        print(f"{worker}: {circular_streaks(worker, cycles=2)}")
