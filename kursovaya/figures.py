# -*- coding: utf-8 -*-
"""Схемы и графики для курсовой. Оригинальные учебные рисунки, не сканы учебников."""

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Circle, Arc

OUT = Path(__file__).resolve().parent / "figures"
OUT.mkdir(parents=True, exist_ok=True)

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.unicode_minus"] = False

NAVY = "#1B3A4B"
INK = "#1A1A1A"
BLUE = "#2C5282"
TEAL = "#2B6E6A"
CLAY = "#9C4A2B"
SAND = "#F4EFE6"
MIST = "#E6EEF2"
GREEN = "#2F6B4F"
GOLD = "#8A5A12"
LINE = "#2C3E50"


def _save(fig, name):
    path = OUT / name
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def _box(ax, x, y, w, h, text, fc, ec=NAVY, fs=9, weight="regular", tc=INK, radius=0.08):
    ax.add_patch(
        FancyBboxPatch(
            (x, y),
            w,
            h,
            boxstyle=f"round,pad=0.012,rounding_size={radius}",
            facecolor=fc,
            edgecolor=ec,
            linewidth=1.15,
        )
    )
    ax.text(
        x + w / 2,
        y + h / 2,
        text,
        ha="center",
        va="center",
        fontsize=fs,
        color=tc,
        fontweight=weight,
        linespacing=1.25,
    )


def _arrow(ax, x1, y1, x2, y2, color=LINE):
    ax.add_patch(
        FancyArrowPatch(
            (x1, y1),
            (x2, y2),
            arrowstyle="-|>",
            mutation_scale=12,
            linewidth=1.15,
            color=color,
            shrinkA=0,
            shrinkB=0,
        )
    )


def fig_levels():
    fig, ax = plt.subplots(figsize=(11.2, 6.4))
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 9.2)
    ax.axis("off")
    ax.set_title(
        "Как гидрологический факт становится объяснением",
        fontsize=13,
        color=NAVY,
        pad=8,
        fontweight="medium",
        loc="left",
    )
    rows = [
        (7.7, MIST, "1. Факт и номенклатура", "«Лена полноводна летом».\nНазвание реки, место на карте."),
        (6.15, "#E5F0EA", "2. Понятие", "Питание, режим, половодье,\nхарактер течения, бассейн."),
        (4.6, "#F8E8DC", "3. Причинно-следственная связь", "Снеговое питание + течение на север\n+ мерзлота → позднее половодье и заторы."),
        (3.05, "#E4E9F4", "4. Закономерность", "У рек холодного континентального климата\nсо снеговым питанием половодье весенне-летнее."),
        (1.5, "#F3EBD3", "5. Перенос и прогноз", "Колыму ученик объясняет сам,\nпо аналогии, без готового абзаца учебника."),
    ]
    for y, color, title, note in rows:
        _box(ax, 0.4, y, 6.3, 1.25, title, color, fs=10.5, weight="medium")
        _box(ax, 7.3, y, 8.2, 1.25, note, "white", fs=9.5)
    for y in (7.7, 6.15, 4.6, 3.05):
        _arrow(ax, 3.55, y, 3.55, y - 0.28)
    ax.text(
        0.4,
        0.35,
        "Связь — не отдельный «параграф», а ступень между понятием и закономерностью.\nПока ученик остаётся на ступени 1–2, он воспроизводит текст, но не объясняет реку.",
        fontsize=8.5,
        color="#333333",
        va="bottom",
    )
    return _save(fig, "fig01_levels.png")


def fig_method():
    fig, ax = plt.subplots(figsize=(11.2, 5.6))
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 8)
    ax.axis("off")
    steps = [
        ("1", "Что объясняем?", "Один факт,\nне весь параграф"),
        ("2", "Какие факторы\nвообще возможны?", "Климат, рельеф,\nчеловек, мерзлота"),
        ("3", "Что видно\nна картах?", "Климат, физическая\nкарта, сток, течения"),
        ("4", "Какой фактор\nздесь ведущий?", "И какой только\nусиливает его"),
        ("5", "Фраза со «потому что»", "Причина и следствие\nназвать разными словами"),
        ("6", "Проверка переносом", "Новая река или вопрос\n«что будет, если…»"),
    ]
    w, h = 4.6, 2.15
    coords = [(0.4, 5.2), (5.6, 5.2), (10.8, 5.2), (0.4, 1.5), (5.6, 1.5), (10.8, 1.5)]
    for (n, title, note), (x, y) in zip(steps, coords):
        circ = Circle((x + 0.42, y + h - 0.42), 0.32, facecolor=NAVY, edgecolor=NAVY)
        ax.add_patch(circ)
        ax.text(x + 0.42, y + h - 0.42, n, ha="center", va="center", color="white", fontsize=11, fontweight="bold")
        ax.text(x + 0.9, y + h - 0.48, title, ha="left", va="center", fontsize=10, color=NAVY, fontweight="medium")
        ax.text(x + 0.25, y + 0.55, note, ha="left", va="center", fontsize=9, color=INK)
        ax.add_patch(
            FancyBboxPatch(
                (x, y),
                w,
                h,
                boxstyle="round,pad=0.012,rounding_size=0.08",
                facecolor="#F7F5F1",
                edgecolor=NAVY,
                linewidth=1.0,
                zorder=0,
            )
        )
    _arrow(ax, 5.05, 6.25, 5.5, 6.25)
    _arrow(ax, 10.25, 6.25, 10.7, 6.25)
    _arrow(ax, 5.05, 2.55, 5.5, 2.55)
    _arrow(ax, 10.25, 2.55, 10.7, 2.55)
    ax.annotate(
        "",
        xy=(2.7, 3.75),
        xytext=(2.7, 5.1),
        arrowprops=dict(arrowstyle="-|>", color=CLAY, lw=1.1),
    )
    ax.text(2.85, 4.25, "сначала вместе\nс учителем", fontsize=8, color=CLAY)
    return _save(fig, "fig02_method.png")


def fig_continuity():
    fig, ax = plt.subplots(figsize=(11.2, 6.6))
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 9.4)
    ax.axis("off")
    cols = [
        (0.35, "#E7F0F4", "6 класс", "Опора", "Части гидросферы.\nПитание и режим реки.\nБассейн, водораздел.\nТёплые и холодные течения.\nСвязь «климат бассейна —\nпитание — режим»\nна простом примере."),
        (5.55, "#F8EDE3", "7 класс", "Планетарный масштаб", "Солёность океана и система\nтечений. Несколько климатов\nвнутри одного бассейна.\nТранзитные реки.\nКонтраст материков:\nАмазонка и крики Австралии.\nЛедовый сток Антарктиды."),
        (10.75, "#E7F2EA", "8 класс", "Своя страна", "Бассейны океанов России.\nРежим и характер течения\nразведены по причинам.\nМерзлота, болота, Каспий.\nОпасные явления.\nВодные ресурсы и свой край."),
    ]
    for x, color, title, subtitle, body in cols:
        ax.add_patch(
            FancyBboxPatch(
                (x, 1.15),
                4.85,
                7.5,
                boxstyle="round,pad=0.02,rounding_size=0.1",
                facecolor=color,
                edgecolor=NAVY,
                linewidth=1.15,
            )
        )
        ax.text(x + 2.42, 8.05, title, ha="center", va="center", fontsize=13, color=NAVY, fontweight="bold")
        ax.text(x + 2.42, 7.35, subtitle, ha="center", va="center", fontsize=10, color=CLAY, fontweight="medium")
        ax.plot([x + 0.45, x + 4.4], [6.95, 6.95], color=NAVY, lw=0.6)
        ax.text(x + 0.35, 6.55, body, ha="left", va="top", fontsize=9, color=INK, linespacing=1.35)
    _arrow(ax, 5.25, 4.8, 5.5, 4.8)
    _arrow(ax, 10.45, 4.8, 10.7, 4.8)
    ax.text(
        0.4,
        0.35,
        "Если в 8 классе заново учить, чем половодье отличается от паводка, цепочку по России собрать уже не успеть:\nна тему «Моря России. Внутренние воды и водные ресурсы» программа даёт 6 часов.",
        fontsize=8.4,
        color="#333333",
    )
    return _save(fig, "fig03_continuity.png")


def fig_chains():
    fig, ax = plt.subplots(figsize=(11.4, 7.2))
    ax.set_xlim(0, 16.2)
    ax.set_ylim(0, 10.2)
    ax.axis("off")
    chains = [
        (
            0.25,
            TEAL,
            "Амазонка, 7 класс",
            [
                "Экватор и оба\nполушария",
                "Дожди круглый год,\nпики притоков\nне совпадают",
                "Дождевое питание,\nбез долгой межени",
                "Очень большой\nи ровный сток",
            ],
        ),
        (
            5.55,
            CLAY,
            "Нил, 7 класс",
            [
                "Истоки: экватор\nи Эфиопское\nнагорье",
                "Летние дожди\nна нагорье",
                "Паводок приходит\nв сухой Египет",
                "Плотина сгладила\nход уровня в низовье",
            ],
        ),
        (
            10.85,
            BLUE,
            "Амур, 8 класс",
            [
                "Муссонная смена\nциркуляции",
                "Летне-осенние\nдожди",
                "Дождевое питание,\nзима — межень",
                "Наводнения во\nвторой половине лета",
            ],
        ),
    ]
    for x, color, title, steps in chains:
        ax.text(x + 2.35, 9.55, title, ha="center", fontsize=11, color=color, fontweight="bold")
        y = 7.85
        for i, step in enumerate(steps):
            _box(ax, x, y, 4.7, 1.35, step, "white", ec=color, fs=9)
            if i < len(steps) - 1:
                _arrow(ax, x + 2.35, y, x + 2.35, y - 0.38, color=color)
            y -= 1.85
    ax.text(
        0.25,
        0.25,
        "Три разные цепочки учат не запоминать «тип реки», а видеть, откуда берётся вода и в какой месяц она приходит.\n"
        "Нил специально дан с человеческой поправкой: природная связь и плотина — два слоя, их нельзя склеивать в одну фразу.",
        fontsize=8.3,
        color="#333333",
    )
    return _save(fig, "fig04_chains.png")


def fig_currents():
    fig, ax = plt.subplots(figsize=(11.2, 6.5))
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 9.3)
    ax.axis("off")
    ax.text(0.3, 8.85, "Тропики: правило, которое спрашивает практическая работа", fontsize=11, color=NAVY, fontweight="medium")
    _box(ax, 0.3, 5.15, 7.3, 3.3, "", "#F8E8DC", ec=CLAY)
    _box(ax, 8.3, 5.15, 7.3, 3.3, "", "#E5F0EA", ec=TEAL)
    ax.text(3.95, 8.05, "Западный берег материка", ha="center", fontsize=10.5, color=CLAY, fontweight="bold")
    ax.text(11.95, 8.05, "Восточный берег материка", ha="center", fontsize=10.5, color=TEAL, fontweight="bold")
    ax.text(
        3.95,
        6.55,
        "Холодное течение, часто подъём\nглубинных вод\n→ воздух устойчив, осадков мало\n→ береговые пустыни\n(Атакама, Намиб)",
        ha="center",
        va="center",
        fontsize=9,
        color=INK,
        linespacing=1.35,
    )
    ax.text(
        11.95,
        6.55,
        "Тёплое течение\n→ воздух влажнее, осадков больше\n→ влажные леса доходят\nдо берега\n(восток Южной Америки, Мадагаскар)",
        ha="center",
        va="center",
        fontsize=9,
        color=INK,
        linespacing=1.35,
    )
    ax.text(0.3, 4.55, "И сразу оговорка, иначе правило превращается в заклинание", fontsize=11, color=NAVY, fontweight="medium")
    _box(
        ax,
        0.3,
        0.45,
        15.3,
        3.75,
        "",
        "#F7F3E8",
        ec=GOLD,
    )
    ax.text(
        8.0,
        2.35,
        "Северная Атлантика устроена иначе. К берегам Европы приходит продолжение\n"
        "Гольфстрима — Северо-Атлантическое течение. Поэтому Мурманск не замерзает,\n"
        "а на той же широте у Гудзона зимой совсем другой лёд.\n"
        "Фраза «западный берег всегда холодный» для 7 класса вредна:\n"
        "в 8 классе её придётся ломать на Баренцевом море.",
        ha="center",
        va="center",
        fontsize=9.5,
        color=INK,
        linespacing=1.4,
    )
    return _save(fig, "fig05_currents.png")


def fig_hydrograph():
    fig, ax = plt.subplots(figsize=(11.2, 5.8))
    months = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
    x = list(range(12))
    amazon = [80, 82, 86, 90, 93, 88, 82, 76, 74, 77, 80, 81]
    lena = [8, 6, 5, 7, 28, 100, 72, 42, 30, 20, 13, 9]
    amur = [14, 10, 12, 22, 34, 48, 78, 100, 84, 52, 28, 16]
    ax.plot(x, amazon, color=TEAL, lw=2.2, marker="o", ms=4.5, label="Амазонка — дождевое, без долгой межени")
    ax.plot(x, lena, color=BLUE, lw=2.2, marker="o", ms=4.5, label="Лена — снеговое, пик в начале лета")
    ax.plot(x, amur, color=CLAY, lw=2.2, marker="o", ms=4.5, label="Амур — дождевое, пик во второй половине лета")
    ax.set_xticks(x)
    ax.set_xticklabels(months, fontsize=9)
    ax.set_ylim(0, 118)
    ax.set_ylabel("Относительная водность", fontsize=10)
    ax.set_xlim(-0.2, 11.2)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", color="#D5DDE3", lw=0.7)
    ax.legend(frameon=False, fontsize=8.5, loc="upper left")
    ax.set_title("Качественная схема хода стока в течение года", fontsize=12, color=NAVY, loc="left", pad=10)
    fig.text(
        0.01,
        -0.02,
        "Это не гидрограф конкретного поста и не величины расхода. На уроке шкала нужна только для того,\n"
        "чтобы увидеть месяц максимума. Подписывать её «тыс. м³/с» нельзя: получится ложный график.",
        fontsize=8,
        color="#333333",
    )
    fig.tight_layout()
    return _save(fig, "fig06_hydrograph.png")


def fig_swamp():
    fig, ax = plt.subplots(figsize=(11.2, 6.8))
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 9.6)
    ax.axis("off")
    _box(ax, 5.15, 4.05, 5.7, 1.7, "Сильная\nзаболоченность\nцентра Западной Сибири", "#E7F2EA", ec=GREEN, fs=11, weight="medium")
    causes = [
        (0.3, 7.35, "Плоский рельеф,\nмалые уклоны"),
        (5.7, 7.55, "Осадки больше\nиспарения"),
        (11.0, 7.35, "Глинистые грунты,\nвода стоит"),
        (0.3, 1.15, "Мерзлота на севере\nдержит воду сверху"),
        (11.0, 1.15, "Широкие поймы,\nдолгий разлив Оби"),
    ]
    for x, y, text in causes:
        _box(ax, x, y, 4.5, 1.25, text, "white", ec=NAVY, fs=9)
    _arrow(ax, 3.3, 7.35, 6.3, 5.75)
    _arrow(ax, 7.9, 7.55, 7.9, 5.8)
    _arrow(ax, 12.5, 7.35, 9.8, 5.75)
    _arrow(ax, 3.4, 2.4, 6.2, 4.05)
    _arrow(ax, 12.4, 2.4, 9.9, 4.05)
    ax.annotate(
        "",
        xy=(5.15, 4.5),
        xytext=(4.3, 3.15),
        arrowprops=dict(arrowstyle="-|>", color=CLAY, lw=1.15, connectionstyle="arc3,rad=0.35"),
    )
    ax.text(2.6, 3.35, "болото само\nухудшает сток", fontsize=8, color=CLAY, ha="center")
    ax.text(
        0.35,
        0.25,
        "Ответ «болота, потому что много воды» ничего не объясняет: вода есть и на Валдае.\nВопрос урока другой — почему вода не уходит.",
        fontsize=8.5,
        color="#333333",
    )
    return _save(fig, "fig07_swamp.png")


def fig_formula():
    fig, ax = plt.subplots(figsize=(11.2, 2.9))
    ax.set_xlim(0, 16.4)
    ax.set_ylim(0, 3.4)
    ax.axis("off")
    parts = [
        (0.15, "Факт", "что видно"),
        (3.35, "Причина", "какой фактор"),
        (6.55, "Механизм", "как влияет"),
        (9.75, "Следствие", "режим, беда,\nхозяйство"),
        (12.95, "Граница", "где и когда\nэто верно"),
    ]
    colors = [MIST, "#F8E8DC", "#E4E9F4", "#E7F2EA", "#F3EBD3"]
    for (x, title, note), color in zip(parts, colors):
        _box(ax, x, 0.85, 3.0, 2.1, "", color, ec=NAVY)
        ax.text(x + 1.5, 2.25, title, ha="center", va="center", fontsize=11, color=NAVY, fontweight="bold")
        ax.text(x + 1.5, 1.45, note, ha="center", va="center", fontsize=8.5, color=INK)
    for x in (3.15, 6.35, 9.55, 12.75):
        _arrow(ax, x, 1.9, x + 0.2, 1.9)
    return _save(fig, "fig08_formula.png")


def build_all():
    return {
        "levels": fig_levels(),
        "method": fig_method(),
        "continuity": fig_continuity(),
        "chains": fig_chains(),
        "currents": fig_currents(),
        "hydro": fig_hydrograph(),
        "swamp": fig_swamp(),
        "formula": fig_formula(),
    }


if __name__ == "__main__":
    for key, path in build_all().items():
        print(key, path)
