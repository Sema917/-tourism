# -*- coding: utf-8 -*-
"""Сборка пояснительной записки."""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from engine import (
    FIG, add_mixed, add_page_break, add_picture, add_table, add_text, new_document,
)
from bib import BIB
from text_a import A
from text_b import B
from text_c import C
from expand import expand_chapter1, expand_economy, expand_rest

OUT = "/agent/repos/-tourism/ВКР_Божко_А_П_система_управления_линии_разделки_бревен.docx"

# Страницы оглавления подставляются после первого прогона.
TOC = [
    ("ВВЕДЕНИЕ", "7"),
    ("1 АНАЛИЗ ИМЕЮЩИХСЯ РЕШЕНИЙ", "9"),
    ("2 РАЗРАБОТКА СТРУКТУРНОЙ СХЕМЫ", "45"),
    ("3 РАЗРАБОТКА ФУНКЦИОНАЛЬНОЙ СХЕМЫ", "49"),
    ("4 ПОДБОР ЭЛЕМЕНТНОЙ БАЗЫ", "52"),
    ("4.1 Контроллер серии С7-200 Смарт", "52"),
    ("4.2 Преобразователь частоты", "53"),
    ("4.3 Блок питания", "55"),
    ("4.4 Кнопки управления", "56"),
    ("4.5 Панель оператора", "57"),
    ("4.6 Датчик угловых перемещений", "58"),
    ("4.7 Концевой выключатель", "59"),
    ("4.8 Электродвигатель", "60"),
    ("4.9 Промежуточное реле", "61"),
    ("5 РАЗРАБОТКА ПРИНЦИПИАЛЬНОЙ ЭЛЕКТРИЧЕСКОЙ СХЕМЫ", "65"),
    ("6 РАЗРАБОТКА СХЕМЫ ОБЩЕГО ВИДА УСТРОЙСТВА", "68"),
    ("7 РАЗРАБОТКА АЛГОРИТМА РАБОТЫ", "70"),
    ("8 РАЗРАБОТКА ПРОГРАММНОГО ОБЕСПЕЧЕНИЯ", "76"),
    ("9 ЭКОНОМИЧЕСКОЕ ОБОСНОВАНИЕ ПРОЕКТА", "83"),
    ("ЗАКЛЮЧЕНИЕ", "92"),
    ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "94"),
]

META = {
    "pages": "101",
    "figures": "28",
    "tables": "8",
    "sources": str(len(BIB)),
}


def emit_blocks(doc, blocks):
    for block in blocks:
        kind = block[0]
        if kind == "p":
            add_text(doc, block[1])
        elif kind == "h":
            add_page_break(doc)
            add_text(doc, block[1], center=True, first=0, bold=True, before=0, after=12)
        elif kind == "h2":
            add_text(doc, block[1], center=True, first=0, bold=True, before=12, after=8)
        elif kind == "img":
            add_picture(doc, os.path.join(FIG, block[1]), block[2], block[3])
        elif kind == "table":
            add_table(doc, block[1], block[2], block[3])
        else:
            raise ValueError(kind)


def title_page(doc):
    add_text(doc, "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ", center=True, first=0, bold=True, after=0)
    add_text(doc, "РОССИЙСКОЙ ФЕДЕРАЦИИ", center=True, first=0, bold=True, after=6)
    add_text(doc, "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ", center=True, first=0, after=0)
    add_text(doc, "УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ", center=True, first=0, after=0)
    add_text(doc, "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ", center=True, first=0, bold=True, after=0)
    add_text(doc, "имени И.С. ТУРГЕНЕВА»", center=True, first=0, bold=True, after=12)
    add_text(doc, "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА", center=True, first=0, bold=True, before=12, after=12)
    add_text(doc, "по направлению подготовки 15.03.04 Автоматизация технологических", center=True, first=0, after=0)
    add_text(doc, "процессов и производств", center=True, first=0, after=6)
    add_text(doc, "направленность (профиль): Киберфизические системы", center=True, first=0, after=12)
    add_text(doc, "Студента Божко Антона Петровича, шифр 201498", center=True, first=0, after=0)
    add_text(doc, "Институт приборостроения, автоматизации и информационных технологий", center=True, first=0, after=0)
    add_text(doc, "Кафедра автоматизированных систем управления и кибернетики", center=True, first=0, after=8)
    add_text(doc, "Тема выпускной квалификационной работы", center=True, first=0, after=0)
    add_text(doc, "Модернизация системы управления линии разделки бревен портального типа", center=True, first=0, bold=True, after=12)
    add_text(doc, "Студент ___________ Божко А.П.", align="left", first=0, before=6)
    add_text(doc, "Научный руководитель ___________ Маркин Н.И.", align="left", first=0)
    add_text(doc, "Нормоконтроль ___________ Татарченков Н.В.", align="left", first=0)
    add_text(doc, "И.о. зав. кафедрой ___________ Маркин Н.И.", align="left", first=0, after=12)
    add_text(doc, "Орёл 2024", center=True, first=0, before=12)


def assignment(doc):
    add_page_break(doc)
    add_text(doc, "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ", center=True, first=0, bold=True, after=0, size=12, line=1.0)
    add_text(doc, "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ", center=True, first=0, size=12, line=1.0)
    add_text(doc, "УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ", center=True, first=0, size=12, line=1.0)
    add_text(doc, "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ имени И.С. ТУРГЕНЕВА»", center=True, first=0, bold=True, size=12, line=1.0, after=4)
    add_text(doc, "Институт приборостроения, автоматизации и информационных технологий", center=True, first=0, size=12, line=1.0)
    add_text(doc, "Кафедра автоматизированных систем управления и кибернетики", center=True, first=0, size=12, line=1.0)
    add_text(doc, "Направление подготовки 15.03.04 Автоматизация технологических процессов и производств", center=True, first=0, size=12, line=1.0)
    add_text(doc, "Направленность (профиль) Киберфизические системы", center=True, first=0, size=12, line=1.0, after=6)
    add_text(doc, "УТВЕРЖДАЮ", align="right", first=0, size=12, line=1.0)
    add_text(doc, "И.о. зав. кафедрой ___________ Н.И. Маркин", align="right", first=0, size=12, line=1.0)
    add_text(doc, "«____» ____________ 20____ г.", align="right", first=0, size=12, line=1.0, after=6)
    add_text(doc, "ЗАДАНИЕ", center=True, first=0, bold=True, size=14, line=1.0)
    add_text(doc, "на выполнение выпускной квалификационной работы", center=True, first=0, size=14, line=1.0, after=6)
    add_text(doc, "Студента Божко Антона Петровича, шифр 201498.", first=1.25, size=14, line=1.15)
    add_text(doc, "1. Тема ВКР: модернизация системы управления линии разделки бревен портального типа. Утверждена приказом по университету от «30» октября 2023 г. № 2-3079.", size=14, line=1.15)
    add_text(doc, "2. Срок сдачи студентом законченной работы «24» мая 2024 г.", size=14, line=1.15)
    add_text(doc, "3. Исходные данные к работе: материалы, предоставленные Центром междисциплинарного инжиниринга ОГУ им. И.С. Тургенева; материалы преподавателя; литература, в том числе электронные ресурсы; ГОСТ 2.105-2019.", size=14, line=1.15)
    add_text(doc, "4. Содержание ВКР: анализ объекта модернизации; разработка структурной схемы; разработка функциональной схемы; подбор элементной базы; разработка принципиальной электрической схемы; разработка схемы общего вида; разработка алгоритма работы; разработка программного обеспечения; экономическое обоснование. Заключение. Список использованных источников.", size=14, line=1.15)
    add_text(doc, "5. Перечень графического материала: 1) структурная схема; 2) функциональная схема; 3) принципиальная электрическая схема; 4) алгоритм работы в виде блок-схем; 5) общий вид устройства; 6) экономическое обоснование.", size=14, line=1.15)
    add_text(doc, "6. Дата выдачи задания «02» ноября 2023 г.", size=14, line=1.15, after=6)
    add_text(doc, "Научный руководитель ___________ Маркин Н.И.", first=0, size=14, line=1.15)
    add_text(doc, "Задание принял к исполнению ___________ Божко А.П.", first=0, size=14, line=1.15, after=8)
    add_text(doc, "КАЛЕНДАРНЫЙ ПЛАН", center=True, first=0, bold=True, size=14, line=1.0, before=6)
    add_table(
        doc,
        "Таблица А.1 – Сроки этапов",
        ["Этап", "Срок"],
        [
            ["Анализ объекта модернизации", "02.11.2023 – 05.12.2023"],
            ["Структурная схема", "06.12.2023 – 13.12.2023"],
            ["Функциональная схема", "14.12.2023 – 15.01.2024"],
            ["Подбор элементной базы", "16.01.2024 – 05.03.2024"],
            ["Принципиальная электрическая схема", "06.03.2024 – 01.04.2024"],
            ["Схема общего вида", "02.04.2024 – 07.04.2024"],
            ["Алгоритм работы", "08.04.2024 – 20.04.2024"],
            ["Программное обеспечение", "21.04.2024 – 08.05.2024"],
            ["Экономическое обоснование", "09.05.2024 – 17.05.2024"],
            ["Оформление пояснительной записки", "18.05.2024 – 23.05.2024"],
        ],
        font=11,
    )
    add_text(doc, "Студент ___________ Божко А.П.", first=0, size=14, line=1.15)
    add_text(doc, "Научный руководитель ___________ Маркин Н.И.", first=0, size=14, line=1.15)


def annotation(doc):
    m = META
    add_page_break(doc)
    add_text(doc, "АННОТАЦИЯ", center=True, first=0, bold=True, after=12)
    add_text(doc, "Выпускная квалификационная работа посвящена модернизации системы управления линии разделки бревен портального типа.")
    add_text(doc, f"Работа содержит введение, девять разделов, заключение и список использованных источников. Общий объем работы составляет {m['pages']} страниц, {m['figures']} рисунков, {m['tables']} таблиц, {m['sources']} источников.")
    add_text(doc, "Ключевые слова: система управления, линия разделки бревен, модернизация, контроллер, преобразователь частоты, конструкторская документация.")
    add_text(doc, "Объект исследования: линия разделки бревен портального типа. Предмет исследования: система управления линией. Цель работы: модернизация системы управления линии разделки бревен портального типа.")
    add_text(doc, "В первом разделе разобраны способы разделки бревен и станок, принятый за объект. Во втором разделе составлена структурная схема. В третьем разделе составлена функциональная схема. В четвертом разделе подобрана элементная база. В пятом разделе разработана принципиальная электрическая схема. В шестом разделе выполнен общий вид устройства. В седьмом разделе составлен алгоритм работы. В восьмом разделе описана программа контроллера. В девятом разделе приведено экономическое обоснование.")
    add_page_break(doc)
    add_text(doc, "ANNOTATION", center=True, first=0, bold=True, after=12)
    add_text(doc, "The final qualifying work is devoted to the modernization of the control system of a portal-type log cutting line.")
    add_text(doc, f"The work contains an introduction, nine sections, a conclusion and a list of references. The total volume is {m['pages']} pages, {m['figures']} figures, {m['tables']} tables and {m['sources']} sources.")
    add_text(doc, "Key words: control system, log cutting line, modernization, controller, frequency converter, design documentation.")
    add_text(doc, "Object of the work: a portal-type log cutting line. Subject of the work: the control system of the line. Purpose of the work: modernization of the control system of a portal-type log cutting line.")
    add_text(doc, "The first section describes log cutting methods and the machine taken as the object. The second section gives the structural diagram. The third section gives the functional diagram. The fourth section selects the components. The fifth section develops the circuit diagram. The sixth section gives the general view. The seventh section gives the operating algorithm. The eighth section describes the controller program. The ninth section gives the economic assessment.")


def contents(doc):
    add_page_break(doc)
    add_text(doc, "СОДЕРЖАНИЕ", center=True, first=0, bold=True, after=12)
    for title, page in TOC:
        dots = "." * max(2, 50 - len(title))
        add_text(doc, f"{title} {dots} {page}", align="left", first=0, line=1.15, size=14)


def bibliography(doc):
    add_page_break(doc)
    add_text(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", center=True, first=0, bold=True, after=12)
    for i, item in enumerate(BIB, 1):
        add_text(doc, f"{i}. {item}", align="justify", first=0, line=1.15, size=14)


def build():
    doc = new_document()
    title_page(doc)
    assignment(doc)
    annotation(doc)
    contents(doc)
    emit_blocks(doc, expand_chapter1(A) + expand_rest(B) + expand_economy(C))
    bibliography(doc)
    doc.save(OUT)
    print("saved", OUT, "sources", len(BIB))


if __name__ == "__main__":
    build()
