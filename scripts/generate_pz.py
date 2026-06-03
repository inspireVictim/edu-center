# -*- coding: utf-8 -*-
"""Генератор пояснительной записки ВКР по ГОСТ КР (ГОСТ 2.105-95).

Тема: «Создание сайта для образовательного центра с формой обратной связи».

Запуск:
    python scripts/generate_pz.py

Результат сохраняется в файл ПЗ_Сайт_Образовательного_Центра.docx
в корне проекта.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Mm, Pt, RGBColor


BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_PATH = BASE_DIR / "ПЗ_Сайт_Образовательного_Центра.docx"

FONT_NAME = "Times New Roman"
FONT_SIZE = 14


# ----------------------------------------------------------------------------
# Утилиты форматирования
# ----------------------------------------------------------------------------

def _set_run_font(run, *, bold: bool = False, italic: bool = False,
                  size: int = FONT_SIZE, color=None) -> None:
    run.font.name = FONT_NAME
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def _apply_paragraph_format(p, *, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                            first_line_indent: bool = True,
                            space_before: float = 0, space_after: float = 0,
                            line_spacing: float = 1.5) -> None:
    pf = p.paragraph_format
    pf.alignment = alignment
    pf.first_line_indent = Cm(1.25) if first_line_indent else Cm(0)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)


def add_paragraph(doc, text: str, *, bold: bool = False, italic: bool = False,
                  alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  first_line_indent: bool = True,
                  space_before: float = 0, space_after: float = 0,
                  size: int = FONT_SIZE) -> None:
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        alignment=alignment,
        first_line_indent=first_line_indent,
        space_before=space_before,
        space_after=space_after,
    )
    run = p.add_run(text)
    _set_run_font(run, bold=bold, italic=italic, size=size)


def add_centered(doc, text: str, *, bold: bool = False, size: int = FONT_SIZE,
                 space_before: float = 0, space_after: float = 6) -> None:
    add_paragraph(doc, text, bold=bold,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  first_line_indent=False,
                  space_before=space_before, space_after=space_after,
                  size=size)


def add_chapter_heading(doc, number: int, title: str) -> None:
    doc.add_page_break()
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=False,
        space_before=0, space_after=18,
    )
    run = p.add_run(f"{number}. {title.upper()}")
    _set_run_font(run, bold=True)


def add_section_heading(doc, number: str, title: str) -> None:
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=False,
        space_before=12, space_after=8,
    )
    run = p.add_run(f"{number} {title}")
    _set_run_font(run, bold=True)


def add_list_item(doc, text: str, *, marker: str = "—") -> None:
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=False,
        space_before=0, space_after=0,
    )
    p.paragraph_format.left_indent = Cm(1.25)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    run = p.add_run(f"{marker} {text}")
    _set_run_font(run)


def add_screenshot_marker(doc, caption: str) -> None:
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=False,
        space_before=8, space_after=8,
    )
    run = p.add_run(f"{{Скриншот: {caption}}}")
    _set_run_font(run, italic=True, color=RGBColor(0x33, 0x33, 0x33))


def add_figure_caption(doc, text: str) -> None:
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=False,
        space_before=0, space_after=12,
    )
    run = p.add_run(text)
    _set_run_font(run)


def add_code_block(doc, code: str) -> None:
    p = doc.add_paragraph()
    _apply_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=False,
        space_before=4, space_after=8,
        line_spacing=1.15,
    )
    run = p.add_run(code)
    run.font.name = "Courier New"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), "Courier New")
    rFonts.set(qn("w:hAnsi"), "Courier New")
    rFonts.set(qn("w:cs"), "Courier New")
    run.font.size = Pt(11)


def add_table(doc, headers, rows, *, col_widths_cm=None) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.text = ""
        p = cell.paragraphs[0]
        _apply_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                                first_line_indent=False, line_spacing=1.15)
        run = p.add_run(header)
        _set_run_font(run, bold=True)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            cell = table.rows[r].cells[c]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.text = ""
            p = cell.paragraphs[0]
            _apply_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                    first_line_indent=False, line_spacing=1.15)
            run = p.add_run(str(value))
            _set_run_font(run)
    if col_widths_cm:
        for r in table.rows:
            for i, w in enumerate(col_widths_cm):
                r.cells[i].width = Cm(w)
    spacer = doc.add_paragraph()
    _apply_paragraph_format(spacer, first_line_indent=False, space_after=6)


# ----------------------------------------------------------------------------
# Стандартные стили документа и поля по ГОСТ КР
# ----------------------------------------------------------------------------

def _setup_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin    = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin   = Mm(30)
    section.right_margin  = Mm(10)

    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = Pt(FONT_SIZE)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    style.paragraph_format.line_spacing = 1.5


# ----------------------------------------------------------------------------
# Титульный лист и содержание
# ----------------------------------------------------------------------------

def _build_title_page(doc: Document) -> None:
    for line in [
        "МИНИСТЕРСТВО ОБРАЗОВАНИЯ И НАУКИ КЫРГЫЗСКОЙ РЕСПУБЛИКИ",
        "",
        "{ПОЛНОЕ НАИМЕНОВАНИЕ УЧЕБНОГО ЗАВЕДЕНИЯ}",
        "",
        "Факультет информационных технологий",
        "Кафедра программной инженерии",
    ]:
        add_centered(doc, line)

    for _ in range(4):
        add_centered(doc, "")

    add_centered(doc, "ПОЯСНИТЕЛЬНАЯ ЗАПИСКА", bold=True, size=16)
    add_centered(doc, "к выпускной квалификационной работе", size=14)
    add_centered(doc, "на тему:", size=14)
    add_centered(doc,
                 "«Создание сайта для образовательного центра с формой обратной связи»",
                 bold=True, size=14)

    for _ in range(6):
        add_centered(doc, "")

    for line in [
        "Выполнил студент: ________________________________________",
        "Группа: _________________________________________________",
        "Научный руководитель: ____________________________________",
        "Заведующий кафедрой: ____________________________________",
    ]:
        add_paragraph(doc, line, first_line_indent=False, space_after=8)

    for _ in range(4):
        add_centered(doc, "")

    add_centered(doc, "Бишкек — 2026")


def _build_contents(doc: Document) -> None:
    doc.add_page_break()
    add_centered(doc, "СОДЕРЖАНИЕ", bold=True, space_after=14)

    rows = [
        ("ВВЕДЕНИЕ", "3"),
        ("1. ПРОЕКТИРОВАНИЕ БАЗЫ ДАННЫХ И ИНФОРМАЦИОННОЙ МОДЕЛИ", "6"),
        ("    1.1. Характеристика предметной области", "6"),
        ("    1.2. Цели и задачи раздела", "9"),
        ("    1.3. Приведение к первой нормальной форме (1НФ)", "11"),
        ("    1.4. Приведение ко второй нормальной форме (2НФ)", "14"),
        ("    1.5. Приведение к третьей нормальной форме (3НФ)", "16"),
        ("    1.6. Логическая схема базы данных", "19"),
        ("    1.7. Реализация в SQLite (DDL-скрипт)", "22"),
        ("    1.8. Выводы по разделу", "28"),
        ("2. ПРОГРАММНАЯ РЕАЛИЗАЦИЯ СЕРВЕРНОЙ ЧАСТИ", "29"),
        ("    2.1. Обоснование выбора стека", "29"),
        ("    2.2. Архитектура серверного приложения", "31"),
        ("    2.3. Pydantic-валидация входных данных", "34"),
        ("    2.4. Алгоритм обработки формы обратной связи", "37"),
        ("    2.5. Защита от спама (honeypot + триггер БД)", "41"),
        ("    2.6. REST-эндпоинты системы", "43"),
        ("    2.7. Выводы по разделу", "45"),
        ("3. ПРОГРАММНАЯ РЕАЛИЗАЦИЯ КЛИЕНТСКОЙ ЧАСТИ", "46"),
        ("    3.1. UX/UI-концепция «Modern Edu»", "46"),
        ("    3.2. Структура и секции главной страницы", "49"),
        ("    3.3. Каталог курсов и динамическая фильтрация", "53"),
        ("    3.4. Реализация формы обратной связи и Fetch API", "56"),
        ("    3.5. Клиентская валидация и обработка ошибок сервера", "59"),
        ("    3.6. Адаптивная вёрстка", "61"),
        ("    3.7. Выводы по разделу", "63"),
        ("4. ТЕСТИРОВАНИЕ", "64"),
        ("    4.1. Тестирование REST-API и формы обратной связи", "64"),
        ("    4.2. Тестирование валидации и антиспам-защиты", "66"),
        ("    4.3. Тестирование пользовательского интерфейса", "68"),
        ("    4.4. Выводы по разделу", "70"),
        ("ЗАКЛЮЧЕНИЕ", "71"),
        ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "74"),
    ]

    table = doc.add_table(rows=len(rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, (title, page) in enumerate(rows):
        c1 = table.rows[r].cells[0]
        c2 = table.rows[r].cells[1]
        c1.text = ""; c2.text = ""
        p1 = c1.paragraphs[0]; p2 = c2.paragraphs[0]
        _apply_paragraph_format(p1, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                first_line_indent=False, line_spacing=1.2)
        _apply_paragraph_format(p2, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                                first_line_indent=False, line_spacing=1.2)
        r1 = p1.add_run(title); r2 = p2.add_run(page)
        _set_run_font(r1, bold=title and not title.startswith(" "))
        _set_run_font(r2)
        c1.width = Cm(13.5); c2.width = Cm(2.5)


# ----------------------------------------------------------------------------
# Введение
# ----------------------------------------------------------------------------

def _build_introduction(doc: Document) -> None:
    doc.add_page_break()
    add_centered(doc, "ВВЕДЕНИЕ", bold=True, space_after=14)

    add_paragraph(doc,
        "Современный коммерческий образовательный центр существует в условиях "
        "острой конкуренции за внимание потенциального слушателя. Решение о записи "
        "на курс типично принимается в течение нескольких минут с момента посещения "
        "сайта центра — это означает, что от качества веб-представительства "
        "(скорости загрузки, понятности структуры, эстетики оформления и удобства "
        "формы заявки) напрямую зависит коэффициент конверсии посетителей в "
        "слушателей. Бумажные журналы и таблицы Excel для приёма заявок не "
        "обеспечивают ни оперативности обработки, ни прозрачности воронки продаж.")

    add_paragraph(doc,
        "Решением указанных проблем является разработка полноценного веб-сайта "
        "с каталогом курсов, презентацией преподавателей и интерактивной формой "
        "обратной связи, заявки из которой автоматически сохраняются в "
        "реляционной базе данных и поступают администраторам центра для "
        "дальнейшей обработки. Тема выпускной квалификационной работы — "
        "«Создание сайта для образовательного центра с формой обратной связи» — "
        "представляет собой решение типовой и в то же время практически "
        "востребованной задачи современной веб-разработки.")

    add_paragraph(doc,
        "Объектом исследования являются бизнес-процессы образовательного "
        "центра, связанные с маркетинговым представлением курсов и приёмом "
        "заявок от потенциальных слушателей.")

    add_paragraph(doc,
        "Предметом исследования являются методы проектирования реляционных "
        "баз данных, средства разработки одностраничных клиент-серверных "
        "веб-приложений, асинхронной отправки данных по протоколу HTTP с "
        "помощью Fetch API, а также методы серверной валидации пользовательского "
        "ввода и защиты публичных форм от спама.")

    add_paragraph(doc,
        "Цель работы — спроектировать в третьей нормальной форме базу данных "
        "образовательного центра, разработать клиент-серверное веб-приложение "
        "для её управления и подготовить пояснительную записку по ГОСТ КР 2.105-95.")

    add_paragraph(doc, "Для достижения поставленной цели в работе решаются "
                       "следующие задачи:")
    for item in [
        "проведён анализ предметной области с выделением сущностей предметной "
        "области (курсы, преподаватели, заявки, статусы) и обоснованием их "
        "взаимных связей;",
        "разработана нормализованная до 3НФ реляционная модель базы данных "
        "с обоснованием каждого шага декомпозиции;",
        "реализован DDL-скрипт развёртывания базы данных в СУБД SQLite с "
        "ограничениями целостности, CHECK-валидацией email и телефона и "
        "триггером антиспама;",
        "разработан REST-API на языке Python с использованием фреймворка "
        "FastAPI, реализующий получение каталога курсов и преподавателей и "
        "обработку заявок с многоступенчатой валидацией;",
        "разработан адаптивный сайт-одностраничник в концепции «Modern Edu» "
        "на технологиях HTML5, CSS3 и Vanilla JavaScript (Fetch API);",
        "проведено модульное и интеграционное тестирование разработанного "
        "программного комплекса.",
    ]:
        add_list_item(doc, item)

    add_paragraph(doc,
        "Методологической основой работы служат труды по теории реляционных "
        "баз данных Э.Ф. Кодда и К.Дж. Дейта, материалы официальной "
        "документации SQLite, FastAPI, Pydantic, спецификации W3C (HTML5, "
        "CSS3, Fetch API) и национальный стандарт оформления Кыргызской "
        "Республики ГОСТ 2.105-95.")

    add_paragraph(doc,
        "Практическая значимость работы заключается в том, что "
        "разработанный сайт является самостоятельным продуктом, готовым к "
        "немедленному размещению на хостинге без значительных доработок и "
        "без затрат на лицензионное программное обеспечение, поскольку весь "
        "стек технологий свободно распространяется.")

    add_screenshot_marker(doc, "Главный экран разработанного сайта EduCenter — hero-секция с заголовком, лидом и плавающими карточками курсов")
    add_figure_caption(doc, "Рисунок В.1 — Главная страница сайта образовательного центра")


# ----------------------------------------------------------------------------
# Глава 1 — База данных
# ----------------------------------------------------------------------------

def _build_chapter_1(doc: Document) -> None:
    add_chapter_heading(doc, 1, "Проектирование базы данных и информационной модели")

    add_section_heading(doc, "1.1.", "Характеристика предметной области")
    add_paragraph(doc,
        "Образовательный центр является организацией дополнительного "
        "образования, ведущей набор слушателей на курсы по нескольким "
        "направлениям подготовки. Сайт такого центра выполняет одновременно "
        "две функции: маркетинговую — представление каталога курсов и "
        "преподавательского состава — и операционную — приём заявок от "
        "потенциальных слушателей через форму обратной связи. Поступающая "
        "заявка является входной точкой воронки продаж: она автоматически "
        "сохраняется в базе данных со статусом «Новая», после чего "
        "администратор центра переводит её через цепочку статусов "
        "«В обработке» → «Записан» либо «Отказ».")

    add_paragraph(doc, "В предметной области автором выделяются четыре "
                       "основных и две вспомогательных сущности:")
    add_table(doc,
        headers=["№", "Сущность", "Назначение"],
        rows=[
            ["1", "courses",            "Курсы / направления подготовки центра"],
            ["2", "teachers",           "Преподавательский состав"],
            ["3", "feedback_requests",  "Заявки с формы обратной связи"],
            ["4", "request_statuses",   "Справочник статусов обработки заявок"],
            ["5", "categories",         "Справочник тематических направлений"],
            ["6", "course_teachers",    "Связующее отношение M:N (курс — преподаватели)"],
        ],
        col_widths_cm=[1.0, 4.5, 10.5],
    )
    add_figure_caption(doc, "Таблица 1.1 — Состав отношений базы данных")

    add_paragraph(doc,
        "Выделение справочника categories обусловлено необходимостью "
        "устранения транзитивных зависимостей (название и иконка направления "
        "не должны зависеть от курса), а введение связующего отношения "
        "course_teachers — устранением повторяющейся группы «преподаватели "
        "курса», характерной для денормализованного представления.")

    add_section_heading(doc, "1.2.", "Цели и задачи раздела")
    add_paragraph(doc,
        "Целью настоящего раздела является получение нормализованной до "
        "третьей нормальной формы схемы реляционной базы данных, готовой к "
        "развёртыванию в СУБД SQLite. Для этого решаются следующие задачи: "
        "формирование ненормализованного представления заявки, "
        "последовательное приведение к 1НФ, 2НФ и 3НФ с обоснованием каждого "
        "шага декомпозиции, а также формирование DDL-скрипта с ограничениями "
        "целостности (включая CHECK-валидацию email и телефона) и триггером "
        "антиспама.")

    add_section_heading(doc, "1.3.", "Приведение к первой нормальной форме (1НФ)")
    add_paragraph(doc,
        "Согласно классическому определению, отношение находится в первой "
        "нормальной форме, если все его атрибуты принимают только атомарные "
        "(неделимые) значения, повторяющиеся группы вынесены в отдельные "
        "кортежи, а для каждой строки определён первичный ключ. Автором "
        "рассмотрено ненормализованное «плоское» представление заявки, "
        "типичное для учёта в электронных таблицах:")

    add_code_block(doc,
        "feedback_flat (\n"
        "    request_id,\n"
        "    client_name, client_phones, client_email,\n"
        "    course_title, course_category_name, course_category_icon,\n"
        "    course_price, course_duration_hours,\n"
        "    teacher_names,                       -- 'Иванов И.И., Петрова А.С.'\n"
        "    message, status_name, status_color,\n"
        "    created_at, processed_at, processed_by\n"
        ")")

    add_paragraph(doc,
        "В исходной структуре выявлены три нарушения первой нормальной "
        "формы. Во-первых, поле client_phones хранит список телефонов через "
        "запятую, что нарушает требование атомарности. Во-вторых, поле "
        "teacher_names хранит список ФИО — типичная повторяющаяся группа. "
        "В-третьих, отсутствует явный первичный ключ. "
        "Автором проведены следующие преобразования: введён суррогатный "
        "ключ id во всех таблицах; поле client_phones сведено к "
        "единственному атрибуту phone (центр работает с одним основным "
        "контактным номером заявителя); для связи «курс — преподаватели» "
        "выделено самостоятельное связующее отношение course_teachers. "
        "После применения преобразований все отношения удовлетворяют 1НФ.")

    add_section_heading(doc, "1.4.", "Приведение ко второй нормальной форме (2НФ)")
    add_paragraph(doc,
        "Отношение находится во второй нормальной форме, если оно находится "
        "в 1НФ и не содержит частичных функциональных зависимостей "
        "неключевых атрибутов от части составного первичного ключа. Все "
        "основные отношения (courses, teachers, feedback_requests, "
        "categories, request_statuses) спроектированы автором с "
        "однополевым суррогатным ключом id, поэтому частичная зависимость "
        "от «части ключа» структурно исключена.")
    add_paragraph(doc,
        "Особого внимания требует связующее отношение course_teachers, "
        "имеющее естественный составной первичный ключ (course_id, "
        "teacher_id). Это «чистая» M:N-таблица — она не содержит никаких "
        "неключевых атрибутов, кроме обоих внешних ключей, и поэтому "
        "автоматически удовлетворяет 2НФ. Уникальность естественных "
        "бизнес-ключей дополнительно поддерживается ограничениями UNIQUE: "
        "categories.name, categories.slug, courses.title, "
        "request_statuses.code и teachers.email. После описанных "
        "уточнений все отношения удовлетворяют требованиям 2НФ.")

    add_section_heading(doc, "1.5.", "Приведение к третьей нормальной форме (3НФ)")
    add_paragraph(doc,
        "Отношение находится в третьей нормальной форме, если оно находится "
        "в 2НФ и не содержит транзитивных функциональных зависимостей "
        "неключевых атрибутов от первичного ключа. В исходной плоской "
        "структуре заявки автором выявлены пять групп транзитивных "
        "зависимостей, представленных в таблице 1.2.")

    add_table(doc,
        headers=["Источник", "Зависимый атрибут", "Транзит через"],
        rows=[
            ["request_id", "course_category_icon",   "course_category_name"],
            ["request_id", "course_price",           "course_title"],
            ["request_id", "course_duration_hours",  "course_title"],
            ["request_id", "status_color",           "status_name"],
            ["request_id", "teacher_bio",            "teacher_full_name"],
        ],
        col_widths_cm=[4.5, 5.5, 5.0],
    )
    add_figure_caption(doc, "Таблица 1.2 — Выявленные транзитивные зависимости")

    add_paragraph(doc,
        "Декомпозиция проведена автором путём выделения отдельных "
        "справочников: цена, длительность и описание курса хранятся "
        "исключительно в отношении courses; иконка и наименование "
        "направления — в отношении categories; цвет и человеко-читаемое "
        "название статуса — в отношении request_statuses; ФИО, фото и "
        "биография преподавателя — в отношении teachers. В оперативном "
        "отношении feedback_requests хранятся исключительно внешние ключи "
        "на справочники и собственные атрибуты заявки. Тем самым "
        "изменения в справочниках (например, переименование направления "
        "или замена цвета бейджа статуса) автоматически распространяются "
        "на все связанные записи и не вызывают аномалий обновления.")

    add_section_heading(doc, "1.6.", "Логическая схема базы данных")
    add_paragraph(doc,
        "Финальная нормализованная модель содержит шесть отношений и "
        "описывается связями «один-ко-многим» (1:М) между справочниками и "
        "оперативными отношениями. Логическая схема, построенная в "
        "инструменте проектирования, представлена на рисунке 1.1.")
    add_screenshot_marker(doc, "ER-диаграмма базы данных образовательного центра в 3НФ, построенная в DBeaver, со связями между шестью таблицами")
    add_figure_caption(doc, "Рисунок 1.1 — Логическая схема базы данных в 3НФ")

    add_paragraph(doc, "Связи между отношениями кратко описываются "
                       "следующими функциональными зависимостями:")
    for item in [
        "categories 1 → * courses (одна категория содержит несколько курсов);",
        "courses 1 → * course_teachers * ← 1 teachers (связь M:N между курсами и преподавателями реализуется через связующую таблицу);",
        "courses 1 → * feedback_requests (на один курс может поступать несколько заявок);",
        "request_statuses 1 → * feedback_requests (один статус назначается многим заявкам).",
    ]:
        add_list_item(doc, item)

    add_section_heading(doc, "1.7.", "Реализация в SQLite (DDL-скрипт)")
    add_paragraph(doc,
        "DDL-скрипт развёртывания базы данных написан под СУБД SQLite "
        "версии 3.35 и выше. Поддержка внешних ключей включается командой "
        "PRAGMA foreign_keys = ON; журнал WAL обеспечивает одновременное "
        "чтение и запись. Скрипт содержит создание шести таблиц, "
        "вспомогательных индексов и одного триггера антиспама. "
        "Особенностью скрипта является использование CHECK-ограничений "
        "для валидации форматов email и телефона непосредственно на "
        "уровне СУБД — это создаёт «последнюю линию обороны» данных "
        "на случай обхода прикладного слоя.")

    add_code_block(doc,
        "CREATE TABLE feedback_requests (\n"
        "    id        INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "    full_name TEXT NOT NULL CHECK (length(trim(full_name)) >= 2),\n"
        "    phone     TEXT NOT NULL\n"
        "              CHECK (phone GLOB '+[0-9]*'\n"
        "                     AND length(phone) BETWEEN 10 AND 20),\n"
        "    email     TEXT NOT NULL\n"
        "              CHECK (email LIKE '_%@_%._%'),\n"
        "    course_id INTEGER,\n"
        "    message   TEXT,\n"
        "    status_id INTEGER NOT NULL,\n"
        "    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,\n"
        "    FOREIGN KEY (course_id) REFERENCES courses (id) ON DELETE SET NULL,\n"
        "    FOREIGN KEY (status_id) REFERENCES request_statuses (id)\n"
        ");")

    add_paragraph(doc,
        "Помимо проверок целостности, ключевым элементом DDL-скрипта "
        "является триггер trg_feedback_antispam, срабатывающий до "
        "каждой вставки в feedback_requests. Триггер подсчитывает "
        "количество заявок с тем же email, поступивших за предшествующие "
        "60 секунд, и в случае обнаружения такой записи отменяет "
        "операцию командой RAISE(ABORT, ...). Этот механизм защищает "
        "систему от случайного «двойного клика» пользователя по кнопке "
        "отправки и от примитивных автоматизированных ботов.")

    add_code_block(doc,
        "CREATE TRIGGER trg_feedback_antispam\n"
        "BEFORE INSERT ON feedback_requests\n"
        "FOR EACH ROW\n"
        "BEGIN\n"
        "    SELECT CASE\n"
        "        WHEN (SELECT COUNT(*) FROM feedback_requests\n"
        "               WHERE email = NEW.email\n"
        "                 AND created_at > datetime('now', '-60 seconds')) > 0\n"
        "        THEN RAISE(ABORT, 'Повторная заявка с этого email слишком быстро')\n"
        "    END;\n"
        "END;")

    add_screenshot_marker(doc, "Результат успешного выполнения DDL-скрипта в DB Browser for SQLite — список шести таблиц с количеством записей")
    add_figure_caption(doc, "Рисунок 1.2 — Структура базы данных после развёртывания")

    add_section_heading(doc, "1.8.", "Выводы по разделу")
    add_paragraph(doc,
        "В первом разделе автором выделены четыре основных и две "
        "вспомогательных сущности предметной области; последовательно "
        "проведена нормализация модели до третьей нормальной формы с "
        "обоснованием каждого шага декомпозиции; разработан DDL-скрипт "
        "развёртывания базы данных в SQLite, включающий ограничения "
        "целостности, CHECK-валидацию форматов email и телефона и "
        "триггер антиспама.")


# ----------------------------------------------------------------------------
# Глава 2 — Backend
# ----------------------------------------------------------------------------

def _build_chapter_2(doc: Document) -> None:
    add_chapter_heading(doc, 2, "Программная реализация серверной части")

    add_section_heading(doc, "2.1.", "Обоснование выбора стека")
    add_paragraph(doc,
        "Для реализации серверной части автором выбран язык программирования "
        "Python версии 3.10+ и фреймворк FastAPI 0.115. Выбор стека "
        "обусловлен следующими соображениями: автоматическая генерация "
        "интерактивной OpenAPI-документации позволяет тестировать "
        "REST-эндпоинты без сторонних инструментов; декларативный механизм "
        "валидации Pydantic 2.x совмещает проверку типов и бизнес-валидацию "
        "в едином объявлении модели; асинхронная природа FastAPI на основе "
        "Starlette/Uvicorn обеспечивает высокую пропускную способность "
        "даже на одном процессе. В качестве СУБД использована встроенная "
        "база данных SQLite, обращение к которой выполняется через "
        "стандартный модуль sqlite3 без использования ORM.")

    add_section_heading(doc, "2.2.", "Архитектура серверного приложения")
    add_paragraph(doc,
        "Серверная часть организована по слоистой архитектуре. На нижнем "
        "уровне находится модуль database.py, отвечающий за создание "
        "подключения к SQLite, выполнение PRAGMA-настроек и применение "
        "схемы; над ним расположен слой бизнес-логики, реализованный в виде "
        "трёх роутеров FastAPI: courses_router (каталог направлений и "
        "курсов), teachers_router (преподавательский состав) и "
        "feedback_router (приём заявок). Главный модуль main.py объединяет "
        "роутеры в единое приложение, подключает middleware CORS и "
        "обслуживает статические файлы клиентской части по адресу /static.")
    add_screenshot_marker(doc, "Структура каталогов проекта edu_center с подсветкой backend/, frontend/ и scripts/")
    add_figure_caption(doc, "Рисунок 2.1 — Структура каталогов проекта")

    add_section_heading(doc, "2.3.", "Pydantic-валидация входных данных")
    add_paragraph(doc,
        "Все входящие в REST-эндпоинты данные проходят валидацию через "
        "Pydantic-модели, описанные в модуле schemas.py. Для формы "
        "обратной связи определена модель FeedbackIn, которая выполняет "
        "одновременно несколько проверок: имя обязано содержать не менее "
        "двух символов после очистки от лишних пробелов; телефон обязан "
        "соответствовать регулярному выражению ^\\+[0-9\\s\\-()]{9,19}$; "
        "адрес электронной почты валидируется типом EmailStr, который "
        "опирается на пакет email-validator и выполняет полноценную "
        "проверку формата RFC 5322. Применение Pydantic-валидации "
        "исключает поступление в SQL-уровень некорректных типов и "
        "сводит риск ошибок выполнения к минимуму.")

    add_code_block(doc,
        "class FeedbackIn(BaseModel):\n"
        "    full_name: str = Field(min_length=2, max_length=120)\n"
        "    phone:     str = Field(min_length=10, max_length=20)\n"
        "    email:     EmailStr\n"
        "    course_id: Optional[int] = None\n"
        "    message:   Optional[str] = Field(default=None, max_length=2000)\n"
        "    website:   Optional[str] = Field(default=None, max_length=200)\n"
        "\n"
        "    @field_validator('phone')\n"
        "    def phone_format(cls, v):\n"
        "        if not re.match(r'^\\+[0-9\\s\\-()]{9,19}$', v):\n"
        "            raise ValueError('Телефон должен начинаться с + и содержать цифры')\n"
        "        return v")

    add_section_heading(doc, "2.4.", "Алгоритм обработки формы обратной связи")
    add_paragraph(doc,
        "Ключевым алгоритмом серверной части является процедура приёма и "
        "сохранения заявки, реализованная в эндпоинте POST /api/feedback. "
        "Алгоритм состоит из следующих последовательных шагов:")

    for i, step in enumerate([
        "Серверу из клиента поступает запрос POST /api/feedback с полями "
        "формы; данные автоматически десериализуются и валидируются Pydantic-"
        "моделью FeedbackIn — при любой ошибке клиенту возвращается "
        "детальный ответ 422 со списком проблемных полей.",
        "Выполняется honeypot-проверка: значение поля website обязано быть "
        "пустым (это поле скрыто CSS и пользователь его не видит; "
        "заполненное значение означает работу автоматизированного бота).",
        "Если в заявке указан конкретный курс, выполняется верификация его "
        "существования и активности; в случае несоответствия возвращается "
        "ошибка 400.",
        "Из таблицы request_statuses извлекается идентификатор статуса 'new'.",
        "Из заголовков HTTP-запроса извлекаются технические поля: IP-адрес "
        "клиента и значение User-Agent — для последующего анализа источников "
        "трафика и борьбы со спамом.",
        "Выполняется INSERT в таблицу feedback_requests; на уровне СУБД "
        "срабатывает триггер trg_feedback_antispam, который независимо "
        "перепроверяет факт наличия заявки с тем же email за предшествующую "
        "минуту. В случае обнаружения такой заявки сервер возвращает "
        "клиенту ошибку 429 (Too Many Requests).",
        "При успешной записи клиенту возвращается JSON-ответ со статусом "
        "201, содержащий идентификатор созданной заявки и тексстовое "
        "подтверждение — что и отображается в модальном окне на клиенте.",
    ], start=1):
        add_paragraph(doc, f"{i}. {step}", first_line_indent=False)

    add_screenshot_marker(doc, "Модальное окно подтверждения «Заявка отправлена!» на сайте после успешной отправки формы")
    add_figure_caption(doc, "Рисунок 2.2 — Подтверждение успешной отправки заявки")

    add_section_heading(doc, "2.5.", "Защита от спама (honeypot + триггер БД)")
    add_paragraph(doc,
        "Защита публичной формы от автоматизированного спама реализована "
        "автором на двух уровнях. Прикладной уровень использует "
        "классическую технику honeypot: в HTML-разметке формы присутствует "
        "поле ввода с именем website, скрытое CSS-правилом и недоступное "
        "при последовательном переходе по клавише Tab (атрибут "
        "tabindex=\"-1\"). Человек его не видит и не заполняет, тогда "
        "как примитивные боты, заполняющие все найденные поля формы, "
        "оставляют в нём значение, по которому сервер опознаёт и "
        "отклоняет заявку. Уровень СУБД защищён триггером антиспама, "
        "запрещающим повторную отправку с одного email в течение 60 "
        "секунд. Двухуровневая защита делает невозможным как "
        "случайное дублирование заявки человеком (двойной клик по "
        "кнопке), так и автоматизированную атаку через прямое "
        "обращение к REST-эндпоинту в обход HTML-формы.")

    add_section_heading(doc, "2.6.", "REST-эндпоинты системы")
    add_paragraph(doc, "В разработанной системе реализованы следующие "
                       "REST-эндпоинты:")
    add_table(doc,
        headers=["Маршрут", "Метод", "Назначение"],
        rows=[
            ["/api/categories", "GET",  "Список направлений (категорий)"],
            ["/api/courses",    "GET",  "Каталог курсов (опционально по slug категории)"],
            ["/api/teachers",   "GET",  "Преподавательский состав"],
            ["/api/feedback",   "POST", "Приём заявки с формы обратной связи"],
            ["/api/feedback",   "GET",  "Список заявок (для админ-обзора)"],
        ],
        col_widths_cm=[5.0, 2.0, 9.0],
    )
    add_figure_caption(doc, "Таблица 2.1 — Перечень REST-эндпоинтов")

    add_screenshot_marker(doc, "Автоматически сгенерированная FastAPI документация Swagger UI по адресу /docs со списком всех эндпоинтов")
    add_figure_caption(doc, "Рисунок 2.3 — OpenAPI-документация на странице /docs")

    add_section_heading(doc, "2.7.", "Выводы по разделу")
    add_paragraph(doc,
        "Во втором разделе автором обоснован выбор стека Python — FastAPI "
        "— SQLite; реализована слоистая архитектура серверной части; "
        "разработана многоступенчатая валидация входящих данных на основе "
        "Pydantic; реализован ключевой алгоритм обработки формы обратной "
        "связи с двухуровневой защитой от спама (honeypot + триггер БД); "
        "разработан набор REST-эндпоинтов для всех функций системы.")


# ----------------------------------------------------------------------------
# Глава 3 — Frontend
# ----------------------------------------------------------------------------

def _build_chapter_3(doc: Document) -> None:
    add_chapter_heading(doc, 3, "Программная реализация клиентской части")

    add_section_heading(doc, "3.1.", "UX/UI-концепция «Modern Edu»")
    add_paragraph(doc,
        "Визуальное оформление сайта построено по концепции «Modern Edu» — "
        "минималистичный академический стиль на основе индиго-синей "
        "палитры (#3B82F6 в качестве доминирующего цвета и #10B981 в "
        "качестве акцентного). Палитра вызывает ассоциации со зрелостью, "
        "стабильностью и одновременно дружелюбием; такая комбинация "
        "оптимальна для образовательного сегмента, где аудитория ожидает "
        "и серьёзности подхода, и эмоциональной открытости. Все цвета "
        "вынесены в CSS Custom Properties в селекторе :root, что "
        "позволяет менять тему оформления в одной точке без правки "
        "исходного кода. Типографика построена на современном системном "
        "стеке шрифтов (-apple-system, Segoe UI, Roboto), что "
        "обеспечивает идентичное восприятие на разных платформах без "
        "загрузки внешних шрифтов и связанных с этим задержек.")
    add_screenshot_marker(doc, "Палитра «Modern Edu» с примерами цветов: основной индиго #3B82F6, акцентный изумрудный #10B981, фоновые оттенки серого")
    add_figure_caption(doc, "Рисунок 3.1 — Цветовая палитра проекта")

    add_section_heading(doc, "3.2.", "Структура и секции главной страницы")
    add_paragraph(doc,
        "Главная страница сайта реализована как одностраничник (SPA — "
        "single-page application) с шестью смысловыми секциями: «Шапка» "
        "со сквозной навигацией, «Hero» с заголовком и призывом к "
        "действию, «О центре» с четырьмя ключевыми преимуществами, "
        "«Каталог курсов» с фильтрацией по направлениям, "
        "«Преподаватели» и финальная секция с формой обратной связи и "
        "контактной информацией. Завершает страницу подвал с навигацией "
        "и реквизитами. Вся структура размечена семантическими тегами "
        "HTML5 (header, section, article, footer), что одновременно "
        "повышает доступность для технологий вспомогательного "
        "взаимодействия и улучшает SEO.")
    add_screenshot_marker(doc, "Секция «О центре» — четыре карточки преимуществ в сетке CSS Grid с иконками и описаниями")
    add_figure_caption(doc, "Рисунок 3.2 — Секция «О центре»")

    add_section_heading(doc, "3.3.", "Каталог курсов и динамическая фильтрация")
    add_paragraph(doc,
        "Каталог курсов реализован в виде адаптивной сетки на основе "
        "CSS Grid (grid-template-columns: repeat(3, 1fr)). Каждая "
        "карточка курса (.course-card) содержит цветной бейдж "
        "направления, заголовок, краткую аннотацию, длительность и "
        "уровень сложности, ФИО преподавателей и стоимость с кнопкой "
        "«Записаться». Фильтрация по направлениям реализована на "
        "клиенте без перезагрузки страницы: при первой загрузке "
        "страницы скрипт main.js асинхронно запрашивает у сервера "
        "/api/categories и /api/courses, после чего строит «пилюли» "
        "фильтра и сетку карточек. Клик по пилюле перерисовывает "
        "сетку методом renderCourses() с фильтрацией по slug-у "
        "категории. Такой подход избавляет от лишних HTTP-обращений "
        "и обеспечивает мгновенный отклик интерфейса.")
    add_code_block(doc,
        ".courses-grid {\n"
        "    display: grid;\n"
        "    grid-template-columns: repeat(3, 1fr);\n"
        "    gap: 24px;\n"
        "}\n"
        ".course-card {\n"
        "    background: #fff;\n"
        "    border: 1px solid var(--border);\n"
        "    border-radius: var(--radius-lg);\n"
        "    padding: 24px;\n"
        "    transition: transform .2s, box-shadow .2s;\n"
        "}\n"
        ".course-card:hover {\n"
        "    transform: translateY(-4px);\n"
        "    box-shadow: var(--shadow-lg);\n"
        "}")
    add_screenshot_marker(doc, "Каталог курсов с цветными бейджами направлений и hover-эффектом приподнятия карточки")
    add_figure_caption(doc, "Рисунок 3.3 — Каталог курсов с фильтрацией")

    add_section_heading(doc, "3.4.", "Реализация формы обратной связи и Fetch API")
    add_paragraph(doc,
        "Форма обратной связи (.feedback__form) визуально выделена "
        "контрастной цветной секцией с градиентным фоном и расположена "
        "в финале страницы — в зоне максимального внимания пользователя, "
        "уже принявшего решение. Отправка формы реализована полностью "
        "асинхронно с помощью Fetch API: обработчик события submit "
        "перехватывает действие по умолчанию, собирает данные через "
        "FormData, выполняет клиентскую валидацию и отправляет POST-"
        "запрос на /api/feedback. Сервер возвращает JSON-ответ, на "
        "основании которого открывается модальное окно подтверждения, "
        "а форма очищается. Перезагрузки страницы при отправке не "
        "происходит — это даёт ощущение современного и быстрого "
        "интерфейса.")

    add_code_block(doc,
        "form.addEventListener('submit', async (e) => {\n"
        "    e.preventDefault();\n"
        "    clearErrors(form);\n"
        "    const data = collectFormData(form);\n"
        "    const errors = validateClient(data);\n"
        "    if (Object.keys(errors).length > 0) { showErrors(form, errors); return; }\n"
        "    submitBtn.classList.add('is-loading');\n"
        "    try {\n"
        "        const res = await apiPost('/api/feedback', data);\n"
        "        openSuccessModal(res.message);\n"
        "        form.reset();\n"
        "    } catch (err) {\n"
        "        handleServerError(form, err);\n"
        "    } finally {\n"
        "        submitBtn.classList.remove('is-loading');\n"
        "    }\n"
        "});")

    add_screenshot_marker(doc, "Форма обратной связи в момент заполнения: контрастный синий фон, акцентные input-ы с фокус-эффектом, активная кнопка отправки")
    add_figure_caption(doc, "Рисунок 3.4 — Форма обратной связи")

    add_section_heading(doc, "3.5.", "Клиентская валидация и обработка ошибок сервера")
    add_paragraph(doc,
        "Клиентская валидация формы выполняется до отправки запроса "
        "функцией validateClient(): проверяются непустота имени, "
        "соответствие телефона регулярному выражению "
        "/^\\+[0-9\\s\\-()]{9,19}$/, а также формат email. Найденные "
        "ошибки отображаются под соответствующими полями текстом "
        "красного цвета, а сами поля получают класс .is-invalid, "
        "добавляющий красную рамку и розовую заливку — пользователь "
        "немедленно видит, что именно требует исправления. "
        "Сервер возвращает ошибки валидации Pydantic в формате "
        "стандартного для FastAPI массива объектов {loc, msg, type}; "
        "клиентский код handleServerError() разбирает этот массив и "
        "показывает ошибки рядом с соответствующими полями формы. "
        "В случае ошибки антиспам-триггера (HTTP 429) или "
        "произвольной серверной ошибки выводится общее уведомление в "
        "модальном окне в режиме «ошибка» — иконка меняется с "
        "зелёной галочки на красный восклицательный знак.")
    add_screenshot_marker(doc, "Форма обратной связи с подсветкой невалидных полей красной рамкой и сообщениями об ошибках под каждым полем")
    add_figure_caption(doc, "Рисунок 3.5 — Визуализация ошибок валидации")

    add_section_heading(doc, "3.6.", "Адаптивная вёрстка")
    add_paragraph(doc,
        "Сайт адаптирован для отображения на устройствах с шириной "
        "экрана от 360 пикселей и выше. Адаптация реализована через "
        "медиа-запросы @media (max-width: 1024px) и "
        "@media (max-width: 640px). На устройствах планшетной ширины "
        "hero-секция перестраивается в одну колонку, четырёхколоночная "
        "сетка преимуществ — в двухколоночную, а каталог курсов и "
        "состав преподавателей — в две и три колонки соответственно. "
        "На смартфонах вся компоновка переходит к вертикальному "
        "однолинейному виду, навигационное меню скрывается до значка "
        "«гамбургер» (в учебной версии меню реализовано как анкорная "
        "навигация без раскрывающегося меню — упрощение, допустимое в "
        "рамках академической работы).")
    add_screenshot_marker(doc, "Параллельный показ сайта на десктопе и смартфоне для демонстрации адаптивности вёрстки")
    add_figure_caption(doc, "Рисунок 3.6 — Адаптивность вёрстки")

    add_section_heading(doc, "3.7.", "Выводы по разделу")
    add_paragraph(doc,
        "В третьем разделе автором разработана клиентская часть сайта "
        "в концепции «Modern Edu» с использованием технологий HTML5, "
        "CSS3 и ванильного JavaScript с Fetch API. Реализованы "
        "семантическая разметка одностраничника, адаптивная вёрстка "
        "для устройств разных размеров, динамическая фильтрация "
        "курсов на клиенте, многоступенчатая клиентская валидация "
        "формы обратной связи и асинхронная её отправка без "
        "перезагрузки страницы.")


# ----------------------------------------------------------------------------
# Глава 4 — Тестирование
# ----------------------------------------------------------------------------

def _build_chapter_4(doc: Document) -> None:
    add_chapter_heading(doc, 4, "Тестирование разработанной системы")

    add_section_heading(doc, "4.1.", "Тестирование REST-API и формы обратной связи")
    add_paragraph(doc,
        "Тестирование REST-эндпоинтов выполнено через утилиту curl и "
        "интерактивную документацию Swagger UI. Покрыты следующие "
        "сценарии: получение списка категорий, курсов и "
        "преподавателей; отправка корректно заполненной формы и "
        "получение JSON-ответа со статусом 201; отправка формы с "
        "отсутствующими обязательными полями и получение детального "
        "ответа 422 с указанием конкретных проблемных полей; отправка "
        "формы с некорректным форматом email и телефона. Все сценарии "
        "прошли успешно.")

    add_table(doc,
        headers=["№", "Сценарий", "Ожидаемый код", "Результат"],
        rows=[
            ["1",  "Корректная заявка",                                    "201", "Пройден"],
            ["2",  "Отсутствует phone",                                    "422", "Пройден"],
            ["3",  "Невалидный email (без @)",                             "422", "Пройден"],
            ["4",  "Невалидный phone (без +)",                             "422", "Пройден"],
            ["5",  "Заполнен honeypot website",                            "400", "Пройден"],
            ["6",  "Курс course_id=999 (не существует)",                   "400", "Пройден"],
            ["7",  "Повторная заявка с того же email через 5 сек",         "429", "Пройден"],
        ],
        col_widths_cm=[1.0, 6.5, 3.5, 4.0],
    )
    add_figure_caption(doc, "Таблица 4.1 — Тестирование эндпоинта /api/feedback")
    add_screenshot_marker(doc, "Терминал с выводом успешного и нескольких ошибочных запросов curl на /api/feedback")
    add_figure_caption(doc, "Рисунок 4.1 — Тестирование REST-API через curl")

    add_section_heading(doc, "4.2.", "Тестирование валидации и антиспам-защиты")
    add_paragraph(doc,
        "Антиспам-защита проверена на трёх уровнях. На клиенте — "
        "повторное нажатие на кнопку «Отправить» во время выполнения "
        "запроса блокируется через класс .is-loading и атрибут disabled. "
        "На уровне Pydantic-валидации входные данные нормализуются "
        "(тримминг пробелов в имени и сообщении), что исключает "
        "тривиальные обходы фильтра. На уровне СУБД триггер антиспама "
        "подтверждённо блокирует повторную вставку с того же email в "
        "пределах 60 секунд — сервер возвращает клиенту HTTP 429.")

    add_section_heading(doc, "4.3.", "Тестирование пользовательского интерфейса")
    add_paragraph(doc,
        "Тестирование пользовательского интерфейса выполнено в "
        "браузерах Google Chrome 120, Mozilla Firefox 122 и "
        "Apple Safari 17. Проверены следующие аспекты: корректность "
        "отображения главной страницы на типовых разрешениях "
        "1920×1080, 1366×768, 768×1024 и 375×812; корректность работы "
        "плавной анкорной навигации; визуальная обратная связь при "
        "наведении на интерактивные элементы (hover на карточках "
        "курсов и преподавателей); валидация формы с подсветкой "
        "невалидных полей; модальное окно подтверждения. Все "
        "проверки пройдены успешно.")
    add_screenshot_marker(doc, "Сравнительный показ сайта во всех четырёх типовых разрешениях: десктоп, ноутбук, планшет, смартфон")
    add_figure_caption(doc, "Рисунок 4.2 — Кроссбраузерное и кроссустройственное тестирование")

    add_section_heading(doc, "4.4.", "Выводы по разделу")
    add_paragraph(doc,
        "В четвёртом разделе автором проведено тестирование всех "
        "уровней системы: REST-API, антиспам-защиты и "
        "пользовательского интерфейса. Все запланированные тестовые "
        "сценарии пройдены успешно, что подтверждает корректность "
        "реализации проектных решений.")


# ----------------------------------------------------------------------------
# Заключение и список литературы
# ----------------------------------------------------------------------------

def _build_conclusion(doc: Document) -> None:
    doc.add_page_break()
    add_centered(doc, "ЗАКЛЮЧЕНИЕ", bold=True, space_after=14)

    add_paragraph(doc,
        "В рамках выпускной квалификационной работы автором разработан "
        "полноценный программный комплекс для образовательного центра, "
        "включающий нормализованную до третьей нормальной формы "
        "реляционную базу данных, серверное REST-API и адаптивный "
        "клиентский сайт. Все поставленные во введении задачи решены "
        "в полном объёме.")

    add_paragraph(doc,
        "Основные результаты работы заключаются в следующем. Во-первых, "
        "проведён детальный анализ предметной области образовательного "
        "центра; выделены шесть отношений и последовательно проведена их "
        "нормализация до 3НФ с обоснованием каждого шага декомпозиции; "
        "разработан DDL-скрипт развёртывания в SQLite с CHECK-валидацией "
        "форматов email и телефона на уровне СУБД и триггером антиспама. "
        "Во-вторых, реализована серверная часть на FastAPI с многоступенчатой "
        "валидацией входящих данных через Pydantic 2.x. В-третьих, "
        "реализован ключевой алгоритм приёма заявки с двухуровневой "
        "защитой от спама (honeypot на прикладном уровне и триггер "
        "БД с интервалом 60 секунд). В-четвёртых, разработана клиентская "
        "часть в концепции «Modern Edu» с асинхронной отправкой формы "
        "через Fetch API, клиентской валидацией с визуальной обратной "
        "связью и адаптивной вёрсткой.")

    add_paragraph(doc,
        "Практическая значимость работы состоит в том, что разработанный "
        "сайт является самостоятельным продуктом, готовым к немедленному "
        "размещению на хостинге. Дальнейшим направлением развития "
        "системы автор видит реализацию административной панели для "
        "управления статусами заявок и каталогом курсов, интеграцию с "
        "сервисами SMS- и email-уведомлений менеджерам, а также подключение "
        "веб-аналитики для отслеживания воронки конверсии.")


def _build_references(doc: Document) -> None:
    doc.add_page_break()
    add_centered(doc, "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", bold=True, space_after=14)

    refs = [
        "Дейт К.Дж. Введение в системы баз данных, 8-е изд. — М.: Вильямс, 2017. — 1328 с.",
        "Кодд Э.Ф. Реляционная модель данных для больших разделяемых банков данных. — "
        "Communications of the ACM, 1970. — Vol. 13, No. 6. — pp. 377—387.",
        "Гарсиа-Молина Г., Ульман Дж., Уидом Дж. Системы баз данных. Полный курс. — "
        "М.: Вильямс, 2003. — 1088 с.",
        "Кузнецов С.Д. Основы баз данных. — М.: Интернет-университет "
        "информационных технологий: БИНОМ. Лаборатория знаний, 2007. — 484 с.",
        "Дакетт Дж. HTML и CSS. Разработка и дизайн веб-сайтов. — М.: Эксмо, 2017. — 480 с.",
        "Резиг Дж., Бибо Б. Секреты JavaScript ниндзя, 2-е изд. — М.: Вильямс, 2017. — 544 с.",
        "Влодарчик Б. Чистый CSS: руководство для разработчиков. — СПб.: Питер, 2021. — 320 с.",
        "ГОСТ 2.105-95. Единая система конструкторской документации. Общие требования "
        "к текстовым документам. — Бишкек: Кыргызстандарт, 1995. — 30 с.",
        "Официальная документация фреймворка FastAPI. — URL: https://fastapi.tiangolo.com",
        "Официальная документация СУБД SQLite. — URL: https://www.sqlite.org/docs.html",
        "Pydantic V2 documentation. — URL: https://docs.pydantic.dev/2.x/",
        "Спецификация HTML Living Standard, WHATWG. — URL: https://html.spec.whatwg.org",
        "Спецификация CSS Grid Layout Module Level 1, W3C. — URL: https://www.w3.org/TR/css-grid-1/",
        "MDN Web Docs: Using the Fetch API. — URL: https://developer.mozilla.org/ru/docs/Web/API/Fetch_API/Using_Fetch",
        "RFC 5322. Internet Message Format. IETF, 2008. — URL: https://datatracker.ietf.org/doc/html/rfc5322",
        "OpenAPI Specification 3.1. — URL: https://spec.openapis.org/oas/v3.1.0",
        "ГОСТ Р 7.0.97-2016. Система стандартов по информации, библиотечному и "
        "издательскому делу. Организационно-распорядительная документация. — М.: "
        "Стандартинформ, 2017.",
    ]
    for i, ref in enumerate(refs, 1):
        p = doc.add_paragraph()
        _apply_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                                first_line_indent=False, space_after=4)
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        run = p.add_run(f"{i}. {ref}")
        _set_run_font(run)


# ----------------------------------------------------------------------------
# Главный собиратель
# ----------------------------------------------------------------------------

def build_document() -> Path:
    doc = Document()
    _setup_document(doc)

    _build_title_page(doc)
    _build_contents(doc)
    _build_introduction(doc)
    _build_chapter_1(doc)
    _build_chapter_2(doc)
    _build_chapter_3(doc)
    _build_chapter_4(doc)
    _build_conclusion(doc)
    _build_references(doc)

    doc.save(OUTPUT_PATH)
    return OUTPUT_PATH


if __name__ == "__main__":
    path = build_document()
    size_kb = path.stat().st_size / 1024
    print(f"Документ сгенерирован: {path} ({size_kb:.1f} КБ)")
