# Файл: /home/detker/Документы/repository/doc-forge/src/schema.py
"""
Pydantic схемы валидации декларативных профилей оформления документов (ГОСТ / СТО).
"""

from typing import Optional, Dict, Any, Literal
from pydantic import BaseModel, Field


class PageMargins(BaseModel):
    left_mm: float = Field(default=30.0, description="Левое поле (под переплет) в мм")
    right_mm: float = Field(default=10.0, description="Правое поле в мм")
    top_mm: float = Field(default=20.0, description="Верхнее поле в мм")
    bottom_mm: float = Field(default=20.0, description="Нижнее поле в мм")


class PageNumbering(BaseModel):
    enabled: bool = Field(default=True, description="Включена ли нумерация")
    start_from_page: int = Field(default=2, description="Номер первой нумеруемой страницы")
    position: Literal["top_right", "bottom_center", "top_center", "bottom_right"] = Field(
        default="top_right", description="Позиция номера страницы"
    )
    font_family: str = Field(default="Times New Roman", description="Шрифт номера страницы")
    font_size_pt: float = Field(default=12.0, description="Кегль номера страницы")


class PageConfig(BaseModel):
    format: str = Field(default="A4", description="Формат страницы")
    orientation: Literal["portrait", "landscape"] = Field(default="portrait", description="Ориентация")
    margins: PageMargins = Field(default_factory=PageMargins)
    page_numbering: PageNumbering = Field(default_factory=PageNumbering)


class TextStyle(BaseModel):
    font_family: str = Field(default="Times New Roman", description="Гарнитура шрифта")
    font_size_pt: float = Field(default=14.0, description="Размер шрифта в пунктах")
    line_spacing: float = Field(default=1.5, description="Межстрочный интервал")
    first_line_indent_cm: float = Field(default=1.25, description="Отступ первой строки (красная строка) в см")
    alignment: Literal["JUSTIFY", "LEFT", "CENTER", "RIGHT"] = Field(
        default="JUSTIFY", description="Выравнивание текста"
    )
    spacing_before_pt: float = Field(default=0.0, description="Отступ перед абзацем в pt")
    spacing_after_pt: float = Field(default=0.0, description="Отступ после абзаца в pt")
    bold: bool = Field(default=False, description="Полужирный")
    italic: bool = Field(default=False, description="Курсив")
    all_caps: bool = Field(default=False, description="Все прописные (КАПС)")
    keep_with_next: bool = Field(default=False, description="Не отрывать от следующего абзаца")
    page_break_before: bool = Field(default=False, description="Разрыв страницы перед элементом")


class HeadingStyles(BaseModel):
    h1: TextStyle = Field(
        default_factory=lambda: TextStyle(
            font_family="Times New Roman",
            font_size_pt=14.0,
            bold=True,
            all_caps=True,
            alignment="CENTER",
            first_line_indent_cm=0.0,
            spacing_before_pt=12.0,
            spacing_after_pt=12.0,
            page_break_before=True,
            keep_with_next=True,
        ),
        description="Заголовок 1 уровня (Главы, Разделы, Введение, Заключение)"
    )
    h2: TextStyle = Field(
        default_factory=lambda: TextStyle(
            font_family="Times New Roman",
            font_size_pt=14.0,
            bold=True,
            all_caps=False,
            alignment="LEFT",
            first_line_indent_cm=1.25,
            spacing_before_pt=12.0,
            spacing_after_pt=6.0,
            page_break_before=False,
            keep_with_next=True,
        ),
        description="Заголовок 2 уровня (Подразделы 1.1, 1.2)"
    )
    h3: TextStyle = Field(
        default_factory=lambda: TextStyle(
            font_family="Times New Roman",
            font_size_pt=14.0,
            bold=False,
            italic=True,
            alignment="LEFT",
            first_line_indent_cm=1.25,
            spacing_before_pt=6.0,
            spacing_after_pt=6.0,
            page_break_before=False,
            keep_with_next=True,
        ),
        description="Заголовок 3 уровня (Пункты 1.1.1)"
    )


class TableStyle(BaseModel):
    font_family: str = Field(default="Times New Roman")
    font_size_pt: float = Field(default=12.0, description="Кегль внутри таблицы")
    line_spacing: float = Field(default=1.0, description="Одинарный интервал внутри таблицы")
    caption_format: str = Field(default="Таблица {id} — {title}", description="Шаблон подписи таблицы")
    caption_alignment: Literal["LEFT", "RIGHT", "CENTER", "JUSTIFY"] = Field(default="LEFT")
    caption_first_line_indent_cm: float = Field(default=0.0)
    spacing_before_pt: float = Field(default=6.0)
    spacing_after_pt: float = Field(default=6.0)
    repeat_header_on_new_page: bool = Field(default=True, description="Повторять шапку на след. странице")
    borders: bool = Field(default=True, description="Отрисовывать сетку границ")


class ListStyle(BaseModel):
    bullet_style: Literal["dash", "bullet", "number"] = Field(default="dash", description="Стиль маркера")
    first_line_indent_cm: float = Field(default=1.25, description="Красная строка для списков")
    left_indent_cm: float = Field(default=1.25, description="Отступ слева")
    spacing_before_pt: float = Field(default=0.0)
    spacing_after_pt: float = Field(default=0.0)


class ReferenceStyle(BaseModel):
    font_family: str = Field(default="Times New Roman")
    font_size_pt: float = Field(default=14.0)
    line_spacing: float = Field(default=1.5)
    first_line_indent_cm: float = Field(default=1.25)
    hanging_indent_cm: float = Field(default=0.0)
    spacing_before_pt: float = Field(default=0.0)
    spacing_after_pt: float = Field(default=0.0)


class ElementStyles(BaseModel):
    tables: TableStyle = Field(default_factory=TableStyle)
    lists: ListStyle = Field(default_factory=ListStyle)
    references: ReferenceStyle = Field(default_factory=ReferenceStyle)


class TitleDefaults(BaseModel):
    ministry: str = Field(default="МИНИСТЕРСТВО ОБРАЗОВАНИЯ И НАУКИ САМАРСКОЙ ОБЛАСТИ")
    institution: str = Field(default="ГБПОУ «Самарский металлургический колледж»")
    faculty: Optional[str] = Field(default=None)
    department: Optional[str] = Field(default="Отделение информационных технологий")
    specialty: str = Field(default="09.02.07 Информационные системы и программирование")
    work_type: str = Field(default="КУРСОВОЙ ПРОЕКТ")
    discipline: Optional[str] = Field(default="МДК 01.01 Разработка программных модулей")
    theme: str = Field(default="Тема курсового проекта")
    student_name: str = Field(default="Еткарев Д. О.")
    student_group: str = Field(default="гр. ИСП-24")
    supervisor_name: str = Field(default="Иванов И. И.")
    supervisor_title: str = Field(default="преподаватель")
    city: str = Field(default="Самара")
    year: int = Field(default=2026)


class MetaInfo(BaseModel):
    name: str = Field(description="Уникальный идентификатор профиля")
    institution: str = Field(description="Учебное заведение")
    standard: str = Field(default="ГОСТ 7.32-2017", description="Стандарт оформления")
    description: Optional[str] = Field(default=None)


class DocProfile(BaseModel):
    meta: MetaInfo
    page: PageConfig = Field(default_factory=PageConfig)
    body: TextStyle = Field(default_factory=TextStyle)
    headings: HeadingStyles = Field(default_factory=HeadingStyles)
    elements: ElementStyles = Field(default_factory=ElementStyles)
    title_defaults: TitleDefaults = Field(default_factory=TitleDefaults)
