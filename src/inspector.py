# Файл: /home/detker/Документы/repository/doc-forge/src/inspector.py
"""
Модуль инспекции и реверс-инжиниринга стилей существующих документов DOCX в декларативные YAML-профили.
"""

from pathlib import Path
from typing import Optional, Dict, Any
import yaml
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from src.schema import (
    DocProfile,
    MetaInfo,
    PageConfig,
    PageMargins,
    PageNumbering,
    TextStyle,
    HeadingStyles,
    ElementStyles,
    TableStyle,
    ListStyle,
    ReferenceStyle,
    TitleDefaults,
)


def inspect_docx(file_path: Path, profile_name: Optional[str] = None) -> DocProfile:
    """
    Анализирует .docx файл и извлекает метрики форматирования в объект DocProfile.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Файл образца '{file_path}' не найден!")

    doc = Document(str(file_path))
    name = profile_name or file_path.stem.lower().replace(" ", "_")

    # 1. Извлекаем параметры страницы из первого раздела
    sec = doc.sections[0]
    left_mm = round(sec.left_margin.mm, 1) if sec.left_margin else 30.0
    right_mm = round(sec.right_margin.mm, 1) if sec.right_margin else 10.0
    top_mm = round(sec.top_margin.mm, 1) if sec.top_margin else 20.0
    bottom_mm = round(sec.bottom_margin.mm, 1) if sec.bottom_margin else 20.0

    page_config = PageConfig(
        format="A4",
        orientation="portrait",
        margins=PageMargins(
            left_mm=left_mm,
            right_mm=right_mm,
            top_mm=top_mm,
            bottom_mm=bottom_mm,
        ),
        page_numbering=PageNumbering(
            enabled=True,
            start_from_page=2,
            position="top_right",
            font_family="Times New Roman",
            font_size_pt=12.0,
        ),
    )

    # 2. Анализируем параграфы для определения стиля основного текста (Normal)
    body_style = TextStyle(
        font_family="Times New Roman",
        font_size_pt=14.0,
        line_spacing=1.5,
        first_line_indent_cm=1.25,
        alignment="JUSTIFY",
        spacing_before_pt=0.0,
        spacing_after_pt=0.0,
    )

    headings = HeadingStyles()

    for p in doc.paragraphs:
        style_name = p.style.name.lower() if p.style else ""
        text = p.text.strip()
        if not text:
            continue

        # Проверяем шрифт первого Run
        run_font = None
        run_size = None
        run_bold = False
        run_italic = False
        if p.runs:
            r0 = p.runs[0]
            run_font = r0.font.name
            if r0.font.size:
                run_size = r0.font.size.pt
            run_bold = bool(r0.font.bold)
            run_italic = bool(r0.font.italic)

        # Выравнивание
        align_str = "JUSTIFY"
        if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            align_str = "CENTER"
        elif p.alignment == WD_ALIGN_PARAGRAPH.LEFT:
            align_str = "LEFT"
        elif p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            align_str = "RIGHT"

        # Отступ первой строки
        indent_cm = 1.25
        if p.paragraph_format.first_line_indent:
            indent_cm = round(p.paragraph_format.first_line_indent.cm, 2)

        # Анализ H1
        if "heading 1" in style_name or "заголовок 1" in style_name or (run_bold and align_str == "CENTER" and len(text) < 60 and text.isupper()):
            headings.h1.font_family = run_font or headings.h1.font_family
            headings.h1.font_size_pt = run_size or headings.h1.font_size_pt
            headings.h1.alignment = align_str
            headings.h1.all_caps = text.isupper()
            if p.paragraph_format.space_before:
                headings.h1.spacing_before_pt = round(p.paragraph_format.space_before.pt, 1)
            if p.paragraph_format.space_after:
                headings.h1.spacing_after_pt = round(p.paragraph_format.space_after.pt, 1)

        # Анализ H2
        elif "heading 2" in style_name or "заголовок 2" in style_name or (run_bold and align_str == "LEFT" and len(text) < 80):
            headings.h2.font_family = run_font or headings.h2.font_family
            headings.h2.font_size_pt = run_size or headings.h2.font_size_pt
            headings.h2.alignment = align_str
            headings.h2.first_line_indent_cm = indent_cm
            if p.paragraph_format.space_before:
                headings.h2.spacing_before_pt = round(p.paragraph_format.space_before.pt, 1)
            if p.paragraph_format.space_after:
                headings.h2.spacing_after_pt = round(p.paragraph_format.space_after.pt, 1)

        # Анализ основного текста
        elif "normal" in style_name or "основной" in style_name or (len(text) > 100):
            if run_font:
                body_style.font_family = run_font
            if run_size:
                body_style.font_size_pt = run_size
            if p.paragraph_format.line_spacing:
                body_style.line_spacing = float(p.paragraph_format.line_spacing)
            body_style.first_line_indent_cm = indent_cm
            body_style.alignment = align_str

    profile = DocProfile(
        meta=MetaInfo(
            name=name,
            institution="Извлечено из образца " + file_path.name,
            standard="ГОСТ 7.32 / СТО",
            description=f"Автоматически сгенерированный профиль из {file_path.name}",
        ),
        page=page_config,
        body=body_style,
        headings=headings,
        elements=ElementStyles(
            tables=TableStyle(
                font_family=body_style.font_family,
                font_size_pt=12.0,
                line_spacing=1.0,
                caption_format="Таблица {id} — {title}",
            ),
            lists=ListStyle(bullet_style="dash", first_line_indent_cm=1.25),
            references=ReferenceStyle(
                font_family=body_style.font_family,
                font_size_pt=body_style.font_size_pt,
                line_spacing=body_style.line_spacing,
                first_line_indent_cm=1.25,
            ),
        ),
        title_defaults=TitleDefaults(),
    )

    return profile


def save_profile_to_yaml(profile: DocProfile, output_path: Path):
    """
    Сохраняет профиль в формате YAML с аккуратной структурой.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    data = profile.model_dump()
    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
