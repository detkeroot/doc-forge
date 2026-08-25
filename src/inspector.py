# Файл: /home/detker/Документы/repository/doc-forge/src/inspector.py
"""
Инспектор DOCX-документов: извлечение метрик форматирования, геометрии и генерация YAML-профиля.
"""

from pathlib import Path
from typing import Dict, Any
import yaml
from docx import Document
from docx.shared import Mm, Pt

from src.schema import (
    DocProfile, MetaInfo, PageConfig, PageMargins, PageNumbering,
    TextStyle, HeadingStyles, ElementStyles, TableStyle, ListStyle, TitleDefaults
)


def inspect_docx(docx_path: Path, profile_name: str = "custom_profile") -> DocProfile:
    """
    Анализирует .docx файл образца и возвращает сгенерированный объект DocProfile.
    """
    doc = Document(str(docx_path))
    
    # 1. Анализируем геометрию страницы первого раздела
    sec = doc.sections[0]
    left_mm = round(sec.left_margin.mm, 1) if sec.left_margin else 30.0
    right_mm = round(sec.right_margin.mm, 1) if sec.right_margin else 10.0
    top_mm = round(sec.top_margin.mm, 1) if sec.top_margin else 20.0
    bottom_mm = round(sec.bottom_margin.mm, 1) if sec.bottom_margin else 20.0

    margins = PageMargins(
        left_mm=left_mm,
        right_mm=right_mm,
        top_mm=top_mm,
        bottom_mm=bottom_mm
    )

    # 2. Статистика шрифтов и отступов
    detected_fonts = {}
    detected_sizes = {}
    detected_indents = []
    detected_line_spacings = []

    for p in doc.paragraphs:
        if p.paragraph_format.first_line_indent:
            indent_cm = round(p.paragraph_format.first_line_indent.cm, 2)
            if indent_cm > 0:
                detected_indents.append(indent_cm)
        if p.paragraph_format.line_spacing:
            detected_line_spacings.append(p.paragraph_format.line_spacing)

        for r in p.runs:
            if r.font.name:
                detected_fonts[r.font.name] = detected_fonts.get(r.font.name, 0) + 1
            if r.font.size:
                size_pt = round(r.font.size.pt, 1)
                detected_sizes[size_pt] = detected_sizes.get(size_pt, 0) + 1

    # Выбираем самый частый шрифт
    main_font = max(detected_fonts, key=detected_fonts.get) if detected_fonts else "Times New Roman"
    main_size = max(detected_sizes, key=detected_sizes.get) if detected_sizes else 14.0
    main_indent = detected_indents[0] if detected_indents else 1.25
    main_spacing = detected_line_spacings[0] if detected_line_spacings else 1.5

    # 3. Собираем профиль
    profile = DocProfile(
        meta=MetaInfo(
            name=profile_name,
            institution="Извлечено из " + docx_path.name,
            standard="СТО / ГОСТ (Автодетект)",
            description=f"Автоматически сгенерированный профиль на основе {docx_path.name}"
        ),
        page=PageConfig(
            format="A4",
            orientation="portrait",
            margins=margins,
            page_numbering=PageNumbering(
                enabled=True,
                start_from_page=2,
                position="top_right",
                font_family=main_font,
                font_size_pt=12.0
            )
        ),
        body=TextStyle(
            font_family=main_font,
            font_size_pt=main_size,
            line_spacing=main_spacing if isinstance(main_spacing, (int, float)) else 1.5,
            first_line_indent_cm=main_indent,
            alignment="JUSTIFY",
            spacing_before_pt=0.0,
            spacing_after_pt=0.0
        ),
        headings=HeadingStyles(),
        elements=ElementStyles(),
        title_defaults=TitleDefaults()
    )

    return profile


def export_profile_yaml(profile: DocProfile, yaml_path: Path):
    """
    Экспортирует профиль в чистый структурированный YAML-файл.
    """
    yaml_path.parent.mkdir(parents=True, exist_ok=True)
    data = profile.model_dump()
    with open(yaml_path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True, sort_keys=False, indent=2)
    return yaml_path
