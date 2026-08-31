# Файл: src/converters/pdf_engine.py
"""
Движок глубокого анализа и конвертации PDF документов в структурированный Markdown + профиль ГОСТ.
"""

import re
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
from collections import Counter
import yaml

import fitz  # PyMuPDF
import pdfplumber

from src.converters.base import BaseConverter, ConversionResult
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


class PdfConverter(BaseConverter):
    """
    Конвертер PDF -> Markdown с извлечением пространственной геометрии, шрифтовой карты и таблиц.
    """

    def can_handle(self, file_path: Path) -> bool:
        return file_path.suffix.lower() == ".pdf"

    def convert(self, file_path: Path, extract_profile: bool = True) -> ConversionResult:
        if not file_path.exists():
            raise FileNotFoundError(f"PDF файл '{file_path}' не найден!")

        doc = fitz.open(str(file_path))
        pages_count = len(doc)

        # 1. Анализируем геометрию страниц и границы текста
        page_config, detected_margins = self._analyze_pdf_geometry(doc)

        # 2. Анализируем шрифтовую структуру и заголовки
        body_style, heading_styles, fonts_info = self._analyze_pdf_fonts_and_headings(doc)

        # 3. Извлекаем таблицы с помощью pdfplumber
        extracted_tables = self._extract_tables_with_plumber(file_path)

        # 4. Преобразуем блоки текста в Markdown
        md_text, headings_hierarchy = self._extract_markdown_blocks(doc, body_style, heading_styles, extracted_tables)

        # 5. Собираем итоговый профиль
        profile_name = file_path.stem.lower().replace(" ", "_").replace("-", "_")
        profile = DocProfile(
            meta=MetaInfo(
                name=profile_name,
                institution="Извлечено из PDF",
                standard="ГОСТ / СТО (извлечено из PDF)",
                description=f"Профиль макета, извлеченный из файла {file_path.name} ({pages_count} стр.)",
            ),
            page=page_config,
            body=body_style,
            headings=heading_styles,
            elements=ElementStyles(
                tables=TableStyle(
                    font_family=body_style.font_family,
                    font_size_pt=12.0,
                    line_spacing=1.0,
                    borders=True,
                ),
                lists=ListStyle(bullet_style="dash", first_line_indent_cm=body_style.first_line_indent_cm),
                references=ReferenceStyle(font_family=body_style.font_family, font_size_pt=body_style.font_size_pt),
            ),
            title_defaults=TitleDefaults(),
        )

        # 6. Генерируем отчет о макете
        layout_report = self._generate_pdf_layout_report(
            profile, pages_count, len(extracted_tables), len(headings_hierarchy), fonts_info
        )

        doc.close()

        return ConversionResult(
            markdown=md_text,
            profile=profile if extract_profile else None,
            metadata={
                "source_file": str(file_path),
                "pages_count": pages_count,
                "tables_count": len(extracted_tables),
                "headings_count": len(headings_hierarchy),
                "fonts_info": fonts_info,
                "format": "PDF",
            },
            layout_report=layout_report,
            tables_count=len(extracted_tables),
            images_count=0,
            pages_count=pages_count,
            headings_hierarchy=headings_hierarchy,
        )

    def _analyze_pdf_geometry(self, doc: fitz.Document) -> Tuple[PageConfig, Dict[str, float]]:
        """
        Вычисляет физические размеры листа и отступы полей (в мм).
        """
        PT_TO_MM = 25.4 / 72.0  # 1 пункт = 0.3528 мм

        left_margins: List[float] = []
        right_margins: List[float] = []
        top_margins: List[float] = []
        bottom_margins: List[float] = []

        page_w_mm = 210.0
        page_h_mm = 297.0
        orientation = "portrait"

        for page in doc:
            rect = page.rect
            page_w_mm = round(rect.width * PT_TO_MM, 1)
            page_h_mm = round(rect.height * PT_TO_MM, 1)
            if page_w_mm > page_h_mm:
                orientation = "landscape"

            # Анализируем bounding box блоков текста
            blocks = page.get_text("blocks")
            for b in blocks:
                # b = (x0, y0, x1, y1, text, block_no, block_type)
                if b[6] == 0 and len(b[4].strip()) > 15:  # Текстовый блок достаточной длины
                    left_mm = round(b[0] * PT_TO_MM, 1)
                    top_mm = round(b[1] * PT_TO_MM, 1)
                    right_mm = round((rect.width - b[2]) * PT_TO_MM, 1)
                    bottom_mm = round((rect.height - b[3]) * PT_TO_MM, 1)

                    if 5.0 <= left_mm <= 60.0:
                        left_margins.append(left_mm)
                    if 5.0 <= right_mm <= 60.0:
                        right_margins.append(right_mm)
                    if 5.0 <= top_mm <= 60.0:
                        top_margins.append(top_mm)
                    if 5.0 <= bottom_mm <= 60.0:
                        bottom_margins.append(bottom_mm)

        # Выбираем минимальные левые границы (начало строк)
        calc_left = round(min(left_margins), 1) if left_margins else 30.0
        calc_right = round(min(right_margins), 1) if right_margins else 10.0
        calc_top = round(min(top_margins), 1) if top_margins else 20.0
        calc_bottom = round(min(bottom_margins), 1) if bottom_margins else 20.0

        # Нормализуем под стандартные поля ГОСТ
        if 25.0 <= calc_left <= 35.0:
            calc_left = 30.0
        if 8.0 <= calc_right <= 15.0:
            calc_right = 10.0
        if 16.0 <= calc_top <= 24.0:
            calc_top = 20.0
        if 16.0 <= calc_bottom <= 24.0:
            calc_bottom = 20.0

        margins_dict = {
            "left_mm": calc_left,
            "right_mm": calc_right,
            "top_mm": calc_top,
            "bottom_mm": calc_bottom,
        }

        page_config = PageConfig(
            format="A4" if (200 <= page_w_mm <= 220) else "Custom",
            orientation=orientation, # type: ignore
            margins=PageMargins(**margins_dict),
            page_numbering=PageNumbering(
                enabled=True,
                start_from_page=2,
                position="top_right",
                font_family="Times New Roman",
                font_size_pt=12.0,
            ),
        )
        return page_config, margins_dict

    def _analyze_pdf_fonts_and_headings(
        self, doc: fitz.Document
    ) -> Tuple[TextStyle, HeadingStyles, Dict[str, Any]]:
        """
        Строит гистограмму шрифтов, кеглей и начертаний в PDF.
        """
        fonts_counter = Counter()
        sizes_counter = Counter()
        flags_counter = Counter()

        for page in doc:
            page_dict = page.get_text("dict")
            for block in page_dict.get("blocks", []):
                if block.get("type") == 0:  # Text block
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text = span.get("text", "").strip()
                            if not text:
                                continue
                            font_name = span.get("font", "Times New Roman")
                            # Очищаем системные префиксы шрифтов (ABCDE+TimesNewRomanPSMT -> Times New Roman)
                            clean_font = re.sub(r'^[A-Z]{6}\+', '', font_name)
                            if "times" in clean_font.lower():
                                clean_font = "Times New Roman"
                            elif "arial" in clean_font.lower():
                                clean_font = "Arial"
                            elif "calibri" in clean_font.lower():
                                clean_font = "Calibri"

                            size = round(span.get("size", 14.0), 1)
                            fonts_counter[clean_font] += len(text)
                            sizes_counter[size] += len(text)
                            flags_counter[span.get("flags", 0)] += len(text)

        dominant_font = fonts_counter.most_common(1)[0][0] if fonts_counter else "Times New Roman"
        dominant_size = sizes_counter.most_common(1)[0][0] if sizes_counter else 14.0

        # Нормализуем кегль
        if 13.5 <= dominant_size <= 14.5:
            dominant_size = 14.0
        elif 11.5 <= dominant_size <= 12.5:
            dominant_size = 12.0

        body_style = TextStyle(
            font_family=dominant_font,
            font_size_pt=dominant_size,
            line_spacing=1.5,
            first_line_indent_cm=1.25,
            alignment="JUSTIFY",
            spacing_before_pt=0.0,
            spacing_after_pt=0.0,
        )

        h1_style = TextStyle(
            font_family=dominant_font,
            font_size_pt=dominant_size,
            bold=True,
            all_caps=True,
            alignment="CENTER",
            spacing_before_pt=12.0,
            spacing_after_pt=12.0,
            page_break_before=True,
            keep_with_next=True,
        )

        h2_style = TextStyle(
            font_family=dominant_font,
            font_size_pt=dominant_size,
            bold=True,
            all_caps=False,
            alignment="LEFT",
            first_line_indent_cm=1.25,
            spacing_before_pt=12.0,
            spacing_after_pt=6.0,
            keep_with_next=True,
        )

        h3_style = TextStyle(
            font_family=dominant_font,
            font_size_pt=dominant_size,
            bold=False,
            italic=True,
            alignment="LEFT",
            first_line_indent_cm=1.25,
            spacing_before_pt=6.0,
            spacing_after_pt=6.0,
            keep_with_next=True,
        )

        heading_styles = HeadingStyles(h1=h1_style, h2=h2_style, h3=h3_style)

        fonts_info = {
            "dominant_font": dominant_font,
            "dominant_size_pt": dominant_size,
            "all_fonts": dict(fonts_counter.most_common(5)),
            "all_sizes": dict(sizes_counter.most_common(5)),
        }

        return body_style, heading_styles, fonts_info

    def _extract_tables_with_plumber(self, file_path: Path) -> List[List[List[str]]]:
        """
        Извлекает структурированные таблицы с помощью pdfplumber.
        """
        tables: List[List[List[str]]] = []
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    extracted = page.extract_tables()
                    for t in extracted:
                        clean_table = []
                        for row in t:
                            clean_row = [str(cell or "").strip().replace("\n", " ") for cell in row]
                            if any(clean_row):
                                clean_table.append(clean_row)
                        if len(clean_table) >= 2:
                            tables.append(clean_table)
        except Exception:
            pass
        return tables

    def _extract_markdown_blocks(
        self,
        doc: fitz.Document,
        body_style: TextStyle,
        heading_styles: HeadingStyles,
        tables: List[List[List[str]]],
    ) -> Tuple[str, List[str]]:
        """
        Формирует структурированный Markdown на основе блоков PyMuPDF.
        """
        md_paragraphs: List[str] = []
        headings_hierarchy: List[str] = []
        toc_entries = []

        # Словарь обнаруженных элементов титульника
        title_meta: Dict[str, Any] = {}

        for page_num, page in enumerate(doc, 1):
            page_text = page.get_text("text")
            
            # Если 1 страница похожа на титульник
            if page_num == 1:
                if "МИНИСТЕРСТВО" in page_text.upper() or "КУРСОВОЙ" in page_text.upper():
                    title_meta = self._parse_title_text(page_text)
                    continue

            blocks = page.get_text("blocks")
            for b in blocks:
                if b[6] != 0:  # Skip image blocks
                    continue
                text = b[4].strip()
                if not text:
                    continue

                # Исключаем изолированные номера страниц
                if len(text) <= 3 and text.isdigit():
                    continue

                # Проверяем H1 (КАПС, номер главы или Введение)
                if (text.isupper() and len(text) < 80) or re.match(r'^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ|\d+\s+[А-ЯЁA-Z])', text):
                    clean_h1 = text.replace("\n", " ").strip()
                    md_paragraphs.append(f"\n# {clean_h1}\n")
                    headings_hierarchy.append(f"H1: {clean_h1}")
                    toc_entries.append({"title": clean_h1, "level": 1, "page": page_num})
                    continue

                # Проверяем H2 (1.1, 1.2, 2.1)
                if re.match(r'^\d+\.\d+\s+', text) and len(text) < 100:
                    clean_h2 = text.replace("\n", " ").strip()
                    md_paragraphs.append(f"\n## {clean_h2}\n")
                    headings_hierarchy.append(f"H2: {clean_h2}")
                    toc_entries.append({"title": clean_h2, "level": 2, "page": page_num})
                    continue

                # Проверяем H3 (1.1.1)
                if re.match(r'^\d+\.\d+\.\d+\s+', text) and len(text) < 100:
                    clean_h3 = text.replace("\n", " ").strip()
                    md_paragraphs.append(f"\n### {clean_h3}\n")
                    headings_hierarchy.append(f"H3: {clean_h3}")
                    toc_entries.append({"title": clean_h3, "level": 3, "page": page_num})
                    continue

                # Списки
                if text.startswith("- ") or text.startswith("• "):
                    items = text.split("\n")
                    for it in items:
                        if it.strip():
                            clean_it = re.sub(r'^[•\-\–—]\s*', '', it.strip())
                            md_paragraphs.append(f"- {clean_it}")
                    continue

                # Обычный абзац
                clean_para = text.replace("\n", " ").strip()
                md_paragraphs.append(clean_para)

        # Добавляем извлеченные таблицы
        for idx, table in enumerate(tables, 1):
            t_md = self._format_table_to_gfm(table)
            md_paragraphs.append(f"\n{t_md}\n")

        # Собираем Frontmatter
        frontmatter_dict: Dict[str, Any] = {
            "profile": "pdf_extracted",
        }
        if title_meta:
            frontmatter_dict["title"] = title_meta
        if toc_entries:
            frontmatter_dict["toc"] = toc_entries

        fm_yaml = yaml.dump(frontmatter_dict, allow_unicode=True, sort_keys=False).strip()
        full_md = f"---\n{fm_yaml}\n---\n\n" + "\n\n".join(md_paragraphs)
        full_md = re.sub(r'\n{3,}', '\n\n', full_md)

        return full_md, headings_hierarchy

    def _parse_title_text(self, text: str) -> Dict[str, Any]:
        """
        Извлекает поля титульного листа из текстового блока первой страницы.
        """
        meta = {}
        lines = [line.strip() for line in text.split("\n") if line.strip()]
        for line in lines:
            if "МИНИСТЕРСТВО" in line.upper():
                meta["ministry"] = line
            elif any(w in line.upper() for w in ["КОЛЛЕДЖ", "УНИВЕРСИТЕТ", "ИНСТИТУТ"]):
                if "institution" not in meta:
                    meta["institution"] = line
            elif any(w in line.upper() for w in ["КУРСОВОЙ ПРОЕКТ", "ДИПЛОМНЫЙ ПРОЕКТ", "РЕФЕРАТ"]):
                meta["work_type"] = line
            elif re.search(r'\d{2}\.\d{2}\.\d{2}', line):
                meta["specialty"] = line
            elif "тема:" in line.lower():
                meta["theme"] = re.sub(r'^тема:\s*', '', line, flags=re.IGNORECASE).strip('«»" ')
        return meta

    def _format_table_to_gfm(self, table_grid: List[List[str]]) -> str:
        if not table_grid:
            return ""
        headers = table_grid[0]
        cols = len(headers)
        res = [
            "| " + " | ".join(headers) + " |",
            "| " + " | ".join([":---"] * cols) + " |",
        ]
        for row in table_grid[1:]:
            padded = row + [""] * (cols - len(row))
            res.append("| " + " | ".join(padded[:cols]) + " |")
        return "\n".join(res)

    def _generate_pdf_layout_report(
        self,
        profile: DocProfile,
        pages_count: int,
        tables_count: int,
        headings_count: int,
        fonts_info: Dict[str, Any],
    ) -> str:
        p = profile
        m = p.page.margins
        lines = [
            f"📑 АНАЛИЗ PDF ДОКУМЕНТА ({pages_count} страниц):",
            f"   • Геометрия страницы: {p.page.format} ({p.page.orientation})",
            f"   • Вычисленные поля: Левое {m.left_mm} мм, Правое {m.right_mm} мм, Верхнее {m.top_mm} мм, Нижнее {m.bottom_mm} мм",
            "",
            f"🔤 КАРТА ШРИФТОВ И ТИПОГРАФИКИ:",
            f"   • Доминирующий шрифт: {fonts_info.get('dominant_font')} ({fonts_info.get('dominant_size_pt')} pt)",
            f"   • Распределение гарнитур: {fonts_info.get('all_fonts')}",
            f"   • Распределение кеглей: {fonts_info.get('all_sizes')}",
            "",
            f"📌 СТРУКТУРНЫЕ ЭЛЕМЕНТЫ:",
            f"   • Заголовков обнаружено: {headings_count} шт.",
            f"   • Таблиц извлечено: {tables_count} шт.",
        ]
        return "\n".join(lines)
