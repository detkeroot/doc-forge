# Файл: /home/detker/Документы/repository/doc-forge/src/converters/router.py
"""
Универсальный маршрутизатор (Router) конвертации документов в Markdown и профили ГОСТ.
"""

from pathlib import Path
from typing import List, Optional

from src.converters.base import BaseConverter, ConversionResult
from src.converters.docx_engine import DocxConverter
from src.converters.pdf_engine import PdfConverter
from src.converters.xlsx_engine import XlsxConverter
from src.converters.pptx_engine import PptxConverter
from src.converters.rtf_engine import RtfConverter
from src.converters.epub_engine import EpubConverter
from src.converters.csv_engine import CsvConverter


CONVERTERS: List[BaseConverter] = [
    DocxConverter(),
    PdfConverter(),
    XlsxConverter(),
    PptxConverter(),
    RtfConverter(),
    EpubConverter(),
    CsvConverter(),
]


def get_converter_for_file(file_path: Path) -> Optional[BaseConverter]:
    """
    Возвращает подходящий движок конвертера для указанного файла.
    """
    for converter in CONVERTERS:
        if converter.can_handle(file_path):
            return converter
    return None


def convert_document(file_path: Path, extract_profile: bool = True) -> ConversionResult:
    """
    Универсальная функция прямой конвертации любого документа (PDF, DOCX, XLSX, PPTX, RTF, EPUB, CSV) в Markdown.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Файл '{file_path}' не существует!")

    converter = get_converter_for_file(file_path)
    if not converter:
        supported = [".docx", ".pdf", ".xlsx", ".pptx", ".rtf", ".epub", ".csv", ".tsv"]
        raise ValueError(
            f"Неподдерживаемый формат файла: '{file_path.suffix}'. Поддерживаются: {', '.join(supported)}"
        )

    return converter.convert(file_path, extract_profile=extract_profile)
