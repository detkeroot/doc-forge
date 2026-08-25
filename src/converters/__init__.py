# Файл: /home/detker/Документы/repository/doc-forge/src/converters/__init__.py
"""
Подсистема универсальной конвертации документов (PDF, DOCX, XLSX, PPTX, RTF, EPUB, CSV) в структурированный Markdown.
"""

from src.converters.base import BaseConverter, ConversionResult
from src.converters.router import convert_document, get_converter_for_file

__all__ = [
    "BaseConverter",
    "ConversionResult",
    "convert_document",
    "get_converter_for_file",
]
