# Файл: src/converters/rtf_engine.py
"""
Движок конвертации RTF (Rich Text Format) документов в Markdown.
"""

import re
from pathlib import Path
from striprtf.striprtf import rtf_to_text

from src.converters.base import BaseConverter, ConversionResult


class RtfConverter(BaseConverter):
    """
    Конвертер RTF -> Markdown.
    """

    def can_handle(self, file_path: Path) -> bool:
        return file_path.suffix.lower() == ".rtf"

    def convert(self, file_path: Path, extract_profile: bool = False) -> ConversionResult:
        if not file_path.exists():
            raise FileNotFoundError(f"RTF файл '{file_path}' не найден!")

        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            raw_rtf = f.read()

        plain_text = rtf_to_text(raw_rtf)
        
        # Структурируем заголовки и списки
        lines = [line.strip() for line in plain_text.split("\n")]
        md_lines = []

        for line in lines:
            if not line:
                continue

            # Заголовки (КАПС или нумерованные разделы)
            if (line.isupper() and len(line) < 70) or re.match(r'^(ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|\d+\s+[А-ЯЁA-Z])', line):
                md_lines.append(f"\n# {line}\n")
            elif re.match(r'^\d+\.\d+\s+', line):
                md_lines.append(f"\n## {line}\n")
            elif line.startswith("- ") or line.startswith("• "):
                clean_it = re.sub(r'^[•\-\–—]\s*', '', line)
                md_lines.append(f"- {clean_it}")
            else:
                md_lines.append(line)

        full_md = f"# Документ RTF: {file_path.name}\n\n" + "\n\n".join(md_lines)
        full_md = re.sub(r'\n{3,}', '\n\n', full_md)

        layout_report = (
            f"📄 ДОКУМЕНТ RTF ({file_path.name}):\n"
            f"   • Формат: Rich Text Format\n"
            f"   • Строк обработано: {len(lines)}"
        )

        return ConversionResult(
            markdown=full_md,
            profile=None,
            metadata={
                "source_file": str(file_path),
                "format": "RTF",
            },
            layout_report=layout_report,
            pages_count=1,
        )
