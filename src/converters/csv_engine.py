# Файл: /home/detker/Документы/repository/doc-forge/src/converters/csv_engine.py
"""
Движок конвертации CSV / TSV файлов в таблицы Markdown.
"""

import csv
from pathlib import Path
from typing import Dict, Any, List

from src.converters.base import BaseConverter, ConversionResult


class CsvConverter(BaseConverter):
    """
    Конвертер CSV / TSV -> Markdown таблица.
    """

    def can_handle(self, file_path: Path) -> bool:
        return file_path.suffix.lower() in [".csv", ".tsv"]

    def convert(self, file_path: Path, extract_profile: bool = False) -> ConversionResult:
        if not file_path.exists():
            raise FileNotFoundError(f"CSV файл '{file_path}' не найден!")

        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except Exception:
                dialect = csv.excel
                if file_path.suffix.lower() == ".tsv":
                    dialect.delimiter = "\t"

            reader = csv.reader(f, dialect)
            rows = list(reader)

        if not rows:
            return ConversionResult(
                markdown=f"# Таблица: {file_path.name}\n\n*(Файл пуст)*\n",
                metadata={"source_file": str(file_path), "format": "CSV"},
                pages_count=1,
            )

        headers = [h.strip() for h in rows[0]]
        cols = len(headers)
        table_lines = [
            f"# Таблица: {file_path.name}\n",
            "| " + " | ".join(headers) + " |",
            "| " + " | ".join([":---"] * cols) + " |",
        ]

        for r in rows[1:]:
            clean_r = [c.strip().replace("\n", " ").replace("|", "\\|") for c in r]
            padded = clean_r + [""] * (cols - len(clean_r))
            table_lines.append("| " + " | ".join(padded[:cols]) + " |")

        full_md = "\n".join(table_lines) + "\n"

        layout_report = (
            f"📊 ТАБЛИЦА CSV/TSV ({file_path.name}):\n"
            f"   • Разделитель: '{dialect.delimiter}'\n"
            f"   • Колонок: {cols} | Строк: {len(rows)}"
        )

        return ConversionResult(
            markdown=full_md,
            profile=None,
            metadata={
                "source_file": str(file_path),
                "columns_count": cols,
                "rows_count": len(rows),
                "format": "CSV",
            },
            layout_report=layout_report,
            tables_count=1,
            pages_count=1,
        )
