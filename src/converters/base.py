# Файл: src/converters/base.py
"""
Базовые интерфейсы и классы данных для подсистемы универсальной конвертации документов.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Any, Optional, List
from src.schema import DocProfile


@dataclass
class ConversionResult:
    """
    Результат прямой конвертации документа любого формата в структурированный Markdown.
    """
    markdown: str
    profile: Optional[DocProfile] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    layout_report: str = ""
    tables_count: int = 0
    images_count: int = 0
    pages_count: int = 1
    headings_hierarchy: List[str] = field(default_factory=list)

    def save_markdown(self, path: Path) -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.markdown)
        return path

    def save_profile_yaml(self, path: Path) -> Optional[Path]:
        if not self.profile:
            return None
        import yaml
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            yaml.dump(self.profile.model_dump(exclude_none=True), f, allow_unicode=True, sort_keys=False)
        return path


class BaseConverter(ABC):
    """
    Абстрактный базовый класс для всех движков конвертации документов в Markdown.
    """

    @abstractmethod
    def can_handle(self, file_path: Path) -> bool:
        """Проверяет, поддерживается ли данный формат файла текущим конвертером."""
        pass

    @abstractmethod
    def convert(self, file_path: Path, extract_profile: bool = True) -> ConversionResult:
        """Выполняет извлечение семантического Markdown и анализ визуального макета."""
        pass
