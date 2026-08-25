# Файл: /home/detker/Документы/repository/doc-forge/src/inspector.py
"""
Модуль инспекции и реверс-инжиниринга документов (DOCX, PDF, etc.) в декларативные YAML-профили ГОСТ.
"""

from pathlib import Path
from typing import Optional, Dict, Any, Tuple
import yaml

from src.schema import DocProfile
from src.converters.router import convert_document
from src.converters.docx_engine import DocxConverter


def inspect_document(file_path: Path, profile_name: Optional[str] = None) -> Tuple[DocProfile, str]:
    """
    Универсальная инспекция документа любого поддерживаемого формата (DOCX, PDF).
    Возвращает сгенерированный объект DocProfile и подробный текстовый отчет о макете.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Файл образца '{file_path}' не найден!")

    result = convert_document(file_path, extract_profile=True)
    if not result.profile:
        raise ValueError(f"Не удалось извлечь профиль оформления из файла '{file_path.name}'!")

    if profile_name:
        result.profile.meta.name = profile_name

    return result.profile, result.layout_report


def inspect_docx(file_path: Path, profile_name: Optional[str] = None) -> DocProfile:
    """
    Обратная совместимость: прямой анализ DOCX файла.
    """
    conv = DocxConverter()
    res = conv.convert(file_path, extract_profile=True)
    if profile_name and res.profile:
        res.profile.meta.name = profile_name
    return res.profile or DocProfile.model_validate({})


def save_profile_to_yaml(profile: DocProfile, output_path: Path):
    """
    Сохраняет профиль в формате YAML с аккуратной структурой.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    data = profile.model_dump(exclude_none=True)
    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
