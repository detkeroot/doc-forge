# Файл: src/artcraft_bridge.py
"""
Модуль интеграции ArtCraft Suite с Doc-Forge:
- Обнаружение высокопроизводительных нативных бинарников ArtCraft на чистом Rust
- Сверхбыстрая конвертация DOCX -> PDF (~50 мс вместо 3-5 с LibreOffice)
- Мгновенный рендер страниц в PNG для визуального нормоконтроля (wordcraft-cli / pdfcraft-cli)
- Структурный аудит и инспекция документов без графического интерфейса
- Интерактивное открытие документов в ArtCraft (WordCraft, DeckCraft, PdfCraft, CadCraft, FigCraft)
"""

import json
import os
import shutil
import subprocess
from pathlib import Path
from typing import Any, Dict, Optional, Tuple


BUILD_DIR_RELEASE = Path("/var/cache/build_dir/release")
LOCAL_BIN_DIR = Path.home() / ".local" / "bin"


def find_craft_binary(binary_name: str) -> Optional[Path]:
    """
    Находит исполняемый бинарник ArtCraft в системе.
    Приоритет:
    1. /var/cache/build_dir/release/<name> (основной кэш быстрых релизных сборок)
    2. ~/.local/bin/<name> (пользовательские симлинки / лаунчеры)
    3. Системный PATH
    """
    # 1. Проверяем /var/cache/build_dir/release
    candidate_cache = BUILD_DIR_RELEASE / binary_name
    if candidate_cache.exists() and os.access(candidate_cache, os.X_OK):
        return candidate_cache

    # 2. Проверяем ~/.local/bin
    candidate_local = LOCAL_BIN_DIR / binary_name
    if candidate_local.exists() and os.access(candidate_local, os.X_OK):
        return candidate_local

    # 3. Проверяем системный PATH
    system_path = shutil.which(binary_name)
    if system_path:
        p = Path(system_path)
        if p.exists() and os.access(p, os.X_OK):
            return p

    return None


def is_artcraft_available() -> bool:
    """
    Проверяет, доступны ли ключевые утилиты ArtCraft (wordcraft-cli или pdfcraft-cli).
    """
    return bool(find_craft_binary("wordcraft-cli") or find_craft_binary("pdfcraft-cli"))


def _get_craft_env() -> Dict[str, str]:
    """
    Подготавливает окружение для запуска бинарников ArtCraft на NixOS (LD_LIBRARY_PATH и др.).
    """
    env = os.environ.copy()
    cache_env = Path(f"/tmp/artcraft-nix-env-{os.getuid()}.sh")
    if cache_env.exists():
        try:
            with open(cache_env, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("export CRAFT_LD_LIBRARY_PATH="):
                        val = line.split("=", 1)[1].strip("'\"")
                        curr_ld = env.get("LD_LIBRARY_PATH", "")
                        env["LD_LIBRARY_PATH"] = f"{val}:{curr_ld}" if curr_ld else val
        except Exception:
            pass
    return env


def fast_convert_docx_to_pdf(
    docx_path: Path,
    pdf_path: Optional[Path] = None,
    fallback_to_libreoffice: bool = True,
) -> Path:
    """
    Сверхбыстро конвертирует DOCX в PDF с помощью WordCraft (~50 мс).
    При отсутствии wordcraft-cli или ошибке производит автоматический откат к LibreOffice.
    """
    docx_path = Path(docx_path).resolve()
    if not docx_path.exists():
        raise FileNotFoundError(f"Файл '{docx_path}' не найден!")

    out_pdf = Path(pdf_path).resolve() if pdf_path else docx_path.with_suffix(".pdf")
    wordcraft_bin = find_craft_binary("wordcraft-cli")

    if wordcraft_bin:
        try:
            cmd = [str(wordcraft_bin), "convert", str(docx_path), str(out_pdf)]
            res = subprocess.run(cmd, capture_output=True, text=True, env=_get_craft_env(), timeout=10)
            if res.returncode == 0 and out_pdf.exists():
                return out_pdf
        except Exception:
            pass

    # Фоллбэк на LibreOffice, если wordcraft-cli недоступен или выдал ошибку
    if fallback_to_libreoffice:
        out_dir = out_pdf.parent
        cmd_lo = ["libreoffice", "--headless", "--convert-to", "pdf", str(docx_path), "--outdir", str(out_dir)]
        res_lo = subprocess.run(cmd_lo, capture_output=True, text=True, timeout=60)
        if res_lo.returncode != 0:
            raise RuntimeError(f"Ошибка экспорта в PDF через LibreOffice: {res_lo.stderr}")
        generated = out_dir / (docx_path.stem + ".pdf")
        if generated != out_pdf and generated.exists():
            generated.rename(out_pdf)
        return out_pdf

    raise RuntimeError("Не удалось сконвертировать DOCX в PDF через ArtCraft (LibreOffice отключен).")


def render_page_preview(
    file_path: Path,
    page: int = 1,
    out_path: Optional[Path] = None,
    dpi: int = 96,
) -> Path:
    """
    Мгновенно рендерит страницу документа (DOCX, PDF, PPTX) в PNG-изображение.
    Использует wordcraft-cli render, pdfcraft-cli render или deckcraft-cli render.
    """
    file_path = Path(file_path).resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"Файл '{file_path}' не найден!")

    suffix = file_path.suffix.lower()
    out_png = Path(out_path).resolve() if out_path else file_path.with_name(f"{file_path.stem}_p{page}.png")

    env = _get_craft_env()

    if suffix in [".docx", ".doc", ".odt", ".rtf", ".txt", ".html"]:
        wc_bin = find_craft_binary("wordcraft-cli")
        if wc_bin:
            cmd = [str(wc_bin), "render", str(file_path), str(out_png), "--page", str(page)]
            res = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=15)
            if res.returncode == 0 and out_png.exists():
                return out_png

    elif suffix == ".pdf":
        pc_bin = find_craft_binary("pdfcraft-cli")
        if pc_bin:
            cmd = [str(pc_bin), "render", str(file_path), "--page", str(page), "--dpi", str(dpi), "--out", str(out_png)]
            res = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=15)
            if res.returncode == 0 and out_png.exists():
                return out_png

    elif suffix in [".pptx", ".deckcraft"]:
        dc_bin = find_craft_binary("deckcraft-cli")
        if dc_bin:
            cmd = [str(dc_bin), "render", str(file_path), "--slide", str(page), str(out_png)]
            res = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=15)
            if res.returncode == 0 and out_png.exists():
                return out_png

    # Универсальный фоллбэк через poppler (pdftoppm) для PDF
    if suffix == ".pdf":
        pdftoppm = shutil.which("pdftoppm")
        if pdftoppm:
            prefix = out_png.with_suffix("")
            cmd = [pdftoppm, "-png", "-f", str(page), "-l", str(page), "-r", str(dpi), str(file_path), str(prefix)]
            res = subprocess.run(cmd, capture_output=True, text=True)
            candidate = prefix.parent / f"{prefix.name}-{page:02d}.png"
            if candidate.exists():
                candidate.rename(out_png)
                return out_png

    raise RuntimeError(f"Не удалось сгенерировать превью страницы {page} для файла '{file_path.name}'.")


def inspect_document_structure(file_path: Path) -> Dict[str, Any]:
    """
    Выполняет глубокий структурный аудит документа с помощью Rust-движка ArtCraft.
    Возвращает словарь с количеством страниц, слов, абзацев, свойствами и временем верстки (layoutMs).
    """
    file_path = Path(file_path).resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"Файл '{file_path}' не найден!")

    suffix = file_path.suffix.lower()
    env = _get_craft_env()

    if suffix in [".docx", ".doc", ".odt", ".rtf", ".txt", ".html"]:
        wc_bin = find_craft_binary("wordcraft-cli")
        if wc_bin:
            res = subprocess.run([str(wc_bin), "info", str(file_path)], capture_output=True, text=True, env=env)
            if res.returncode == 0:
                try:
                    return json.loads(res.stdout)
                except Exception:
                    pass

    elif suffix == ".pdf":
        pc_bin = find_craft_binary("pdfcraft-cli")
        if pc_bin:
            res = subprocess.run([str(pc_bin), "info", str(file_path)], capture_output=True, text=True, env=env)
            if res.returncode == 0:
                try:
                    return json.loads(res.stdout)
                except Exception:
                    pass

    elif suffix in [".pptx", ".deckcraft"]:
        dc_bin = find_craft_binary("deckcraft-cli")
        if dc_bin:
            res = subprocess.run([str(dc_bin), "info", str(file_path)], capture_output=True, text=True, env=env)
            if res.returncode == 0:
                try:
                    return json.loads(res.stdout)
                except Exception:
                    pass

    return {
        "file": str(file_path.name),
        "file_size": file_path.stat().st_size,
        "format": suffix.lstrip("."),
        "engine": "fallback",
    }


def open_in_artcraft(file_path: Path, control_port: Optional[int] = None) -> subprocess.Popen:
    """
    Интерактивно открывает документ в соответствующем приложении ArtCraft.
    Опционально передает --control <port> для подключения ИИ-агента через MCP.
    """
    file_path = Path(file_path).resolve()
    suffix = file_path.suffix.lower()

    app_map = {
        ".docx": "wordcraft",
        ".doc": "wordcraft",
        ".odt": "wordcraft",
        ".rtf": "wordcraft",
        ".pptx": "deckcraft",
        ".deckcraft": "deckcraft",
        ".pdf": "pdfcraft",
        ".xlsx": "gridcraft",
        ".xls": "gridcraft",
        ".ods": "gridcraft",
        ".csv": "gridcraft",
        ".psd": "photocraft",
        ".svg": "vectorcraft",
        ".fig": "figcraft",
        ".dxf": "cadcraft",
        ".dwg": "cadcraft",
    }

    app_name = app_map.get(suffix, "wordcraft")
    binary_path = find_craft_binary(app_name)
    if not binary_path:
        craft_launcher = find_craft_binary("craft") or shutil.which("craft")
        if craft_launcher:
            cmd = [str(craft_launcher), app_name, str(file_path)]
            if control_port:
                cmd.extend(["--control", str(control_port)])
            return subprocess.Popen(cmd, env=_get_craft_env())
        raise FileNotFoundError(f"Приложение ArtCraft '{app_name}' не найдено.")

    cmd = [str(binary_path), str(file_path)]
    if control_port:
        cmd.extend(["--control", str(control_port)])

    return subprocess.Popen(cmd, env=_get_craft_env())
