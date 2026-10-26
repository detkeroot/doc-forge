# Интеграция ArtCraft Suite в Doc-Forge: План реализации

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Дополнить академический компилятор Doc-Forge нативными Rust-движками ArtCraft Suite (`wordcraft-cli`, `pdfcraft-cli`, `deckcraft-cli`), обеспечив сверхбыструю конвертацию DOCX->PDF (~50 мс вместо 3-5 с LibreOffice), мгновенный рендер превью страниц в PNG для визуального контроля нормоконтроля, структурный аудит документов и интерактивное открытие файлов.

**Architecture:** 
1. Создать легковесный мост `src/artcraft_bridge.py` с автоматическим обнаружением бинарников (`/var/cache/build_dir/release/` и `$PATH`) и прозрачным фоллбэком на LibreOffice/Poppler.
2. Расширить `src/cli.py` новыми командами (`preview`, `audit`, `open`) и обновить пайплайн экспорта PDF в командах `build` и `test` с поддержкой флага быстрого движка `--engine [artcraft|libreoffice]`.
3. Покрыть функционал тестами в `tests/test_artcraft_bridge.py` на базе `unittest` и верифицировать сквозной запуск на эталонном сэмпле `test_samples/sample_4page.md`.

**Tech Stack:** Python 3.12+, Typer, Rich, ArtCraft CLI (`wordcraft-cli`, `pdfcraft-cli`, `deckcraft-cli`), NixOS devShell.

## Global Constraints

- **Не ломать существующий API и контракты:** Все существующие команды `convert`, `inspect`, `build`, `deck`, `test`, `list`, `show`, `set` обязаны работать без регрессий.
- **Graceful Fallback:** Если бинарники ArtCraft отсутствуют (например, на сервере без ArtCraft), doc-forge обязан продолжать работать штатно через LibreOffice, выводя понятное предупреждение без падений.
- **Surgical Minimalism:** Изменения вносятся точечно, без раздувания зависимостей и без лишних абстракций.

---

### Task 1: Создание модуля моста `src/artcraft_bridge.py`

**Files:**
- Create: `src/artcraft_bridge.py`
- Test: `tests/test_artcraft_bridge.py`

**Interfaces:**
- Produces:
  - `find_craft_binary(name: str) -> Optional[Path]`
  - `render_page_preview(file_path: Path, page: int = 1, out_path: Optional[Path] = None, dpi: int = 96) -> Path`
  - `inspect_document_structure(file_path: Path) -> dict`
  - `fast_convert_docx_to_pdf(docx_path: Path, pdf_path: Optional[Path] = None) -> Path`
  - `open_in_artcraft(file_path: Path, control_port: Optional[int] = None) -> None`

- [ ] **Step 1: Написать падающий юнит-тест для обнаружения бинарников и интерфейсов моста**
- [ ] **Step 2: Запустить тест и зафиксировать падение**
- [ ] **Step 3: Реализовать `src/artcraft_bridge.py`**
- [ ] **Step 4: Запустить юнит-тест и подтвердить успешное прохождение**

---

### Task 2: Интеграция быстрого PDF-экспорта и команд в `src/cli.py`

**Files:**
- Modify: `src/cli.py`

**Interfaces:**
- Consumes: функции из `src/artcraft_bridge.py`
- Produces:
  - Команду `doc-forge preview <file> [--page N] [-o out.png]`
  - Команду `doc-forge audit <file>`
  - Команду `doc-forge open <file> [--control PORT]`
  - Расширение команд `build` и `test` поддержкой быстрого движка ArtCraft для PDF

- [ ] **Step 1: Модифицировать функцию `export_docx_to_pdf` в `src/cli.py` с поддержкой `engine="artcraft"` и авто-фоллбэком**
- [ ] **Step 2: Добавить команды `preview`, `audit` и `open` в Typer CLI**
- [ ] **Step 3: Протестировать работу CLI через `nix develop --command python3 src/cli.py --help`**

---

### Task 3: Сквозная верификация и стресс-тест

**Files:**
- Test: сквозной прогон CLI команд на реальных данных `test_samples/sample_4page.md`

- [ ] **Step 1: Сборка эталонного документа `doc-forge test samek --pdf` через движок ArtCraft**
- [ ] **Step 2: Генерация превью страниц через `doc-forge preview` и проверка PNG**
- [ ] **Step 3: Запуск аудита документа через `doc-forge audit` и проверка вывода метрик**
- [ ] **Step 4: Очистка временных файлов тестирования и финальная фиксация в Git**
