# Файл: /home/detker/Документы/repository/doc-forge/src/cli.py
"""
CLI интерфейс для Doc-Forge: компиляция, инспекция и тестирование академических документов.
"""

import sys
from pathlib import Path
from typing import Optional
import yaml
import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel

from src.schema import DocProfile
from src.compiler import build_document
from src.inspector import inspect_docx, export_profile_yaml

app = typer.Typer(
    help="🚀 Doc-Forge: Декларативный компилятор и инспектор документов по ГОСТу",
    add_completion=False
)
console = Console()

BASE_DIR = Path(__file__).resolve().parent.parent
PROFILES_DIR = BASE_DIR / "profiles"


def load_profile(name_or_path: str) -> DocProfile:
    """Загружает профиль по имени из папки profiles/ или по прямому пути к .yaml файлу."""
    path = Path(name_or_path)
    if not path.exists():
        path = PROFILES_DIR / f"{name_or_path}.yaml"
    if not path.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Профиль '{name_or_path}' не найден!")
        sys.exit(1)

    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return DocProfile.model_validate(data)


@app.command(name="test")
def test_cmd(
    profile_name: str = typer.Argument("samek", help="Имя профиля оформления (samek, gost_7_32, samgtu)"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь к результирующему .docx файлу")
):
    """
    Генерирует эталонный 4-страничный тестовый документ (Титульник -> Содержание -> Введение/Гл1 -> Таблица/Гл1/Источники).
    """
    console.print(Panel(f"[bold green]Генерация 4-страничного тестового документа[/bold green]\nПрофиль: [cyan]{profile_name}[/cyan]", expand=False))
    profile = load_profile(profile_name)

    sample_md_path = BASE_DIR / "test_samples" / "sample_4page.md"
    if not sample_md_path.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Файл примера {sample_md_path} не найден!")
        sys.exit(1)

    with open(sample_md_path, "r", encoding="utf-8") as f:
        md_content = f.read()

    out_file = output or (BASE_DIR / "output" / f"test_4page_{profile_name}.docx")
    out_file = build_document(md_content, profile, out_file)

    console.print(f"✨ [bold green]Успешно создан 4-страничный документ:[/bold green] [cyan]{out_file}[/cyan]")
    console.print("📄 [bold]Структура страниц:[/bold]")
    console.print("  • [yellow]Стр. 1:[/yellow] Титульный лист (Министерство, Колледж, Специальность, Тема, Подписи, Город, Год)")
    console.print("  • [yellow]Стр. 2:[/yellow] Содержание (Нативное авто-оглавление с точками-отточиями до правого края)")
    console.print("  • [yellow]Стр. 3:[/yellow] Введение + 1 Аналитический раздел (1.1 Текст по ГОСТ, 1.25 см отступ, 1.5 интервал)")
    console.print("  • [yellow]Стр. 4:[/yellow] Продолжение раздела 1 (1.2 Спецификация + Таблица + Маркированный список + Источники)")


@app.command(name="build")
def build_cmd(
    input_file: Path = typer.Argument(..., help="Входной .md файл с текстом работы"),
    profile_name: str = typer.Option("samek", "-p", "--profile", help="Имя профиля оформления"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь к выходному .docx")
):
    """
    Компилирует Markdown в готовый стилизованный DOCX по правилам профиля.
    """
    if not input_file.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Файл {input_file} не найден!")
        sys.exit(1)

    with open(input_file, "r", encoding="utf-8") as f:
        md_content = f.read()

    profile = load_profile(profile_name)
    out_file = output or input_file.with_suffix(".docx")
    out_file = build_document(md_content, profile, out_file)
    console.print(f"✅ [bold green]Документ скомпилирован:[/bold green] [cyan]{out_file}[/cyan]")


@app.command(name="inspect")
def inspect_cmd(
    sample_docx: Path = typer.Argument(..., help="Путь к образцовому .docx документу"),
    output_yaml: Optional[Path] = typer.Option(None, "-o", "--output", help="Куда сохранить YAML профиль"),
    name: str = typer.Option("custom_profile", "--name", help="Имя нового профиля")
):
    """
    Анализирует существующий .docx (методичку/образец) и генерирует YAML профиль.
    """
    if not sample_docx.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Файл {sample_docx} не найден!")
        sys.exit(1)

    console.print(f"🔍 [bold cyan]Анализируем образец:[/bold cyan] {sample_docx.name}...")
    profile = inspect_docx(sample_docx, profile_name=name)
    
    out_path = output_yaml or (PROFILES_DIR / f"{name}.yaml")
    export_profile_yaml(profile, out_path)

    console.print(f"✅ [bold green]Профиль успешно сгенерирован:[/bold green] [cyan]{out_path}[/cyan]")
    show_profile_table(profile)


@app.command(name="list")
def list_cmd():
    """
    Выводит список всех доступных профилей оформления.
    """
    table = Table(title="📋 Доступные профили оформления Doc-Forge")
    table.add_column("Имя профиля", style="cyan", no_wrap=True)
    table.add_column("Учебное заведение", style="green")
    table.add_column("Стандарт", style="magenta")
    table.add_column("Описание", style="white")

    for p in PROFILES_DIR.glob("*.yaml"):
        try:
            with open(p, "r", encoding="utf-8") as f:
                d = yaml.safe_load(f)
            meta = d.get("meta", {})
            table.add_row(
                p.stem,
                meta.get("institution", "—"),
                meta.get("standard", "—"),
                meta.get("description", "—")
            )
        except Exception:
            continue

    console.print(table)


@app.command(name="show")
def show_cmd(profile_name: str = typer.Argument(..., help="Имя профиля")):
    """
    Детально отображает все метрики выбранного профиля.
    """
    profile = load_profile(profile_name)
    show_profile_table(profile)


def show_profile_table(profile: DocProfile):
    """Выводит сводную таблицу параметров профиля."""
    table = Table(title=f"📐 Параметры профиля: {profile.meta.name} ({profile.meta.standard})")
    table.add_column("Параметр", style="cyan")
    table.add_column("Значение", style="green")

    table.add_row("Поля (Левое/Правое/Верх/Низ)", f"{profile.page.margins.left_mm} / {profile.page.margins.right_mm} / {profile.page.margins.top_mm} / {profile.page.margins.bottom_mm} мм")
    table.add_row("Основной шрифт", f"{profile.body.font_family} {profile.body.font_size_pt}pt, интервал {profile.body.line_spacing}")
    table.add_row("Красная строка", f"{profile.body.first_line_indent_cm} см")
    table.add_row("Выравнивание", profile.body.alignment)
    table.add_row("Заголовок 1 (Главы)", f"{profile.headings.h1.font_size_pt}pt, bold={profile.headings.h1.bold}, капс={profile.headings.h1.all_caps}, отступы {profile.headings.h1.spacing_before_pt}/{profile.headings.h1.spacing_after_pt}pt")
    table.add_row("Заголовок 2 (Подразделы)", f"{profile.headings.h2.font_size_pt}pt, bold={profile.headings.h2.bold}, отступ {profile.headings.h2.first_line_indent_cm} см")
    table.add_row("Таблицы", f"шрифт {profile.elements.tables.font_size_pt}pt, интервал {profile.elements.tables.line_spacing}, повтор шапки={profile.elements.tables.repeat_header_on_new_page}")
    table.add_row("Списки", f"стиль={profile.elements.lists.bullet_style}, отступ={profile.elements.lists.first_line_indent_cm} см")

    console.print(table)


if __name__ == "__main__":
    app()
