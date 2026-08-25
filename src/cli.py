# Файл: /home/detker/Документы/repository/doc-forge/src/cli.py
"""
Консольный интерфейс (CLI) комплекса Doc-Forge: компиляция, инспекция и интерактивное управление профилями ГОСТ.
"""

import sys
import yaml
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import print as rprint

from src.schema import DocProfile
from src.compiler import build_document, parse_frontmatter
from src.inspector import inspect_docx, save_profile_to_yaml

app = typer.Typer(
    name="doc-forge",
    help="Doc-Forge: Декларативный компилятор и инспектор академических документов по ГОСТу",
    add_completion=False,
)
console = Console()

REPO_ROOT = Path(__file__).resolve().parent.parent
PROFILES_DIR = REPO_ROOT / "profiles"
SAMPLES_DIR = REPO_ROOT / "test_samples"


def load_profile(name_or_path: str) -> DocProfile:
    """
    Загружает профиль по имени (из папки profiles/) или по прямому пути к YAML файлу.
    """
    p_path = Path(name_or_path)
    if not p_path.exists():
        # Проверяем в profiles/
        candidate = PROFILES_DIR / f"{name_or_path}.yaml"
        if candidate.exists():
            p_path = candidate
        else:
            candidate_yml = PROFILES_DIR / f"{name_or_path}.yml"
            if candidate_yml.exists():
                p_path = candidate_yml
            else:
                raise FileNotFoundError(f"Профиль '{name_or_path}' не найден в {PROFILES_DIR}!")

    with open(p_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    return DocProfile.model_validate(data)


@app.command("build")
def build_cmd(
    input_file: Path = typer.Argument(..., help="Путь к входному .md файлу"),
    profile_name: Optional[str] = typer.Option(None, "-p", "--profile", help="Имя или путь к YAML профилю"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь к результирующему .docx файлу"),
):
    """
    Скомпилировать Markdown документ в стилизованный DOCX по профилю ГОСТ.
    """
    if not input_file.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Файл '{input_file}' не найден!")
        raise typer.Exit(1)

    with open(input_file, "r", encoding="utf-8") as f:
        md_text = f.read()

    fm, _ = parse_frontmatter(md_text)
    
    # Определяем профиль
    chosen_profile_name = profile_name or fm.get("profile") or "samek"
    try:
        profile = load_profile(chosen_profile_name)
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка загрузки профиля:[/bold red] {e}")
        raise typer.Exit(1)

    out_path = output or input_file.with_suffix(".docx")

    console.print(f"[cyan]⚙️ Компиляция:[/cyan] {input_file.name} -> {out_path.name}")
    console.print(f"[cyan]📋 Профиль:[/cyan] [bold green]{profile.meta.name}[/bold green] ({profile.meta.institution})")

    try:
        build_document(md_text, profile, out_path)
        console.print(Panel.fit(
            f"[bold green]✓ Документ успешно создан:[/bold green] {out_path}\n"
            f"[dim]Стандарт: {profile.meta.standard} | Шрифт: {profile.body.font_family} {profile.body.font_size_pt}pt | Интервал: {profile.body.line_spacing}[/dim]",
            title="Doc-Forge Успех",
            border_style="green",
        ))
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка при сборке документа:[/bold red] {e}")
        raise typer.Exit(1)


@app.command("test")
def test_cmd(
    profile_name: str = typer.Argument("samek", help="Имя профиля для генерации теста"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь для сохранения тестового .docx"),
):
    """
    Сгенерировать эталонный 4-страничный тестовый документ (Титульник + Содержание + Глава 1) для проверки верстки.
    """
    sample_file = SAMPLES_DIR / "sample_4page.md"
    if not sample_file.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Тестовый сэмпл '{sample_file}' не найден!")
        raise typer.Exit(1)

    with open(sample_file, "r", encoding="utf-8") as f:
        md_text = f.read()

    try:
        profile = load_profile(profile_name)
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка загрузки профиля:[/bold red] {e}")
        raise typer.Exit(1)

    out_path = output or Path.cwd() / f"test_{profile.meta.name}_4page.docx"

    console.print(Panel(
        f"[bold yellow]📄 Генерация 4-страничного эталонного документа нормоконтроля[/bold yellow]\n"
        f"• Стр. 1: Титульный лист (сетка реквизитов, без нумерации)\n"
        f"• Стр. 2: Содержание (нативная табуляция с отточиями)\n"
        f"• Стр. 3: Введение и Раздел 1.1 (Times New Roman 14pt, 1.5 инт, 1.25 см отступ)\n"
        f"• Стр. 4: Раздел 1.2, Таблица со спецификацией, список с дефисами, Литература",
        title=f"Doc-Forge Test ({profile.meta.name})",
        border_style="yellow",
    ))

    try:
        build_document(md_text, profile, out_path)
        console.print(Panel.fit(
            f"[bold green]✓ 4-страничный тестовый файл создан:[/bold green] [bold]{out_path}[/bold]\n"
            f"[dim]Открой файл в LibreOffice Writer или MS Word для визуальной проверки.[/dim]",
            border_style="green",
        ))
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка генерации теста:[/bold red] {e}")
        raise typer.Exit(1)


@app.command("inspect")
def inspect_cmd(
    docx_file: Path = typer.Argument(..., help="Путь к файлу-образцу .docx для реверс-инжиниринга"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь для сохранения полученного YAML профиля"),
    name: Optional[str] = typer.Option(None, "-n", "--name", help="Имя создаваемого профиля"),
):
    """
    Провести реверс-инжиниринг параметров и стилей образца .docx и сгенерировать YAML профиль.
    """
    if not docx_file.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Файл '{docx_file}' не найден!")
        raise typer.Exit(1)

    prof_name = name or docx_file.stem.lower().replace(" ", "_")
    out_yaml = output or (PROFILES_DIR / f"{prof_name}.yaml")

    console.print(f"[cyan]🔍 Инспекция OpenXML структуры:[/cyan] {docx_file.name}")

    try:
        profile = inspect_docx(docx_file, profile_name=prof_name)
        save_profile_to_yaml(profile, out_yaml)

        table = Table(title=f"Параметры извлеченного профиля: {prof_name}", border_style="cyan")
        table.add_column("Параметр", style="bold yellow")
        table.add_column("Значение", style="green")

        table.add_row("Поля (Л/П/В/Н)", f"{profile.page.margins.left_mm} / {profile.page.margins.right_mm} / {profile.page.margins.top_mm} / {profile.page.margins.bottom_mm} мм")
        table.add_row("Шрифт основного текста", f"{profile.body.font_family} {profile.body.font_size_pt} pt")
        table.add_row("Межстрочный интервал", f"{profile.body.line_spacing}")
        table.add_row("Красная строка (абзац)", f"{profile.body.first_line_indent_cm} см")
        table.add_row("Заголовок 1 уровня (H1)", f"{profile.headings.h1.font_family} {profile.headings.h1.font_size_pt} pt ({profile.headings.h1.alignment})")

        console.print(table)
        console.print(f"[bold green]✓ Профиль сохранен в:[/bold green] {out_yaml}")
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка инспекции:[/bold red] {e}")
        raise typer.Exit(1)


@app.command("list")
def list_cmd():
    """
    Показать список всех доступных профилей оформления в системе.
    """
    yaml_files = list(PROFILES_DIR.glob("*.yaml")) + list(PROFILES_DIR.glob("*.yml"))
    if not yaml_files:
        console.print("[yellow]В папке profiles/ нет профилей.[/yellow]")
        return

    table = Table(title="Доступные профили оформления Doc-Forge", border_style="blue")
    table.add_column("Имя профиля", style="bold green")
    table.add_column("Учреждение", style="white")
    table.add_column("Стандарт", style="yellow")
    table.add_column("Шрифт / Кегль", style="cyan")
    table.add_column("Интервал", style="magenta")
    table.add_column("Поля (Л/П/В/Н)", style="dim white")

    for yf in sorted(yaml_files):
        try:
            with open(yf, "r", encoding="utf-8") as f:
                d = yaml.safe_load(f)
            prof = DocProfile.model_validate(d)
            m = prof.page.margins
            table.add_row(
                prof.meta.name,
                prof.meta.institution,
                prof.meta.standard,
                f"{prof.body.font_family} {prof.body.font_size_pt}pt",
                str(prof.body.line_spacing),
                f"{m.left_mm}/{m.right_mm}/{m.top_mm}/{m.bottom_mm} мм",
            )
        except Exception as e:
            table.add_row(yf.stem, f"[red]Ошибка чтения: {e}[/red]", "-", "-", "-", "-")

    console.print(table)


@app.command("show")
def show_cmd(profile_name: str = typer.Argument(..., help="Имя профиля")):
    """
    Отобразить полную конфигурацию YAML профиля.
    """
    try:
        profile = load_profile(profile_name)
        data = profile.model_dump()
        yaml_str = yaml.dump(data, allow_unicode=True, default_flow_style=False, sort_keys=False)
        console.print(Panel(yaml_str, title=f"Профиль: {profile_name}", border_style="cyan"))
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка:[/bold red] {e}")
        raise typer.Exit(1)


@app.command("set")
def set_cmd(
    profile_name: str = typer.Argument(..., help="Имя профиля"),
    key: str = typer.Argument(..., help="Путь к ключу (например: body.font_size_pt или headings.h1.spacing_after_pt)"),
    value: str = typer.Argument(..., help="Новое значение"),
):
    """
    Интерактивно изменить параметр профиля (например: doc-forge set samek body.font_size_pt 14.0).
    """
    prof_file = PROFILES_DIR / f"{profile_name}.yaml"
    if not prof_file.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Профиль '{prof_file}' не найден!")
        raise typer.Exit(1)

    with open(prof_file, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    # Проходим по вложенным ключам
    keys = key.split(".")
    curr = data
    for k in keys[:-1]:
        if k not in curr or not isinstance(curr[k], dict):
            curr[k] = {}
        curr = curr[k]

    # Преобразуем значение в float/int/bool если возможно
    final_val: Any = value
    if value.lower() == "true":
        final_val = True
    elif value.lower() == "false":
        final_val = False
    else:
        try:
            if "." in value:
                final_val = float(value)
            else:
                final_val = int(value)
        except ValueError:
            final_val = value

    curr[keys[-1]] = final_val

    # Валидируем Pydantic схемой перед сохранением
    try:
        updated_prof = DocProfile.model_validate(data)
        save_profile_to_yaml(updated_prof, prof_file)
        console.print(f"[bold green]✓ Параметр '{key}' успешно обновлен на:[/bold green] [yellow]{final_val}[/yellow]")
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка валидации схемы:[/bold red] {e}")
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
