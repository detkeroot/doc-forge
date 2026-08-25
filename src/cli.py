# Файл: /home/detker/Документы/repository/doc-forge/src/cli.py
"""
Консольный интерфейс (CLI) комплекса Doc-Forge:
- Двусторонняя конвертация (Все форматы <-> Markdown)
- Инспекция и реверс-инжиниринг макетов документов (DOCX, PDF) в YAML-профили
- Декларативная компиляция в DOCX / PDF по стандартам ГОСТ / СТО
"""

import sys
import subprocess
from pathlib import Path
from typing import Optional

import typer
import yaml
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import print as rprint

from src.schema import DocProfile
from src.compiler import build_document, parse_frontmatter
from src.inspector import inspect_document, save_profile_to_yaml
from src.converters.router import convert_document

app = typer.Typer(
    name="doc-forge",
    help="Doc-Forge: Универсальный двусторонний конвертер и инспектор документов по ГОСТу",
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


def export_docx_to_pdf(docx_path: Path, output_pdf: Optional[Path] = None) -> Path:
    """
    Экспортирует сгенерированный DOCX в PDF с помощью LibreOffice в headless режиме.
    """
    out_dir = output_pdf.parent if output_pdf else docx_path.parent
    cmd = ["libreoffice", "--headless", "--convert-to", "pdf", str(docx_path), "--outdir", str(out_dir)]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"Ошибка экспорта в PDF через LibreOffice: {res.stderr}")
    
    generated_pdf = out_dir / (docx_path.stem + ".pdf")
    if output_pdf and output_pdf != generated_pdf and generated_pdf.exists():
        generated_pdf.rename(output_pdf)
        return output_pdf
    return generated_pdf


@app.command("convert")
def convert_cmd(
    input_file: Path = typer.Argument(..., help="Путь к документу (PDF, DOCX, XLSX, PPTX, RTF, EPUB, CSV)"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь для сохранения выходного .md файла"),
    profile_out: Optional[Path] = typer.Option(None, "-p", "--profile-out", help="Путь для сохранения извлеченного YAML профиля"),
    report: bool = typer.Option(True, "--report/--no-report", help="Выводить ли отчет о визуальной структуре макета"),
):
    """
    [ПРЯМОЙ ХОД] Конвертировать документ любого формата в структурированный Rich Markdown + YAML профиль.
    """
    if not input_file.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Файл '{input_file}' не найден!")
        raise typer.Exit(1)

    console.print(f"[bold cyan]🚀 Прямая конвертация:[/bold cyan] [yellow]{input_file.name}[/yellow] -> Markdown")

    try:
        result = convert_document(input_file, extract_profile=True)
        
        # Определяем путь для сохранения MD
        out_md_path = output or input_file.with_suffix(".md")
        result.save_markdown(out_md_path)

        # Сохраняем профиль при наличии
        saved_profile_path = None
        if result.profile:
            p_out = profile_out or (PROFILES_DIR / f"{input_file.stem.lower().replace(' ', '_')}.yaml")
            saved_profile_path = result.save_profile_yaml(p_out)

        # Выводим визуальный отчет
        if report and result.layout_report:
            console.print(Panel(result.layout_report, title=f"📊 Анализ макета: {input_file.name}", border_style="cyan"))

        summary_lines = [
            f"[bold green]✓ Markdown сохранен:[/bold green] [bold]{out_md_path}[/bold]",
            f"[dim]Формат: {result.metadata.get('format')} | Таблиц: {result.tables_count} | Страниц/Слайдов: {result.pages_count}[/dim]",
        ]
        if saved_profile_path:
            summary_lines.append(f"[bold yellow]✓ Извлеченный YAML-профиль:[/bold yellow] [bold]{saved_profile_path}[/bold]")

        console.print(Panel.fit("\n".join(summary_lines), title="Doc-Forge Успех", border_style="green"))

    except Exception as e:
        console.print(f"[bold red]❌ Ошибка при конвертации документа:[/bold red] {e}")
        import traceback
        console.print(f"[dim]{traceback.format_exc()}[/dim]")
        raise typer.Exit(1)


@app.command("inspect")
def inspect_cmd(
    input_file: Path = typer.Argument(..., help="Путь к файлу-образцу (DOCX, PDF) для реверс-инжиниринга"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь для сохранения полученного YAML профиля"),
    name: Optional[str] = typer.Option(None, "-n", "--name", help="Имя создаваемого профиля"),
    markdown_out: Optional[Path] = typer.Option(None, "-m", "--markdown-out", help="Сохранить также извлеченный Markdown"),
):
    """
    [РЕВЕРС-ИНЖИНИРИНГ] Глубокий анализ геометрии, шрифтов, отступов и титульного листа образца (DOCX, PDF).
    """
    if not input_file.exists():
        console.print(f"[bold red]❌ Ошибка:[/bold red] Файл '{input_file}' не найден!")
        raise typer.Exit(1)

    prof_name = name or input_file.stem.lower().replace(" ", "_")
    out_yaml = output or (PROFILES_DIR / f"{prof_name}.yaml")

    console.print(f"[cyan]🔍 Глубокая инспекция макета:[/cyan] {input_file.name}")

    try:
        profile, layout_report = inspect_document(input_file, profile_name=prof_name)
        save_profile_to_yaml(profile, out_yaml)

        # Выводим подробный отчет
        console.print(Panel(layout_report, title=f"Отчет инспекции макета: {input_file.name}", border_style="cyan"))

        table = Table(title=f"Ключевые параметры профиля: {prof_name}", border_style="green")
        table.add_column("Параметр", style="bold yellow")
        table.add_column("Значение", style="green")

        table.add_row("Поля (Л/П/В/Н)", f"{profile.page.margins.left_mm} / {profile.page.margins.right_mm} / {profile.page.margins.top_mm} / {profile.page.margins.bottom_mm} мм")
        table.add_row("Шрифт основного текста", f"{profile.body.font_family} {profile.body.font_size_pt} pt")
        table.add_row("Межстрочный интервал", f"{profile.body.line_spacing}")
        table.add_row("Красная строка (абзац)", f"{profile.body.first_line_indent_cm} см")
        table.add_row("Выравнивание", f"{profile.body.alignment}")
        table.add_row("Заголовок 1 уровня (H1)", f"{profile.headings.h1.font_family} {profile.headings.h1.font_size_pt} pt ({profile.headings.h1.alignment})")
        table.add_row("Заголовок 2 уровня (H2)", f"{profile.headings.h2.font_family} {profile.headings.h2.font_size_pt} pt ({profile.headings.h2.alignment})")

        console.print(table)
        console.print(f"[bold green]✓ Профиль успешно сохранен в:[/bold green] {out_yaml}")

        if markdown_out:
            res = convert_document(input_file, extract_profile=False)
            res.save_markdown(markdown_out)
            console.print(f"[bold green]✓ Семантический Markdown сохранен в:[/bold green] {markdown_out}")

    except Exception as e:
        console.print(f"[bold red]❌ Ошибка инспекции документа:[/bold red] {e}")
        import traceback
        console.print(f"[dim]{traceback.format_exc()}[/dim]")
        raise typer.Exit(1)


@app.command("build")
def build_cmd(
    input_file: Path = typer.Argument(..., help="Путь к входному .md файлу"),
    profile_name: Optional[str] = typer.Option(None, "-p", "--profile", help="Имя или путь к YAML профилю"),
    output: Optional[Path] = typer.Option(None, "-o", "--output", help="Путь к результирующему .docx файлу"),
    pdf: bool = typer.Option(False, "--pdf", help="Автоматически экспортировать результат в PDF через LibreOffice"),
):
    """
    [ОБРАТНЫЙ ХОД] Скомпилировать Markdown документ в стилизованный DOCX (и PDF) по профилю ГОСТ.
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
        messages = [
            f"[bold green]✓ Документ DOCX успешно создан:[/bold green] {out_path}",
            f"[dim]Стандарт: {profile.meta.standard} | Шрифт: {profile.body.font_family} {profile.body.font_size_pt}pt | Интервал: {profile.body.line_spacing}[/dim]",
        ]

        if pdf:
            pdf_path = out_path.with_suffix(".pdf")
            console.print(f"[cyan]🖨️ Экспорт в PDF (LibreOffice Headless):[/cyan] {pdf_path.name}")
            export_docx_to_pdf(out_path, pdf_path)
            messages.append(f"[bold green]✓ Документ PDF успешно создан:[/bold green] {pdf_path}")

        console.print(Panel.fit(
            "\n".join(messages),
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
    pdf: bool = typer.Option(False, "--pdf", help="Автоматически сгенерировать также PDF"),
):
    """
    Сгенерировать эталонный 4-страничный тестовый документ (Титульник + Содержание + Введение + Разд. 1.1 + Разд. 1.2) для проверки верстки.
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
        messages = [f"[bold green]✓ 4-страничный тестовый DOCX создан:[/bold green] [bold]{out_path}[/bold]"]
        
        if pdf:
            pdf_path = out_path.with_suffix(".pdf")
            export_docx_to_pdf(out_path, pdf_path)
            messages.append(f"[bold green]✓ Тестовый PDF создан:[/bold green] [bold]{pdf_path}[/bold]")

        console.print(Panel.fit(
            "\n".join(messages) + "\n[dim]Открой файл в LibreOffice Writer или MS Word для визуальной проверки.[/dim]",
            border_style="green",
        ))
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка генерации теста:[/bold red] {e}")
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
                data = yaml.safe_load(f)
            p = DocProfile.model_validate(data)
            margins_str = f"{p.page.margins.left_mm}/{p.page.margins.right_mm}/{p.page.margins.top_mm}/{p.page.margins.bottom_mm}"
            table.add_row(
                p.meta.name,
                p.meta.institution,
                p.meta.standard,
                f"{p.body.font_family} {p.body.font_size_pt}pt",
                f"{p.body.line_spacing}",
                margins_str,
            )
        except Exception as e:
            table.add_row(yf.stem, f"[red]Ошибка чтения: {e}[/red]", "-", "-", "-", "-")

    console.print(table)


@app.command("show")
def show_cmd(profile_name: str = typer.Argument(..., help="Имя профиля для просмотра")):
    """
    Показать детальные параметры конкретного профиля оформления.
    """
    try:
        p = load_profile(profile_name)
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка:[/bold red] {e}")
        raise typer.Exit(1)

    console.print(Panel.fit(
        f"[bold cyan]Профиль:[/bold cyan] {p.meta.name}\n"
        f"[bold cyan]Учреждение:[/bold cyan] {p.meta.institution}\n"
        f"[bold cyan]Стандарт:[/bold cyan] {p.meta.standard}\n"
        f"[dim]{p.meta.description or ''}[/dim]",
        title="Метаданные",
        border_style="cyan",
    ))

    t_body = Table(title="Параметры основного текста", border_style="yellow")
    t_body.add_column("Свойство", style="bold")
    t_body.add_column("Значение", style="green")
    t_body.add_row("Шрифт", p.body.font_family)
    t_body.add_row("Кегль", f"{p.body.font_size_pt} pt")
    t_body.add_row("Межстрочный интервал", str(p.body.line_spacing))
    t_body.add_row("Абзацный отступ (красная строка)", f"{p.body.first_line_indent_cm} см")
    t_body.add_row("Выравнивание", p.body.alignment)
    t_body.add_row("Интервалы До / После", f"{p.body.spacing_before_pt} pt / {p.body.spacing_after_pt} pt")
    console.print(t_body)


@app.command("set")
def set_cmd(
    profile_name: str = typer.Argument(..., help="Имя профиля для модификации"),
    field_path: str = typer.Argument(..., help="Путь к полю (например: body.font_size_pt или page.margins.left_mm)"),
    value: str = typer.Argument(..., help="Новое значение"),
):
    """
    Интерактивно изменить один параметр в YAML-профиле.
    """
    target_file = PROFILES_DIR / f"{profile_name}.yaml"
    if not target_file.exists():
        target_file = PROFILES_DIR / f"{profile_name}.yml"
        if not target_file.exists():
            console.print(f"[bold red]❌ Ошибка:[/bold red] Профиль '{profile_name}' не найден в {PROFILES_DIR}!")
            raise typer.Exit(1)

    with open(target_file, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    # Приводим типы
    val_parsed: Any = value
    if value.lower() == "true":
        val_parsed = True
    elif value.lower() == "false":
        val_parsed = False
    elif value.lower() == "null" or value.lower() == "none":
        val_parsed = None
    else:
        try:
            if "." in value:
                val_parsed = float(value)
            else:
                val_parsed = int(value)
        except ValueError:
            val_parsed = value

    # Проходим по пути к ключу (например, body.font_size_pt)
    keys = field_path.split(".")
    curr = data
    for k in keys[:-1]:
        if k not in curr or not isinstance(curr[k], dict):
            curr[k] = {}
        curr = curr[k]

    curr[keys[-1]] = val_parsed

    # Валидируем через Pydantic
    try:
        updated_profile = DocProfile.model_validate(data)
        save_profile_to_yaml(updated_profile, target_file)
        console.print(f"[bold green]✓ Поле '{field_path}' успешно обновлено на '{val_parsed}' в профиле '{profile_name}'![/bold green]")
    except Exception as e:
        console.print(f"[bold red]❌ Ошибка валидации нового значения:[/bold red] {e}")
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
