# Файл: src/converters/epub_engine.py
"""
Движок конвертации электронных книг EPUB (.epub) в структурированный Markdown.
"""

import zipfile
import re
from pathlib import Path
from typing import Dict, Any, List
from bs4 import BeautifulSoup
from lxml import etree

from src.converters.base import BaseConverter, ConversionResult


class EpubConverter(BaseConverter):
    """
    Конвертер EPUB -> Markdown.
    """

    def can_handle(self, file_path: Path) -> bool:
        return file_path.suffix.lower() == ".epub"

    def convert(self, file_path: Path, extract_profile: bool = False) -> ConversionResult:
        if not file_path.exists():
            raise FileNotFoundError(f"EPUB файл '{file_path}' не найден!")

        chapters_md: List[str] = []
        headings_hierarchy: List[str] = []
        book_title = file_path.stem

        with zipfile.ZipFile(str(file_path), "r") as zf:
            # 1. Читаем container.xml для поиска .opf
            try:
                container_data = zf.read("META-INF/container.xml")
                root = etree.fromstring(container_data)
                opf_path = root.xpath("//*[@full-path]/@full-path")[0]
            except Exception:
                opf_path = "content.opf"

            # 2. Читаем manifest и spine из OPF
            try:
                opf_data = zf.read(opf_path)
                opf_root = etree.fromstring(opf_data)
                
                # Ищем заголовок книги
                title_elem = opf_root.xpath("//*[local-name()='title']")
                if title_elem and title_elem[0].text:
                    book_title = title_elem[0].text.strip()

                # Ищем порядок глав в spine
                itemrefs = opf_root.xpath("//*[local-name()='spine']/*[local-name()='itemref']/@idref")
                manifest_items = {}
                for item in opf_root.xpath("//*[local-name()='manifest']/*[local-name()='item']"):
                    manifest_items[item.attrib.get("id")] = item.attrib.get("href")

                opf_dir = Path(opf_path).parent

                # Собираем список файлов глав
                ordered_files = []
                for ref in itemrefs:
                    if ref in manifest_items:
                        full_item_path = (opf_dir / manifest_items[ref]).as_posix().lstrip("/")
                        ordered_files.append(full_item_path)

            except Exception:
                # Fallback: все .xhtml / .html файлы
                ordered_files = [n for n in zf.namelist() if n.endswith((".xhtml", ".html", ".htm")) and "nav" not in n.lower()]

            # 3. Парсим каждую главу
            for ch_file in ordered_files:
                if ch_file not in zf.namelist():
                    continue
                try:
                    ch_html = zf.read(ch_file).decode("utf-8", errors="ignore")
                    soup = BeautifulSoup(ch_html, "html.parser")
                    
                    # Извлекаем контент
                    body = soup.find("body") or soup
                    ch_md = self._html_to_markdown(body)
                    if ch_md.strip():
                        chapters_md.append(ch_md)
                        # Ищем заголовки
                        for h in body.find_all(["h1", "h2", "h3"]):
                            headings_hierarchy.append(f"{h.name.upper()}: {h.get_text().strip()}")
                except Exception:
                    pass

        full_md = f"# Книга: {book_title}\n\n" + "\n\n---\n\n".join(chapters_md)
        full_md = re.sub(r'\n{3,}', '\n\n', full_md)

        layout_report = (
            f"📚 ЭЛЕКТРОННАЯ КНИГА EPUB ({file_path.name}):\n"
            f"   • Название: {book_title}\n"
            f"   • Количество глав/разделов: {len(chapters_md)}\n"
            f"   • Заголовков в структуре: {len(headings_hierarchy)}"
        )

        return ConversionResult(
            markdown=full_md,
            profile=None,
            metadata={
                "source_file": str(file_path),
                "book_title": book_title,
                "chapters_count": len(chapters_md),
                "format": "EPUB",
            },
            layout_report=layout_report,
            pages_count=len(chapters_md),
            headings_hierarchy=headings_hierarchy,
        )

    def _html_to_markdown(self, soup_elem) -> str:
        """
        Преобразует HTML элементы в структурированный Markdown.
        """
        md_parts: List[str] = []

        for elem in soup_elem.children:
            if isinstance(elem, str):
                t = elem.strip()
                if t:
                    md_parts.append(t)
                continue

            tag = elem.name
            text = elem.get_text().strip()
            if not text and tag != "hr":
                continue

            if tag == "h1":
                md_parts.append(f"\n# {text}\n")
            elif tag == "h2":
                md_parts.append(f"\n## {text}\n")
            elif tag == "h3":
                md_parts.append(f"\n### {text}\n")
            elif tag == "h4" or tag == "h5" or tag == "h6":
                md_parts.append(f"\n#### {text}\n")
            elif tag == "p":
                # Обработка внутренних тегов (strong, em, code)
                formatted = self._format_inline_html(elem)
                md_parts.append(formatted)
            elif tag in ["ul", "ol"]:
                for li in elem.find_all("li", recursive=False):
                    li_text = li.get_text().strip()
                    md_parts.append(f"- {li_text}")
            elif tag == "blockquote":
                md_parts.append(f"> {text}")
            elif tag == "pre" or tag == "code":
                md_parts.append(f"```\n{text}\n```")
            elif tag == "hr":
                md_parts.append("\n---\n")
            else:
                formatted = self._format_inline_html(elem)
                if formatted:
                    md_parts.append(formatted)

        return "\n\n".join(md_parts)

    def _format_inline_html(self, elem) -> str:
        res = []
        for child in elem.contents:
            if isinstance(child, str):
                res.append(child)
            elif child.name in ["strong", "b"]:
                res.append(f"**{child.get_text()}**")
            elif child.name in ["em", "i"]:
                res.append(f"*{child.get_text()}*")
            elif child.name == "code":
                res.append(f"`{child.get_text()}`")
            else:
                res.append(child.get_text())
        return "".join(res).strip()
