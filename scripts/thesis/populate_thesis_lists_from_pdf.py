#!/usr/bin/env python3
"""Populate cached TOC/list page numbers from a rendered thesis PDF.

Word's TOC field remains marked dirty for future refreshes. This script makes
the document immediately readable even when Word/LibreOffice does not refresh
the cached field while running headlessly.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import tempfile
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Cm, Pt


def normalized(text: str) -> str:
    text = unicodedata.normalize("NFC", text).replace("–", "-").replace("−", "-")
    return re.sub(r"\s+", " ", text).strip().casefold()


def pdf_pages(pdf: Path) -> list[str]:
    with tempfile.NamedTemporaryFile(suffix=".txt") as tmp:
        subprocess.run(["pdftotext", "-layout", str(pdf), tmp.name], check=True)
        raw = Path(tmp.name).read_text(errors="ignore")
    return [normalized(page) for page in raw.split("\f")]


def find_page(pages: list[str], text: str, start_at: int = 1) -> int | None:
    needle = normalized(text)
    if needle == "phụ lục":
        for idx, page in enumerate(pages[start_at - 1 :], start_at):
            if "phụ lục a." in page:
                return idx
    if needle == "tài liệu tham khảo":
        for idx in range(len(pages), start_at - 1, -1):
            if needle in pages[idx - 1]:
                return idx
    for idx, page in enumerate(pages[start_at - 1 :], start_at):
        if needle in page:
            return idx
    # Wrapped/ligature-heavy fallback based on a distinctive prefix.
    prefix = " ".join(needle.split()[:8])
    for idx, page in enumerate(pages[start_at - 1 :], start_at):
        if prefix and prefix in page:
            return idx
    return None


def configure_entry(paragraph: Paragraph, style, text: str, page: int) -> None:
    paragraph.text = f"{text}\t{page}"
    paragraph.style = style
    fmt = paragraph.paragraph_format
    fmt.tab_stops.clear_all()
    fmt.tab_stops.add_tab_stop(Cm(15.45), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    fmt.keep_together = True
    fmt.widow_control = True
    for run in paragraph.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(13)


def paragraph_after(paragraph: Paragraph) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)


def clear_field_shell(paragraph: Paragraph) -> None:
    # Retain the paragraph itself but remove field runs/bookmarks and placeholder.
    for child in list(paragraph._p):
        if child.tag.endswith("}pPr"):
            continue
        paragraph._p.remove(child)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    pages = pdf_pages(args.pdf)
    doc = Document(args.docx)
    toc_styles = {
        1: doc.styles["toc 1"],
        2: doc.styles["toc 2"],
        3: doc.styles["toc 3"],
    }
    body_heading = next(p for p in doc.paragraphs if p.text.strip().startswith("CHƯƠNG 1"))
    body_physical_page = find_page(pages, body_heading.text)
    if body_physical_page is None:
        raise RuntimeError("Could not locate Chapter 1 in rendered PDF")

    toc_items: list[tuple[str, str, int]] = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        style = paragraph.style.name
        if style not in {"Heading 1", "Heading 2", "Heading 3"} or not text:
            continue
        physical = find_page(pages, text, body_physical_page)
        if physical is None:
            raise RuntimeError(f"Could not locate TOC heading in PDF: {text}")
        printed = physical - body_physical_page + 1
        toc_style = toc_styles[{"Heading 1": 1, "Heading 2": 2, "Heading 3": 3}[style]]
        toc_items.append((toc_style, text, printed))

    placeholder = next(p for p in doc.paragraphs if "Đang cập nhật mục lục" in p.text)
    clear_field_shell(placeholder)
    current = placeholder
    for idx, (style, text, page) in enumerate(toc_items):
        if idx:
            current = paragraph_after(current)
        configure_entry(current, style, text, page)

    misses: list[str] = []
    in_figures = False
    in_tables = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text == "DANH MỤC HÌNH":
            in_figures, in_tables = True, False
            continue
        if text == "DANH MỤC BẢNG":
            in_figures, in_tables = False, True
            continue
        if text.startswith("CHƯƠNG 1"):
            in_figures = in_tables = False
        if not ((in_figures and text.startswith("Hình ")) or (in_tables and text.startswith("Bảng "))):
            continue
        physical = find_page(pages, text, body_physical_page)
        if physical is None:
            misses.append(text)
            continue
        printed = physical - body_physical_page + 1
        configure_entry(paragraph, toc_styles[1], text, printed)

    if misses:
        raise RuntimeError("Could not locate list entries in PDF:\n" + "\n".join(misses))

    doc.save(args.output)
    print(f"TOC entries: {len(toc_items)}; lists populated; body starts on physical page {body_physical_page}")


if __name__ == "__main__":
    main()
