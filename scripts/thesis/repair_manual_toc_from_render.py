#!/usr/bin/env python3
"""Repair page numbers in the thesis' manual TOC/list paragraphs from a rendered PDF."""

from __future__ import annotations

import argparse
import re
import subprocess
import tempfile
from pathlib import Path

from docx import Document


FRONT_PAGES = {
    "LỜI MỞ ĐẦU": "i",
    "LỜI CẢM ƠN": "iii",
    "NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN": "iv",
    "NHẬN XÉT CỦA GIẢNG VIÊN PHẢN BIỆN": "v",
    "MỤC LỤC": "vi",
    "DANH MỤC BẢNG BIỂU": "x",
    "DANH MỤC HÌNH": "xi",
    "DANH MỤC CÁC TỪ VIẾT TẮT": "xii",
}

ORPHAN_ENTRIES = {"Phụ lục D. Mã nguồn chính"}


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip().casefold()


def title_without_page(text: str) -> str:
    if "\t" in text:
        left, right = text.rsplit("\t", 1)
        if re.fullmatch(r"[ivxlcdm]+|\d+", right.strip(), re.I):
            return left.strip()
    return text.strip()


def extract_pages(pdf_path: Path) -> dict[int, str]:
    info = subprocess.run(["pdfinfo", str(pdf_path)], check=True, capture_output=True, text=True).stdout
    page_count = int(re.search(r"^Pages:\s+(\d+)", info, re.M).group(1))
    pages: dict[int, str] = {}
    with tempfile.TemporaryDirectory(prefix="kltn_toc_pages_") as temp_dir:
        for page in range(1, page_count + 1):
            output = Path(temp_dir) / f"page-{page}.txt"
            subprocess.run(
                ["pdftotext", "-f", str(page), "-l", str(page), "-layout", str(pdf_path), str(output)],
                check=True,
            )
            pages[page] = output.read_text(errors="ignore")
    return pages


def find_page(title: str, pages: dict[int, str]) -> str | None:
    if title in FRONT_PAGES:
        return FRONT_PAGES[title]
    needle = normalize(title)
    # Prefer an exact rendered line so short headings such as "PHỤ LỤC" are not
    # mistaken for an earlier narrative mention.
    for physical_page in range(15, max(pages) + 1):
        lines = [normalize(line) for line in pages[physical_page].splitlines() if normalize(line)]
        if needle in lines:
            return str(physical_page - 14)
    for physical_page in range(15, max(pages) + 1):
        if needle in normalize(pages[physical_page]):
            return str(physical_page - 14)
    return None


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    doc = Document(args.docx)
    pages = extract_pages(args.pdf)
    candidates = [p for p in doc.paragraphs if p.style.name.lower().startswith("toc") and p.text.strip()]
    for paragraph in list(candidates):
        if title_without_page(paragraph.text) in ORPHAN_ENTRIES:
            print(f"REMOVE ORPHAN\t{title_without_page(paragraph.text)}")
            candidates.remove(paragraph)
            if not args.dry_run:
                paragraph._element.getparent().remove(paragraph._element)
    unresolved: list[str] = []
    updates: list[tuple[object, str, str]] = []
    for paragraph in candidates:
        title = title_without_page(paragraph.text)
        page = find_page(title, pages)
        if page is None:
            unresolved.append(title)
        else:
            updates.append((paragraph, title, page))

    if unresolved:
        raise RuntimeError("Unresolved manual-list entries:\n- " + "\n- ".join(unresolved))
    for paragraph, title, page in updates:
        print(f"{title}\t{page}")
        if not args.dry_run:
            paragraph.text = f"{title}\t{page}"
    if not args.dry_run:
        doc.save(args.docx)
    print(f"Updated {len(updates)} manual-list entries")


if __name__ == "__main__":
    main()
