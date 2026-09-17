#!/usr/bin/env python3
"""Remove TC fields used during TOC repair while preserving visible paragraph text."""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def remove_tc_fields(paragraph) -> None:
    children = list(paragraph._p)
    for child in children:
        if child.tag == qn("w:fldSimple") and (child.get(qn("w:instr")) or "").strip().startswith("TC "):
            paragraph._p.remove(child)

    children = list(paragraph._p)
    index = 0
    while index < len(children):
        child = children[index]
        fld_char = child.find(qn("w:fldChar")) if child.tag == qn("w:r") else None
        if fld_char is None or fld_char.get(qn("w:fldCharType")) != "begin":
            index += 1
            continue
        end = index + 1
        is_tc = False
        while end < len(children):
            instr = children[end].find(qn("w:instrText")) if children[end].tag == qn("w:r") else None
            if instr is not None and (instr.text or "").strip().startswith("TC "):
                is_tc = True
            end_char = children[end].find(qn("w:fldChar")) if children[end].tag == qn("w:r") else None
            if end_char is not None and end_char.get(qn("w:fldCharType")) == "end":
                break
            end += 1
        if is_tc and end < len(children):
            for field_child in children[index : end + 1]:
                paragraph._p.remove(field_child)
            children = list(paragraph._p)
            continue
        index = end + 1


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()
    doc = Document(args.docx)
    for paragraph in doc.paragraphs:
        remove_tc_fields(paragraph)
    doc.save(args.docx)


if __name__ == "__main__":
    main()
