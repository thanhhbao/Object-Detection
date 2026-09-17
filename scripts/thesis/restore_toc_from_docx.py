#!/usr/bin/env python3
"""Copy a known-good Word TOC field block from one DOCX into another."""

from copy import deepcopy
from pathlib import Path
import sys

from docx import Document
from docx.oxml.ns import qn


def toc_region(document):
    body = document.element.body
    children = list(body.iterchildren())
    start = None
    stop = None
    for index, child in enumerate(children):
        if not child.tag.endswith("}p"):
            continue
        instructions = [node.text or "" for node in child.xpath(".//w:instrText")]
        if any(text.strip().startswith("TOC ") for text in instructions):
            start = child
            for following in children[index + 1 :]:
                if not following.tag.endswith("}p"):
                    stop = following
                    break
                style = following.find("./w:pPr/w:pStyle", namespaces=following.nsmap)
                style_id = style.get(qn("w:val")) if style is not None else ""
                if not style_id.startswith("TOC"):
                    stop = following
                    break
            break
    if start is None or stop is None:
        raise RuntimeError("Không tìm thấy vùng TOC.")
    region = []
    current = start
    while current is not stop:
        region.append(current)
        current = current.getnext()
    return body, region, stop


def main(donor_path: str, target_path: str) -> None:
    donor = Document(Path(donor_path))
    target = Document(Path(target_path))
    _, donor_region, _ = toc_region(donor)
    target_body, target_region, target_stop = toc_region(target)
    for element in target_region:
        target_body.remove(element)
    for element in donor_region:
        target_stop.addprevious(deepcopy(element))
    for field in target.element.xpath(".//w:fldChar[@w:fldCharType='begin']"):
        field.set(qn("w:dirty"), "true")
    target.save(target_path)
    print(f"Đã khôi phục TOC field từ {donor_path} vào {target_path}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
