#!/usr/bin/env python3
"""Reduce unnecessary English wording in Chapter 4 without changing metrics."""

from pathlib import Path
import sys

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from rewrite_thesis_chapter3 import mark_fields_dirty, update_manual_figure_table_lists


REPLACEMENTS = [
    ("một số run chính", "một số lượt huấn luyện chính"),
    ("Run copy-paste", "Lần huấn luyện copy-paste"),
    ("run copy-paste", "lần huấn luyện copy-paste"),
    ("Ablation SE", "Thí nghiệm tách thành phần SE"),
    ("Benchmark cho thấy", "Thực nghiệm so sánh cho thấy"),
    ("khả năng zero-shot", "khả năng suy luận khi chưa thích nghi"),
    ("cải thiện zero-shot thuần túy", "cải thiện trong điều kiện chưa thích nghi"),
    ("split", "phân hoạch"),
    ("metadata", "siêu dữ liệu"),
    ("chọn checkpoint", "chọn bộ trọng số"),
    ("chuỗi checkpoint", "chuỗi bộ trọng số"),
    ("Full text log", "Nhật ký văn bản đầy đủ"),
    ("full text log", "nhật ký văn bản đầy đủ"),
    ("checkpoint best.pt", "tệp trọng số best.pt"),
    ("Checkpoint best.pt", "Tệp trọng số best.pt"),
    ("checkpoint final", "bộ trọng số cuối cùng"),
    ("Checkpoint final", "Bộ trọng số cuối cùng"),
    ("mô hình final", "mô hình cuối cùng"),
    ("cấu hình final", "cấu hình cuối cùng"),
    ("best.pt final", "best.pt cuối cùng"),
    ("quan hệ checkpoint", "quan hệ kế thừa trọng số"),
    ("tái đánh giá final", "tái đánh giá mô hình cuối cùng"),
    ("held-out", "phần dữ liệu giữ lại"),
    ("provenance", "bằng chứng truy vết"),
    ("benchmark", "thực nghiệm so sánh"),
    ("validation", "kiểm định"),
    ("các run", "các lần huấn luyện"),
    ("mỗi run", "mỗi lần huấn luyện"),
    ("run v3_960", "lần huấn luyện v3_960"),
    ("run đó", "lần huấn luyện đó"),
    ("log", "nhật ký"),
    ("optimizer", "bộ tối ưu"),
    ("batch size", "kích thước lô"),
    ("loss", "hàm mất mát"),
    ("attention", "chú ý"),
    ("evaluation", "quá trình đánh giá"),
]


def polish_paragraph(paragraph: Paragraph) -> None:
    for run in paragraph.runs:
        value = run.text
        for source, target in REPLACEMENTS:
            value = value.replace(source, target)
        run.text = value


def main(path_str: str) -> None:
    path = Path(path_str)
    document = Document(path)
    active = False
    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            paragraph = Paragraph(child, document)
            if paragraph.style.name == "Heading 1" and paragraph.text.startswith("CHƯƠNG 4"):
                active = True
            if active and paragraph.style.name == "Heading 1" and paragraph.text.startswith("TÀI LIỆU THAM KHẢO"):
                break
            if active:
                polish_paragraph(paragraph)
        elif active and child.tag.endswith("}tbl"):
            table = Table(child, document)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        polish_paragraph(paragraph)
    update_manual_figure_table_lists(document)
    mark_fields_dirty(document)
    document.save(path)
    print(f"Đã biên tập ngôn ngữ Chương 4: {path}")


if __name__ == "__main__":
    main(sys.argv[1])
