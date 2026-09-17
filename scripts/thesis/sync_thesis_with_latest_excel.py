#!/usr/bin/env python3
"""Synchronize thesis content with the user-confirmed latest Excel workbook.

The script preserves covers and Chapters 1–2, updates stale Chapter 3 claims,
rebuilds Chapter 4 from workbook values, and refreshes manual lists/Word fields.
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import load_workbook

from rewrite_thesis_chapter3 import ChapterBuilder, mark_fields_dirty, update_manual_figure_table_lists
from rewrite_thesis_chapter4_academic import remove_chapter4_body


def fmt(value, digits=3):
    if value is None or value == "—":
        return "—"
    if isinstance(value, str):
        return value
    return f"{value:.{digits}f}".replace(".", ",")


def signed(value, digits=3):
    return ("+" if value >= 0 else "") + fmt(value, digits)


def replace_runs(paragraph: Paragraph, replacements: list[tuple[str, str]]) -> None:
    for run in paragraph.runs:
        original = run.text
        text = original
        for source, target in replacements:
            text = text.replace(source, target)
        if text != original:
            run.text = text


def find_paragraph(document: Document, prefix: str) -> Paragraph:
    for paragraph in document.paragraphs:
        if paragraph.text.strip().startswith(prefix):
            return paragraph
    raise RuntimeError(f"Không tìm thấy đoạn bắt đầu bằng: {prefix}")


def set_paragraph(document: Document, prefix: str, text: str) -> None:
    paragraph = find_paragraph(document, prefix)
    paragraph.text = text


def remove_paragraph(document: Document, prefix: str) -> None:
    paragraph = find_paragraph(document, prefix)
    paragraph._element.getparent().remove(paragraph._element)


def next_table(paragraph: Paragraph, document: Document) -> Table:
    current = paragraph._p.getnext()
    while current is not None:
        if current.tag.endswith("}tbl"):
            return Table(current, document)
        current = current.getnext()
    raise RuntimeError(f"Không tìm thấy bảng sau caption: {paragraph.text}")


def replace_table(document: Document, caption_prefix: str, headers, rows, widths=None) -> None:
    caption = next(
        paragraph for paragraph in document.paragraphs
        if paragraph.style.name.startswith("Caption") and paragraph.text.strip().startswith(caption_prefix)
    )
    old_table = next_table(caption, document)
    builder = ChapterBuilder(document, old_table._tbl)
    builder.table(headers, rows, widths=widths)
    builder.commit()
    old_table._tbl.getparent().remove(old_table._tbl)


def replace_table_after_paragraph(document: Document, prefix: str, headers, rows, widths=None) -> None:
    paragraph = find_paragraph(document, prefix)
    old_table = next_table(paragraph, document)
    builder = ChapterBuilder(document, old_table._tbl)
    builder.table(headers, rows, widths=widths)
    builder.commit()
    old_table._tbl.getparent().remove(old_table._tbl)


def set_heading(document: Document, prefix: str, text: str) -> None:
    for paragraph in document.paragraphs:
        if paragraph.style.name.startswith("Heading") and paragraph.text.strip().startswith(prefix):
            paragraph.text = text
            return
    raise RuntimeError(f"Không tìm thấy heading bắt đầu bằng: {prefix}")


def reset_or_create_toc(document: Document) -> None:
    """Replace the cached/manual TOC area with a Word TOC field."""
    heading = find_paragraph(document, "MỤC LỤC")
    stop = find_paragraph(document, "DANH MỤC CHỮ VIẾT TẮT")
    current = heading._p.getnext()
    while current is not None and current is not stop._p:
        following = current.getnext()
        current.getparent().remove(current)
        current = following
    paragraph = document.add_paragraph(style="toc 1")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    paragraph.add_run()._r.append(begin)
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = ' TOC \\o "1-3" \\h \\z '
    paragraph.add_run()._r.append(instruction)
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    paragraph.add_run()._r.append(separate)
    paragraph.add_run("Đang cập nhật mục lục...")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(end)
    stop._p.addprevious(paragraph._p)


def update_chapter3(document: Document) -> None:
    replacements = [
        ("DAWN val", "DAWN test"),
        ("ACDC val", "ACDC test"),
        ("person, bicycle, motorcycle, car, bus và truck", "person, bicycle, car, motorcycle, bus và truck"),
        ("người, xe đạp, xe máy, ô tô, xe buýt và xe tải", "người, xe đạp, ô tô, xe máy, xe buýt và xe tải"),
        ("người đi bộ, xe đạp, xe máy, ô tô, xe buýt và xe tải", "người đi bộ, xe đạp, ô tô, xe máy, xe buýt và xe tải"),
    ]
    for paragraph in document.paragraphs:
        replace_runs(paragraph, replacements)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_runs(paragraph, replacements)

    set_paragraph(
        document,
        "Về phạm vi, đề tài tập trung vào sáu lớp",
        "Về phạm vi, đề tài tập trung vào sáu lớp đối tượng giao thông phổ biến gồm người, xe đạp, ô tô, xe máy, xe buýt "
        "và xe tải. Quy trình huấn luyện gồm Stage 1 trên BDD100K, Stage 2 trên XWOD và Phase 2 của RT-DETR-L với tập gộp "
        "XWOD + ACDC + BDD30K. Các mô hình được đánh giá trên BDD, XWOD, DAWN và ACDC theo đúng vai trò của từng giai đoạn."
    )
    set_paragraph(
        document,
        "Để khắc phục hạn chế về quy mô nhỏ",
        "Để khắc phục hạn chế về quy mô và mất cân bằng của dữ liệu thời tiết bất lợi, đề tài áp dụng học chuyển giao theo "
        "chuỗi pretrained → BDD100K → XWOD. Sau thực nghiệm so sánh, RT-DETR-L được tiếp tục ở Phase 2 bằng tập gộp "
        "XWOD + ACDC + BDD30K cùng lấy mẫu lớp hiếm. Cách tổ chức này cho phép đánh giá đồng thời khả năng thích nghi miền, "
        "chuyển miền và phục hồi tri thức BDD."
    )
    set_paragraph(
        document,
        "Xây dựng quy trình huấn luyện theo chiến lược học chuyển giao lũy tiến",
        "Xây dựng quy trình học chuyển giao gồm pretrained → BDD100K → XWOD; tiếp tục RT-DETR-L ở Phase 2 bằng XWOD, "
        "ACDC và toàn bộ BDD30K; sử dụng DAWN làm đánh giá ngoài miền và diễn giải ACDC theo đúng thời điểm trước/sau thích nghi."
    )
    set_paragraph(
        document,
        "Về dữ liệu, khóa luận sử dụng BDD100K",
        "Về dữ liệu, BDD100K được dùng cho Stage 1 và phát lại toàn bộ BDD30K ở Phase 2; XWOD là dữ liệu thời tiết bất lợi "
        "chính cho Stage 2 và Phase 2; DAWN chỉ dùng để đánh giá ngoài miền; ACDC được đánh giá chéo dữ liệu trước Phase 2 "
        "và cung cấp cả dữ liệu train lẫn tập test giữ lại sau khi được đưa vào Phase 2."
    )
    set_paragraph(
        document,
        "Tiếp theo, đề tài xây dựng quy trình chuẩn bị dữ liệu",
        "Tiếp theo, đề tài xây dựng quy trình chuẩn bị dữ liệu cho sáu lớp mục tiêu. Dữ liệu BDD100K, XWOD, DAWN và ACDC "
        "được ánh xạ về cùng thứ tự person, bicycle, car, motorcycle, bus và truck, chuyển đổi nhãn sang định dạng phù hợp "
        "và kiểm tra theo cùng quy tắc trước khi huấn luyện hoặc đánh giá."
    )
    set_paragraph(
        document,
        "Về huấn luyện, đề tài áp dụng chiến lược học chuyển giao",
        "Về huấn luyện, các mô hình bắt đầu từ trọng số pretrained, tinh chỉnh trên BDD100K ở Stage 1 và tiếp tục trên "
        "XWOD ở Stage 2. Sau khi RT-DETR-L đạt độ chính xác cao nhất trên XWOD và DAWN test, mô hình này được mở rộng ở "
        "Phase 2 bằng XWOD + ACDC + BDD30K. Kết quả được đánh giá trên XWOD, BDD, DAWN và ACDC test."
    )
    set_paragraph(
        document,
        "Quy trình huấn luyện được xây dựng theo hướng lũy tiến",
        "Quy trình huấn luyện được xây dựng theo hướng lũy tiến. Mô hình bắt đầu từ trọng số pretrained, fine-tune trên "
        "BDD100K để thích nghi miền giao thông và tiếp tục trên XWOD để học điều kiện thời tiết bất lợi. RT-DETR-L sau "
        "Stage 2 được đưa vào Phase 2 với XWOD, ACDC và toàn bộ BDD30K; DAWN tiếp tục là tập đánh giá ngoài miền."
    )
    set_paragraph(
        document,
        "Trong đề tài, ACDC được dùng ở hai vai trò",
        "Trong đề tài, ACDC được dùng ở hai vai trò. Trước Phase 2, ACDC test là phép đánh giá chéo dữ liệu cho các mô hình "
        "Stage 1 và Stage 2. Ở Phase 2, ACDC train được gộp với XWOD và BDD30K, còn ACDC test được giữ lại để đánh giá sau "
        "thích nghi; vì vậy kết quả Phase 2 trên ACDC không được gọi là zero-shot."
    )
    set_paragraph(
        document,
        "Cách phân chia này giúp quy trình nghiên cứu rõ ràng hơn",
        "Cách phân chia này giúp quy trình nghiên cứu rõ ràng hơn. Bốn mô hình được đánh giá trên BDD, XWOD, DAWN và ACDC "
        "ở Stage 1–Stage 2. Sau đó, RT-DETR-L được tiếp tục ở Phase 2 bằng dữ liệu gộp và được đánh giá lại trên cùng bốn "
        "tập, qua đó đo đồng thời khả năng duy trì XWOD, phục hồi BDD và thích nghi ACDC."
    )
    set_paragraph(
        document,
        "ACDC được sử dụng nhằm kiểm tra khả năng tổng quát hóa",
        "ACDC được sử dụng để đo chênh lệch miền trước Phase 2 và hiệu quả thích nghi sau Phase 2. Việc tách ACDC train khỏi "
        "ACDC test cho phép đánh giá RT-DETR-L trên phần dữ liệu giữ lại sau khi mô hình đã học từ cùng nguồn, đồng thời "
        "tránh diễn giải sai kết quả này như một phép đánh giá chưa thích nghi."
    )
    replace_table(
        document,
        "Bảng 2.2.",
        ["Bộ dữ liệu", "Vai trò", "Giai đoạn sử dụng"],
        [
            ["BDD100K", "Thích nghi miền giao thông; phát lại để hạn chế quên", "Stage 1; đánh giá Stage 2; train/eval Phase 2"],
            ["XWOD", "Dữ liệu thời tiết bất lợi chính", "Train/test Stage 2; train/test Phase 2"],
            ["DAWN", "Đánh giá ngoài miền, không tham gia huấn luyện", "Đánh giá Stage 1, Stage 2 và Phase 2"],
            ["ACDC", "Đánh giá chéo dữ liệu; thích nghi ở Phase 2", "Test trước Phase 2; train/test Phase 2"],
        ],
        widths=[1.15, 2.75, 2.5],
    )

    replace_table(
        document,
        "Bảng 3.2.",
        ["Bộ dữ liệu", "Quy mô theo workbook", "Vai trò", "Tập đánh giá"],
        [
            ["BDD100K", "BDD30K train", "Stage 1; phát lại toàn bộ ở Phase 2", "BDD test"],
            ["XWOD", "Khoảng 6.000 ảnh train", "Stage 2; thành phần Phase 2", "XWOD val/test"],
            ["DAWN", "Theo split test", "Đánh giá ngoài miền; không tham gia huấn luyện", "DAWN test"],
            ["ACDC", "Khoảng 1.200 ảnh train", "Thành phần Phase 2", "ACDC test"],
        ],
        widths=[1.0, 1.55, 2.4, 1.25],
    )

    set_paragraph(
        document,
        "Phase 2 được xây dựng sau khi đã chọn YOLOv8n",
        "Phase 2 được thực hiện với RT-DETR-L, là mô hình đạt độ chính xác cao nhất trong thực nghiệm so sánh Stage 2. "
        "Theo workbook kết quả mới nhất, tập huấn luyện gộp gồm khoảng 6.000 ảnh XWOD, khoảng 1.200 ảnh ACDC và toàn bộ "
        "BDD30K, tổng cộng khoảng 37.200 ảnh. BDD100K trong giai đoạn này được dùng làm dữ liệu phát lại để phục hồi và "
        "duy trì năng lực trên miền giao thông thông thường.",
    )
    set_paragraph(
        document,
        "Số 2.000 là số ảnh BDD replay",
        "Các quy mô 6k, 1,2k và 30k trong workbook là số liệu tổng hợp đã được làm tròn. Báo cáo vì vậy sử dụng ký hiệu "
        "“khoảng” khi mô tả từng thành phần và giữ tổng quy mô Phase 2 ở mức khoảng 37,2 nghìn ảnh, tránh tạo ra độ chính "
        "xác giả vượt quá dữ liệu nguồn.",
    )
    set_paragraph(
        document,
        "Phiên bản v3 tăng tần suất xuất hiện",
        "Chiến lược lấy mẫu lớp hiếm tăng tần suất ảnh chứa bicycle lên 2 lần, motorcycle lên 3 lần và bus lên 3 lần. "
        "Đây là lấy mẫu ở mức ảnh; ảnh và nhãn gốc không bị biến đổi. Workbook không cung cấp tổng số ảnh sau lấy mẫu, "
        "do đó báo cáo không nêu một quy mô tổng chưa được xác nhận.",
    )
    replace_table(
        document,
        "Bảng 3.4.",
        ["Thành phần", "Quy mô theo workbook", "Mục đích sử dụng"],
        [
            ["XWOD train", "Khoảng 6.000 ảnh", "Dữ liệu thời tiết bất lợi chính."],
            ["ACDC train", "Khoảng 1.200 ảnh", "Bổ sung cảnh đô thị trong sương mù, mưa, tuyết và ban đêm."],
            ["BDD100K train", "Toàn bộ BDD30K", "Phát lại miền giao thông để hạn chế quên thảm họa."],
            ["Tổng tập Phase 2", "Khoảng 37.200 ảnh", "Tập gộp dùng cho RT-DETR-L Phase 2."],
            ["Lấy mẫu lớp hiếm", "bicycle×2; motorcycle×3; bus×3", "Tăng cơ hội cập nhật cho các lớp ít xuất hiện."],
        ],
        widths=[1.55, 1.75, 3.1],
    )
    replace_table(
        document,
        "Bảng 3.5.",
        ["Mô hình", "Hướng tiếp cận", "Vai trò trong so sánh", "Đánh đổi chính"],
        [
            ["YOLOv8n", "Một giai đoạn", "Mốc tốc độ cao", "Nhanh và gọn; mAP thấp hơn RT-DETR-L."],
            ["YOLO11n", "Một giai đoạn", "Đối chiếu thế hệ YOLO mới", "mAP XWOD nhỉnh hơn YOLOv8n nhưng FPS thấp hơn."],
            ["Faster R-CNN", "Hai giai đoạn", "Đối chứng proposal-based", "Recall cao ở một số tập nhưng precision/mAP thấp hơn."],
            ["RT-DETR-L", "Transformer detector", "Mô hình được tiếp tục ở Phase 2", "Độ chính xác cao nhất; tốc độ thấp hơn YOLO nano."],
        ],
        widths=[1.05, 1.35, 2.0, 2.4],
    )
    set_paragraph(
        document,
        "Tiêu chí lựa chọn được tách thành hai khái niệm",
        "Kết quả Stage 2 cho thấy RT-DETR-L đạt mAP50-95 cao nhất trên XWOD test và DAWN test. Theo workbook mới nhất, "
        "đây cũng là mô hình duy nhất được tiếp tục đánh giá ở Phase 2; các mô hình còn lại chưa có kết quả Phase 2 trong "
        "nguồn dữ liệu đã xác nhận. Vì vậy, RT-DETR-L được xác định là mô hình cuối trong phạm vi bằng chứng hiện có.",
    )
    set_paragraph(
        document,
        "Sau benchmark, YOLOv8n được đưa vào Phase 2",
        "Sau thực nghiệm so sánh, RT-DETR-L được đưa vào Phase 2 từ bộ trọng số Stage 2. Giai đoạn này sử dụng tập gộp "
        "XWOD + ACDC + BDD30K và chiến lược lấy mẫu lớp hiếm. Kích thước ảnh được workbook xác nhận là 640; các cấu hình "
        "không có trong nguồn tổng hợp không được đưa vào so sánh Phase 2.",
    )
    set_paragraph(
        document,
        "Trong từng run, checkpoint tốt nhất",
        "Trong mỗi lần huấn luyện, bộ trọng số tốt nhất được chọn theo tập kiểm định thay vì mặc nhiên dùng epoch cuối. "
        "Các trường epoch hoàn thành, thời gian huấn luyện và mAP kiểm định của RT-DETR-L Phase 2 đang để trống trong "
        "workbook; vì vậy báo cáo chỉ nêu cấu hình đã xác nhận gồm tối đa 50 epoch, batch 16, lr0 = 5×10⁻⁵, patience 20, "
        "kích thước ảnh 640 và lấy mẫu lớp hiếm.",
    )
    set_heading(document, "3.5.2. Cải tiến dữ liệu", "3.5.2. Mở rộng dữ liệu Phase 2 và lấy mẫu lớp hiếm")
    set_paragraph(
        document,
        "Sau ablation kiến trúc",
        "Hướng cải tiến dữ liệu tập trung vào hai vấn đề: chênh lệch miền và mất cân bằng lớp. Phase 2 kết hợp XWOD, "
        "ACDC và BDD30K để vừa học thêm các điều kiện bất lợi, vừa duy trì miền giao thông thông thường. Trên tập gộp này, "
        "bicycle, motorcycle và bus được lấy mẫu tăng cường theo các hệ số đã ghi trong workbook.",
    )
    for prefix in [
        "Run v3_960 dùng cùng chiến lược",
        "YOLOv8s Phase 2 v3",
        "Copy-paste được thử nghiệm",
    ]:
        remove_paragraph(document, prefix)
    set_paragraph(
        document,
        "Bộ trọng số cuối cùng được kiểm tra trực tiếp ở cấp kiến trúc",
        "Khảo sát SE thuộc nhánh YOLOv8n riêng và không phải cấu hình RT-DETR-L Phase 2 cuối cùng. Workbook mới nhất không "
        "chứa số liệu định lượng của nhánh SE, vì vậy báo cáo chỉ giữ mô tả thiết kế và không dùng kết quả cũ để lựa chọn "
        "mô hình. CBAM chưa được triển khai và chỉ được xem là hướng nghiên cứu tiếp theo.",
    )
    set_paragraph(
        document,
        "Các run được tổ chức thành những nhóm",
        "Các nhóm thực nghiệm trả lời những câu hỏi riêng: Stage 1 đo khả năng thích nghi miền giao thông; Stage 2 đo "
        "khả năng học miền XWOD và chuyển sang DAWN/ACDC; khảo sát SE kiểm tra thay đổi kiến trúc; Phase 2 kiểm tra tác "
        "động của dữ liệu gộp, phát lại BDD30K và lấy mẫu lớp hiếm đối với RT-DETR-L. Ma trận kết quả chính thức chỉ "
        "bao gồm các nhóm có số liệu trong nguồn tổng hợp cập nhật.",
    )
    replace_table(
        document,
        "Bảng 3.6.",
        ["Nhóm", "Cấu hình", "Biến khảo sát", "Tập đánh giá", "Câu hỏi cần trả lời"],
        [
            ["Stage 1", "Bốn kiến trúc từ pretrained → BDD30K", "Họ mô hình", "BDD/DAWN/XWOD/ACDC test", "Khả năng thích nghi miền giao thông và chuyển miền ban đầu."],
            ["Stage 2", "Stage 1 → XWOD", "Họ mô hình", "XWOD/BDD/DAWN/ACDC test", "Độ chính xác miền đích và mức quên BDD."],
            ["SE", "YOLOv8n chuẩn và YOLOv8n + SE", "Cơ chế chú ý", "Nguồn ablation riêng", "SE có tạo cải thiện ổn định hay không."],
            ["Phase 2", "RT-DETR-L + XWOD + ACDC + BDD30K", "Thành phần dữ liệu", "Bốn tập test", "Dữ liệu gộp có phục hồi BDD và cải thiện ACDC hay không."],
            ["Lớp hiếm", "bicycle×2; motorcycle×3; bus×3", "Tần suất lấy mẫu", "Theo lớp và theo thời tiết", "Những lớp/điều kiện nào còn là hạn chế."],
        ],
        widths=[0.85, 1.65, 1.25, 1.35, 2.0],
    )
    set_paragraph(
        document,
        "Đối với cấu hình cuối, bộ trọng số YOLOv8n",
        "Theo nguồn mới nhất, cấu hình cuối được báo cáo là RT-DETR-L Phase 2 ở kích thước ảnh 640. Workbook cung cấp "
        "kết quả tổng thể, theo lớp và theo điều kiện thời tiết trên XWOD, BDD, DAWN và ACDC test. Tuy nhiên, các trường "
        "epoch hoàn thành, thời gian huấn luyện và mAP kiểm định tốt nhất của Phase 2 chưa được điền; đây là giới hạn truy "
        "vết còn lại và không được bổ sung bằng suy đoán.",
    )
    set_paragraph(
        document,
        "Từ thiết kế trên, Chương 4",
        "Từ thiết kế trên, Chương 4 trình bày môi trường triển khai, kết quả Stage 1–Stage 2, Phase 2 của RT-DETR-L, "
        "phân tích theo lớp, theo thời tiết và hiện tượng quên–phục hồi trên BDD. Các kết luận chỉ sử dụng số liệu có trong "
        "workbook mới nhất; những kết quả cũ không còn nguồn xác nhận được loại khỏi phần định lượng chính thức.",
    )


def add_chapter4(builder: ChapterBuilder, workbook_path: Path) -> None:
    wb = load_workbook(workbook_path, data_only=True)
    overview = wb["Overview — All Models"]
    stage1 = wb["Stage 1 — BDD"]
    stage2 = wb["Stage 2 — XWOD + ACDC"]
    phase2 = wb["Phase 2 — Merged"]
    training = wb["Training Details"]
    weather_dawn = wb["Weather — DAWN"]
    weather_xwod = wb["Weather — XWOD"]
    weather_acdc = wb["Weather — ACDC"]

    p, h2, h3, cap, table = builder.paragraph, builder.heading2, builder.heading3, builder.caption, builder.table

    p(
        "Chương này sử dụng bảng tổng hợp adverse_weather_model_results_google_sheets.xlsx làm nguồn số liệu thực nghiệm cập nhật. "
        "Các kết quả Stage 1, Stage 2 và Phase 2 được trình bày theo đúng mô hình, tập dữ liệu và phân hoạch ghi trong "
        "workbook. Các cấu hình không xuất hiện trong nguồn tổng hợp hiện hành không được đưa vào bảng kết quả chính thức."
    )
    p(
        "Pipeline thực nghiệm gồm pretrained → Stage 1 trên BDD100K → Stage 2 trên XWOD → Phase 2 bằng XWOD + ACDC + "
        "BDD30K. DAWN không tham gia huấn luyện và được dùng để đánh giá ngoài miền. Sau khi ACDC train được đưa vào Phase "
        "2, kết quả trên ACDC test phản ánh đánh giá giữ lại sau thích nghi, không còn là phép đo chưa thích nghi."
    )

    h2("4.1. Môi trường triển khai và cấu hình thực nghiệm")
    h3("4.1.1. Phần cứng và phần mềm")
    p(
        f"Theo sheet Training Details, môi trường được ghi nhận sử dụng {training['B3'].value}, CUDA {training['B4'].value}, "
        f"PyTorch {training['B5'].value}, Ultralytics {training['B6'].value} trên nền tảng {training['B9'].value}. "
        f"Các phép đánh giá dùng kích thước ảnh {training['B7'].value} và batch {training['B8'].value}. Vì các giá trị "
        "tốc độ phụ thuộc trực tiếp vào phần cứng và giao thức đo, báo cáo chỉ so sánh chúng trong điều kiện này."
    )
    cap("Bảng 4.1. Môi trường thực nghiệm được xác nhận trong workbook")
    table(
        ["Thành phần", "Giá trị"],
        [[training[f"A{row}"].value, str(training[f"B{row}"].value)] for row in range(3, 10)],
        widths=[2.2, 4.0],
    )

    h3("4.1.2. Cấu hình huấn luyện theo từng giai đoạn")
    p(
        "Stage 1 sử dụng BDD100K 30k; Stage 2 tiếp tục từ bộ trọng số Stage 1 và tinh chỉnh trên XWOD 6k. Phase 2 chỉ "
        "được workbook ghi nhận cho RT-DETR-L, sử dụng tập gộp khoảng 37,2 nghìn ảnh gồm XWOD, ACDC và toàn bộ BDD30K, "
        "kèm lấy mẫu tăng cường bicycle×2, motorcycle×3 và bus×3. Các trường epoch hoàn thành, dừng sớm, thời gian và mAP "
        "kiểm định tốt nhất của Phase 2 đang để trống nên không được suy diễn."
    )
    cap("Bảng 4.2. Tóm tắt cấu hình và trạng thái huấn luyện")
    training_rows = []
    for row in range(13, 22):
        values = [training.cell(row, col).value for col in range(1, 13)]
        training_rows.append([
            str(values[0]), str(values[1]), str(values[3]), str(values[4]),
            "—" if values[5] is None else str(values[5]),
            "—" if values[6] is None else str(values[6]), str(values[10]),
        ])
    table(
        ["Mô hình", "Giai đoạn", "Dữ liệu", "Epoch dự kiến", "Epoch hoàn thành", "Dừng sớm", "Cấu hình chính"],
        training_rows,
        widths=[1.15, 0.9, 1.35, 0.75, 0.9, 0.8, 1.8],
    )

    h2("4.2. Kết quả thực nghiệm so sánh bốn kiến trúc")
    h3("4.2.1. Stage 1 trên BDD100K")
    p(
        "Trên BDD test, RT-DETR-L đạt mAP50-95 cao nhất là 0,362, tiếp theo là YOLOv8n 0,297, YOLO11n 0,288 và "
        "Faster R-CNN 0,242. Faster R-CNN có recall cao nhất, đạt 0,731, nhưng precision chỉ 0,406. Điều này cho thấy mô "
        "hình phát hiện được nhiều đối tượng hơn nhưng đồng thời tạo nhiều dự đoán sai hơn các kiến trúc còn lại."
    )
    cap("Bảng 4.3. Kết quả Stage 1 trên BDD test")
    s1_rows = []
    for row in range(3, 7):
        s1_rows.append([overview[f"A{row}"].value] + [fmt(overview.cell(row, col).value) for col in range(5, 9)] + [fmt(overview[f"I{row}"].value, 1), fmt(overview[f"J{row}"].value, 1)])
    table(["Mô hình", "Precision", "Recall", "mAP50", "mAP50-95", "ms/ảnh", "FPS"], s1_rows, widths=[1.25, 0.9, 0.85, 0.8, 0.95, 0.85, 0.7])
    p(
        "Khi các bộ trọng số Stage 1 được đánh giá trực tiếp trên dữ liệu bất lợi chưa dùng để tinh chỉnh, RT-DETR-L "
        "tiếp tục dẫn đầu trên XWOD test và DAWN test. Riêng trên ACDC test, YOLOv8n đạt mAP50-95 0,181, nhỉnh hơn "
        "YOLO11n 0,179 và RT-DETR-L 0,174. Kết quả này cho thấy thứ hạng mô hình có thể thay đổi theo miền dữ liệu."
    )

    h3("4.2.2. Stage 2 trên XWOD và đánh giá chuyển miền")
    p(
        "Sau khi tinh chỉnh trên XWOD, cả bốn mô hình đều tăng mạnh trên XWOD test. RT-DETR-L đạt mAP50-95 0,500 và "
        "đứng đầu; YOLO11n đạt 0,454, cao hơn nhẹ YOLOv8n 0,452; Faster R-CNN đạt 0,381. Vì vậy, nhận định cũ rằng "
        "YOLOv8n đứng thứ hai về mAP50-95 không còn đúng theo workbook mới nhất."
    )
    cap("Bảng 4.4. Kết quả Stage 2 trên XWOD test")
    s2_rows = []
    for row in range(19, 23):
        s2_rows.append([overview[f"A{row}"].value] + [fmt(overview.cell(row, col).value) for col in range(5, 9)] + [fmt(overview[f"I{row}"].value, 1), fmt(overview[f"J{row}"].value, 1)])
    table(["Mô hình", "Precision", "Recall", "mAP50", "mAP50-95", "ms/ảnh", "FPS"], s2_rows, widths=[1.25, 0.9, 0.85, 0.8, 0.95, 0.85, 0.7])
    p(
        "Trên DAWN test, RT-DETR-L đạt mAP50-95 0,526 và precision 0,850, đều cao nhất trong nhóm; Faster R-CNN có "
        "recall cao nhất là 0,864. Trên ACDC test trước Phase 2, YOLO11n đứng đầu về mAP50-95 với 0,161, tiếp theo là "
        "YOLOv8n 0,155, RT-DETR-L 0,128 và Faster R-CNN 0,107. ACDC là miền mà quá trình tinh chỉnh chỉ bằng XWOD chưa "
        "tạo ra cải thiện ổn định."
    )
    cap("Bảng 4.5. Kết quả Stage 2 trên DAWN test và ACDC test")
    cross_rows = []
    for row in list(range(27, 31)) + list(range(31, 35)):
        cross_rows.append([overview[f"D{row}"].value, overview[f"A{row}"].value] + [fmt(overview.cell(row, col).value) for col in range(5, 9)])
    table(["Tập", "Mô hình", "Precision", "Recall", "mAP50", "mAP50-95"], cross_rows, widths=[1.1, 1.2, 0.9, 0.85, 0.85, 1.0])

    h3("4.2.3. Lựa chọn mô hình cho Phase 2")
    p(
        "RT-DETR-L được lựa chọn cho Phase 2 vì đạt mAP50-95 cao nhất trên XWOD test và DAWN test sau Stage 2. Workbook "
        "chỉ ghi nhận kết quả Phase 2 của RT-DETR-L; do đó chưa thể so sánh Phase 2 giữa các kiến trúc khi chưa có dữ liệu "
        "tương ứng. Về tốc độ, YOLOv8n vẫn nhanh hơn, nhưng kết quả "
        "này chỉ phản ánh đánh đổi ở kích thước 640 trên RTX 5090, không thay đổi mô hình đã được chọn trong nguồn mới nhất."
    )

    h2("4.3. Kết quả khảo sát cơ chế chú ý SE")
    h3("4.3.1. Phạm vi bằng chứng")
    p(
        "SE đã được triển khai như một khảo sát kiến trúc trên YOLOv8n. Tuy nhiên, workbook mới nhất không chứa bảng kết "
        "quả hoặc cấu hình của thí nghiệm này. Vì vậy, các giá trị SE cũ không được lặp lại trong phiên bản đồng bộ này và "
        "không được dùng để lựa chọn mô hình Phase 2."
    )
    h3("4.3.2. Kết luận sử dụng trong báo cáo")
    p(
        "Trong phạm vi bằng chứng còn truy vết được, SE không được đưa vào cấu hình RT-DETR-L Phase 2. Kết quả này được "
        "giữ như một hướng khảo sát riêng nhưng không được gán số liệu mới khi workbook không cung cấp. CBAM chưa được "
        "triển khai và chỉ được xem là hướng nghiên cứu tiếp theo."
    )

    h2("4.4. Phase 2 và kết quả của mô hình cuối cùng")
    h3("4.4.1. Ảnh hưởng của dữ liệu gộp và phát lại BDD")
    p(
        "Phase 2 khởi tạo từ bộ trọng số RT-DETR-L Stage 2 và sử dụng XWOD + ACDC + toàn bộ BDD30K. Trên BDD test, "
        "mAP50-95 phục hồi từ 0,183 ở Stage 2 lên 0,354, chỉ còn thấp hơn 0,008 so với Stage 1. Điều này cho thấy phát lại "
        "BDD30K giúp giảm đáng kể hiện tượng quên thảm họa sau khi mô hình được tinh chỉnh trên XWOD."
    )
    p(
        "Trên ACDC test, mAP50-95 tăng từ 0,128 lên 0,247 sau khi ACDC train được đưa vào tập gộp. Mức tăng 0,119 phản "
        "ánh hiệu quả thích nghi trên tập test giữ lại của ACDC, không phải cải thiện khi mô hình chưa từng học từ ACDC. "
        "XWOD tăng nhẹ 0,010, trong khi DAWN giảm 0,011 so với Stage 2; vì vậy Phase 2 không cải thiện đồng đều trên mọi miền."
    )
    cap("Bảng 4.6. Kết quả tổng thể của RT-DETR-L Phase 2")
    p2_rows = []
    for row in range(35, 39):
        p2_rows.append([overview[f"D{row}"].value] + [fmt(overview.cell(row, col).value) for col in range(5, 9)] + [fmt(overview[f"I{row}"].value, 1), fmt(overview[f"J{row}"].value, 1)])
    table(["Tập", "Precision", "Recall", "mAP50", "mAP50-95", "ms/ảnh", "FPS"], p2_rows, widths=[1.2, 0.9, 0.85, 0.8, 0.95, 0.85, 0.7])

    h3("4.4.2. So sánh RT-DETR-L qua các giai đoạn")
    cap("Bảng 4.7. Biến động mAP50-95 của RT-DETR-L qua các giai đoạn")
    evolution = [
        ["XWOD test", fmt(overview["H11"].value), fmt(overview["H19"].value), fmt(overview["H35"].value), signed(overview["H35"].value - overview["H19"].value)],
        ["BDD test", fmt(overview["H3"].value), fmt(overview["H23"].value), fmt(overview["H36"].value), signed(overview["H36"].value - overview["H23"].value)],
        ["DAWN test", fmt(overview["H7"].value), fmt(overview["H27"].value), fmt(overview["H37"].value), signed(overview["H37"].value - overview["H27"].value)],
        ["ACDC test", fmt(overview["H15"].value), fmt(overview["H31"].value), fmt(overview["H38"].value), signed(overview["H38"].value - overview["H31"].value)],
    ]
    table(["Tập", "Stage 1", "Stage 2", "Phase 2", "Phase 2 – Stage 2"], evolution, widths=[1.3, 1.0, 1.0, 1.0, 1.45])
    p(
        "Kết quả cho thấy Phase 2 tạo lợi ích lớn nhất trên BDD và ACDC, trong khi XWOD chỉ tăng nhẹ và DAWN giảm nhẹ. "
        "Do đó, kết luận phù hợp là dữ liệu gộp giúp cân bằng lại miền đã học và cải thiện ACDC, nhưng không tạo một mức "
        "tăng đồng thời trên tất cả các tập đánh giá."
    )

    h3("4.4.3. Phân tích theo lớp")
    cap("Bảng 4.8. mAP50-95 theo lớp của RT-DETR-L Phase 2")
    class_rows = []
    class_names = [phase2.cell(11, col).value for col in range(2, 8)]
    for row in range(12, 16):
        class_rows.append([phase2[f"A{row}"].value] + [fmt(phase2.cell(row, col).value) for col in range(2, 8)])
    table(["Tập"] + class_names, class_rows, widths=[1.15, 0.85, 0.85, 0.85, 1.0, 0.75, 0.8])
    p(
        "Trên ACDC test, bicycle và motorcycle có mAP50-95 lần lượt 0,116 và 0,119, thấp nhất trong sáu lớp; car đạt "
        "0,454 và truck đạt 0,372. Trên XWOD test, bus là lớp thấp nhất với 0,396, trong khi bicycle đạt 0,652. Kết quả "
        "cho thấy hạn chế theo lớp thay đổi theo miền dữ liệu; không thể kết luận một lớp luôn khó nhất trên mọi tập."
    )

    h3("4.4.4. Phân tích theo điều kiện thời tiết")
    cap("Bảng 4.9. mAP50-95 theo điều kiện thời tiết của RT-DETR-L Phase 2")
    weather_rows = [
        ["DAWN test", fmt(weather_dawn["C11"].value), fmt(weather_dawn["D11"].value), fmt(weather_dawn["E11"].value), fmt(weather_dawn["F11"].value), "—", "—"],
        ["XWOD test", fmt(weather_xwod["C11"].value), fmt(weather_xwod["D11"].value), fmt(weather_xwod["E11"].value), fmt(weather_xwod["F11"].value), fmt(weather_xwod["G11"].value), fmt(weather_xwod["H11"].value)],
        ["ACDC test", fmt(weather_acdc["C11"].value), fmt(weather_acdc["D11"].value), "—", fmt(weather_acdc["F11"].value), fmt(weather_acdc["G11"].value), "—"],
    ]
    table(["Tập", "Fog", "Rain", "Sand", "Snow", "Night", "Wildfire"], weather_rows, widths=[1.25, 0.75, 0.75, 0.75, 0.75, 0.8, 0.9])
    p(
        "Trên XWOD, wildfire đạt mAP50-95 cao nhất là 0,656, còn rain thấp nhất với 0,243. Trên ACDC, night thấp nhất "
        "với 0,121, cho thấy thiếu sáng vẫn là điều kiện khó sau Phase 2. Trên DAWN, snow đạt 0,616 và rain đạt 0,648; "
        "các giá trị này chỉ được so sánh trong từng bộ dữ liệu vì số lượng ảnh và phân bố lớp giữa các nhóm thời tiết khác nhau."
    )

    h2("4.5. Các phân tích bổ trợ")
    h3("4.5.1. Quên thảm họa và khả năng phục hồi")
    p(
        "Sau Stage 2, mAP50-95 trên BDD test giảm ở cả bốn kiến trúc: RT-DETR-L giảm 0,179; YOLOv8n giảm 0,098; "
        "YOLO11n giảm 0,090; Faster R-CNN giảm 0,115. Đây là dấu hiệu của quên thảm họa khi mô hình chuyển từ BDD sang "
        "XWOD. Với RT-DETR-L, Phase 2 đưa BDD test trở lại 0,354, gần mức 0,362 của Stage 1."
    )
    h3("4.5.2. Đánh đổi giữa độ chính xác và tốc độ")
    p(
        "Ở XWOD test Stage 2, RT-DETR-L đạt mAP50-95 cao nhất 0,500 với 33,7 FPS; YOLOv8n đạt 0,452 với 112,4 FPS; "
        "YOLO11n đạt 0,454 với 95,2 FPS; Faster R-CNN đạt 0,381 với 61,3 FPS. RT-DETR-L ưu tiên độ chính xác, trong khi "
        "các mô hình YOLO nano có lợi thế tốc độ. Các số liệu này được đo ở batch 1, kích thước 640 trên RTX 5090 và không "
        "được khái quát thành tốc độ cố định trên phần cứng khác."
    )

    h2("4.6. Thảo luận")
    h3("4.6.1. Các phát hiện chính và lựa chọn mô hình")
    p(
        "Trong thực nghiệm so sánh, RT-DETR-L đạt độ chính xác cao nhất trên BDD test ở Stage 1 và trên XWOD/DAWN test "
        "ở Stage 2. Mô hình này được tiếp tục ở Phase 2 và đạt mAP50-95 lần lượt 0,510 trên XWOD, 0,354 trên BDD, 0,515 "
        "trên DAWN và 0,247 trên ACDC. Vì workbook chưa có Phase 2 của ba kiến trúc còn lại, kết luận chỉ giới hạn ở việc "
        "RT-DETR-L là mô hình cuối được xác nhận, không khẳng định nó luôn tối ưu cho mọi yêu cầu triển khai."
    )
    p(
        "Phase 2 đặc biệt hữu ích cho khả năng phục hồi BDD và thích nghi ACDC. Tuy nhiên, DAWN giảm nhẹ so với Stage 2 "
        "và ACDC vẫn có recall 0,384; bicycle, motorcycle và điều kiện ban đêm tiếp tục là những hạn chế quan trọng. Kết "
        "quả cho thấy dữ liệu gộp cải thiện sự cân bằng đa miền nhưng chưa giải quyết triệt để hiện tượng bỏ sót."
    )
    h3("4.6.2. Hạn chế và hướng phát triển")
    p(
        "Workbook chưa ghi epoch hoàn thành, thời gian huấn luyện và mAP kiểm định tốt nhất của RT-DETR-L Phase 2. Ngoài "
        "ra, chưa có kết quả Phase 2 cho YOLOv8n, YOLO11n và Faster R-CNN trong cùng giao thức. Đây là các khoảng trống cần "
        "được bổ sung trước khi đưa ra so sánh Phase 2 toàn diện hoặc kết luận chắc chắn về đánh đổi triển khai."
    )
    p(
        "Hướng phát triển phù hợp là hoàn thiện nhật ký Phase 2, đánh giá ba kiến trúc còn lại trên cùng tập gộp, tăng dữ "
        "liệu thật cho bicycle và motorcycle, đồng thời khảo sát riêng các cảnh ban đêm. SE và CBAM chỉ nên được đưa trở "
        "lại bảng định lượng khi có kết quả được ghi trong nguồn dữ liệu mới với giao thức đánh giá xác định."
    )


def update_appendices(document: Document) -> None:
    set_paragraph(
        document,
        "Bảng dưới đây tổng hợp các siêu tham số",
        "Bảng dưới đây tổng hợp cấu hình tham chiếu được workbook xác nhận. Các kiến trúc có bộ tối ưu và lịch huấn luyện "
        "khác nhau; cấu hình RT-DETR-L Phase 2 sử dụng kích thước ảnh 640, batch 16, lr0 = 5×10⁻⁵ và patience 20.",
    )
    replace_table_after_paragraph(
        document,
        "Bảng dưới đây tổng hợp cấu hình tham chiếu",
        ["Thành phần", "Giá trị theo workbook"],
        [
            ["Mô hình Phase 2", "RT-DETR-L"],
            ["Dữ liệu", "XWOD khoảng 6k + ACDC khoảng 1,2k + BDD30K"],
            ["Số epoch tối đa", "50"],
            ["Epoch hoàn thành", "Chưa được ghi trong workbook"],
            ["Dừng sớm", "patience = 20; trạng thái chưa được ghi"],
            ["Kích thước ảnh", "640 × 640"],
            ["Batch size", "16"],
            ["Tốc độ học ban đầu", "5×10⁻⁵"],
            ["Lấy mẫu lớp hiếm", "bicycle×2; motorcycle×3; bus×3"],
            ["Phần cứng đánh giá", "NVIDIA RTX 5090, 32 GB VRAM"],
        ],
        widths=[2.25, 4.0],
    )
    set_paragraph(
        document,
        "Một số kết quả phát hiện của mô hình cuối cùng (YOLOv8n @960)",
        "Các hình trong phụ lục minh họa kết quả phát hiện trên XWOD, DAWN và ACDC. Do workbook mới nhất xác định mô hình "
        "cuối là RT-DETR-L Phase 2 nhưng không chứa siêu dữ liệu liên kết từng ảnh minh họa với checkpoint, các hình này "
        "không được dùng làm bằng chứng định lượng cho mô hình cuối.",
    )
    set_paragraph(
        document,
        "Toàn bộ mã nguồn được tổ chức",
        "Toàn bộ mã nguồn được tổ chức trong các tệp mô tả ở Phụ lục A. Phần này trích dẫn mô-đun SE dùng cho khảo sát "
        "kiến trúc trên YOLOv8n. Mô-đun này không thuộc cấu hình RT-DETR-L Phase 2 cuối cùng được xác nhận trong workbook."
    )


def main(source: str, workbook: str, output: str) -> None:
    source_path = Path(source)
    workbook_path = Path(workbook)
    output_path = Path(output)
    document = Document(source_path)
    update_chapter3(document)
    _, references = remove_chapter4_body(document)
    builder = ChapterBuilder(document, references)
    add_chapter4(builder, workbook_path)
    builder.commit()
    update_appendices(document)
    update_manual_figure_table_lists(document)
    reset_or_create_toc(document)
    mark_fields_dirty(document)
    document.save(output_path)
    print(f"Đã tạo bản đồng bộ nội dung: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit("Usage: sync_thesis_with_latest_excel.py SOURCE.docx RESULTS.xlsx OUTPUT.docx")
    main(*sys.argv[1:])
