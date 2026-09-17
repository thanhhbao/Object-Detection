#!/usr/bin/env python3
"""Rewrite thesis Chapter 4 with verified metrics and academic prose."""

from pathlib import Path
import sys

from docx import Document

from rewrite_thesis_chapter3 import (
    ChapterBuilder,
    FIGURES,
    mark_fields_dirty,
    update_manual_figure_table_lists,
)
from rewrite_thesis_chapter3_academic import reset_toc_cache


def remove_chapter4_body(document: Document):
    body = document.element.body
    chapter4 = None
    references = None
    for child in body.iterchildren():
        if not child.tag.endswith("}p"):
            continue
        text = "".join(child.xpath(".//w:t/text()" )).strip()
        style = child.find("./w:pPr/w:pStyle", namespaces=child.nsmap)
        style_id = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if style is not None else ""
        if text.startswith("CHƯƠNG 4") and style_id == "Heading1":
            chapter4 = child
        elif chapter4 is not None and text.startswith("TÀI LIỆU THAM KHẢO") and style_id == "Heading1":
            references = child
            break
    if chapter4 is None or references is None:
        raise RuntimeError("Không tìm thấy ranh giới Chương 4/Tài liệu tham khảo.")
    current = chapter4.getnext()
    while current is not None and current is not references:
        following = current.getnext()
        body.remove(current)
        current = following
    return chapter4, references


def add_chapter4(builder: ChapterBuilder) -> None:
    p = builder.paragraph
    h2 = builder.heading2
    h3 = builder.heading3
    cap = builder.caption
    table = builder.table

    p(
        "Chương 3 đã xác định kiến trúc hệ thống, vai trò của từng bộ dữ liệu và các nhóm thực nghiệm. Chương 4 trình bày "
        "quá trình triển khai những thiết kế đó, từ môi trường huấn luyện đến kết quả benchmark, khảo sát SE và các cấu hình "
        "Phase 2. Các số liệu của mô hình cuối được ưu tiên từ lần đánh giá tái lập trực tiếp trên checkpoint best.pt; những "
        "kết quả benchmark chỉ còn ở dạng bảng tổng hợp được ghi rõ giới hạn nguồn thay vì được xem như log tái tạo đầy đủ."
    )
    p(
        "Phần phân tích tuân theo hai nguyên tắc. Kết quả trên XWOD val được tách khỏi XWOD test vì hai split có chức năng "
        "khác nhau. ACDC val cũng được diễn giải theo đúng thời điểm sử dụng: trước Phase 2 là đánh giá chéo dữ liệu, còn sau "
        "khi đã dùng ACDC train thì đây là phép đánh giá trên phần dữ liệu được giữ lại. Nhờ đó, mức cải thiện không bị gán "
        "cho khả năng zero-shot khi mô hình thực tế đã được thích nghi với ACDC."
    )

    h2("4.1. Môi trường triển khai và cấu hình thực nghiệm")
    h3("4.1.1. Phần cứng và phần mềm")
    p(
        "Các thí nghiệm được thực hiện trên máy chủ GPU thuê theo phiên, chủ yếu thông qua Vast.ai. Môi trường huấn luyện "
        "sử dụng Ubuntu Linux, Python, PyTorch và CUDA tương ứng với GPU của từng phiên. Một số run chính được thực hiện trên "
        "GPU NVIDIA có bộ nhớ lớn, trong đó môi trường được ghi trong báo cáo gồm RTX 5090 32 GB. Do phần cứng có thể khác "
        "giữa các đợt, giá trị thời gian suy luận chỉ được so sánh khi cùng xuất phát từ một môi trường đo."
    )
    p(
        "Các phiên huấn luyện dài được duy trì bằng tmux; checkpoint, cấu hình và kết quả theo epoch được lưu ngoài tiến trình "
        "đang chạy. Cách tổ chức này giúp hạn chế mất kết quả khi phiên kết nối bị gián đoạn. Đối với lần tái xác minh mô hình "
        "cuối, checkpoint được nạp lại trong môi trường Ultralytics, kiến trúc được in trực tiếp và evaluation được chạy lại "
        "riêng trên XWOD test, DAWN val và ACDC val ở kích thước ảnh 960."
    )
    cap("Bảng 4.1. Các thư viện chính và vai trò trong hệ thống")
    table(
        ["Thư viện/công cụ", "Vai trò"],
        [
            ["PyTorch", "Nền tảng tính toán và huấn luyện mô hình học sâu."],
            ["Ultralytics", "Huấn luyện, đánh giá và suy luận YOLOv8n, YOLO11n và RT-DETR."],
            ["torchvision", "Cài đặt Faster R-CNN với backbone ResNet-50 và FPN."],
            ["OpenCV", "Đọc ảnh/video, tiền xử lý và trực quan hóa hộp giới hạn."],
            ["NumPy, Pandas", "Xử lý số liệu, kết hợp bảng và tổng hợp kết quả."],
            ["pycocotools", "Hỗ trợ biểu diễn và đánh giá dữ liệu theo quy ước COCO."],
            ["Albumentations", "Thực hiện các phép tăng cường dữ liệu khi cấu hình yêu cầu."],
            ["Streamlit", "Xây dựng giao diện minh họa quá trình suy luận."],
        ],
        widths=[1.8, 4.6],
    )

    h3("4.1.2. Cấu hình dữ liệu và huấn luyện")
    p(
        "Stage 1 sử dụng khoảng 30.000 ảnh BDD train và khoảng 10.000 ảnh validation. Stage 2 sử dụng 5.945 ảnh XWOD train "
        "và 1.744 ảnh XWOD val. Sau benchmark, Phase 2 kết hợp 5.945 ảnh XWOD, 1.572 ảnh ACDC và 2.000 ảnh BDD replay, tạo "
        "tập cơ sở 9.517 ảnh. Phiên bản v3 tăng mẫu ở mức ảnh cho bicycle, motorcycle và bus, đưa tập train lên 13.430 ảnh; "
        "XWOD val tiếp tục được dùng để chọn checkpoint."
    )
    p(
        "Bảng 4.2 trình bày cấu hình tham chiếu được dùng cho các run YOLO chính. Đây không phải khẳng định rằng Faster R-CNN, "
        "RT-DETR và mọi run bổ trợ dùng cùng một optimizer hoặc batch size. Các kiến trúc thuộc thư viện khác nhau có cách "
        "cài đặt loss và lịch học riêng; khi có khác biệt, cấu hình trong args.yaml hoặc log của chính run đó được ưu tiên."
    )
    cap("Bảng 4.2. Cấu hình tham chiếu của các run YOLO")
    table(
        ["Thành phần", "Thiết lập tham chiếu", "Ghi chú"],
        [
            ["Số epoch tối đa", "50", "Dừng sớm theo patience khi run áp dụng."],
            ["Kích thước ảnh", "640; cấu hình cuối 960", "Là biến thực nghiệm, không cố định cho mọi run."],
            ["Batch size", "16 ở cấu hình cơ sở", "Có thể điều chỉnh theo bộ nhớ ở imgsz 960."],
            ["Optimizer", "AdamW", "Áp dụng cho nhóm run YOLO được ghi nhận."],
            ["Learning rate", "lr0 = 0,0005; lrf = 0,01", "Lịch cosine theo cấu hình run."],
            ["Weight decay", "0,0005", "Lấy từ cấu hình tham chiếu."],
            ["Seed", "42", "Bật deterministic khi môi trường hỗ trợ."],
            ["Mixed precision", "AMP", "Giảm bộ nhớ và tăng tốc trên GPU phù hợp."],
        ],
        widths=[1.45, 1.75, 3.2],
    )
    p(
        "Đường cong huấn luyện của run v3_960 cho thấy loss huấn luyện giảm đều trong 50 epoch, trong khi mAP50-95 trên "
        "XWOD val tăng nhanh ở giai đoạn đầu rồi dao động quanh vùng hội tụ. Theo results.csv, checkpoint tốt nhất đạt "
        "mAP50-95 xấp xỉ 0,34685 ở epoch 38. Việc chọn best.pt thay vì last.pt giúp tránh sử dụng epoch cuối chỉ vì nó kết "
        "thúc quá trình huấn luyện."
    )
    builder.picture(FIGURES / "hinh_4_2b_training_curves.png", width=6.3)
    cap("Hình 4.1. Đường cong huấn luyện và validation của run YOLOv8n Phase 2 v3 ở imgsz 960")

    h2("4.2. Kết quả benchmark bốn kiến trúc")
    h3("4.2.1. Kết quả validation theo từng giai đoạn")
    p(
        "Bảng 4.3 tổng hợp các mốc validation còn được xác nhận nhất quán cho Stage 1 và Stage 2. Stage 1 cho thấy RT-DETR "
        "đạt mAP50-95 cao hơn hai biến thể YOLO nano trên BDD validation. Sang XWOD val, RT-DETR tiếp tục đứng đầu với 0,359; "
        "YOLOv8n đạt 0,325, YOLO11n đạt 0,311 và Faster R-CNN đạt 0,272. Faster R-CNN có recall 0,710 nhưng precision thấp hơn, "
        "cho thấy mô hình tạo nhiều phát hiện hơn nhưng đồng thời tăng dự đoán sai."
    )
    cap("Bảng 4.3. Kết quả validation của benchmark theo giai đoạn")
    table(
        ["Giai đoạn", "Mô hình", "Precision", "Recall", "mAP50", "mAP50-95"],
        [
            ["Stage 1 – BDD val", "YOLOv8n", "0,598", "0,453", "0,488", "0,292"],
            ["Stage 1 – BDD val", "YOLO11n", "0,605", "0,434", "0,475", "0,287"],
            ["Stage 1 – BDD val", "RT-DETR", "0,699", "0,579", "0,635", "0,380"],
            ["Stage 1 – BDD val", "Faster R-CNN", "—", "—", "—", "—"],
            ["Stage 2 – XWOD val", "YOLOv8n", "0,637", "0,500", "0,536", "0,325"],
            ["Stage 2 – XWOD val", "YOLO11n", "0,660", "0,465", "0,521", "0,311"],
            ["Stage 2 – XWOD val", "RT-DETR", "0,632", "0,569", "0,595", "0,359"],
            ["Stage 2 – XWOD val", "Faster R-CNN", "0,429", "0,710", "0,495", "0,272"],
        ],
        widths=[1.35, 1.15, 0.9, 0.8, 0.8, 0.9],
    )
    p(
        "Các số trong bảng phản ánh validation, không được thay thế trực tiếp cho XWOD test. Chúng được dùng để theo dõi "
        "khả năng thích nghi sau từng stage và lựa chọn checkpoint. Việc giữ riêng hai loại kết quả tránh nhầm giá trị 0,325 "
        "của YOLOv8n trên XWOD val với giá trị 0,266 trong bảng tổng hợp XWOD test của benchmark ban đầu."
    )

    h3("4.2.2. Kết quả trên XWOD test và DAWN")
    p(
        "Bảng 4.4 là kết quả XWOD test đã xuất hiện nhất quán trong báo cáo và bảng tổng hợp cũ. Full text log của cả bốn "
        "mô hình hiện chưa đầy đủ như provenance của mô hình final, vì vậy bảng được giữ ở mức kết quả tổng hợp, không được "
        "mô tả như một lần tái đánh giá mới. Theo nguồn này, RT-DETR có mAP50-95 cao nhất; YOLOv8n đứng thứ hai và có thời "
        "gian suy luận thấp hơn đáng kể trong môi trường đo được ghi nhận."
    )
    cap("Bảng 4.4. Kết quả tổng hợp benchmark trên XWOD test")
    table(
        ["Mô hình", "Precision", "Recall", "mAP50", "mAP50-95", "ms/ảnh"],
        [
            ["RT-DETR", "0,584", "0,490", "0,500", "0,290", "36,4"],
            ["YOLOv8n", "0,570", "0,459", "0,474", "0,266", "3,0"],
            ["YOLO11n", "0,559", "0,440", "0,451", "0,250", "3,0"],
            ["Faster R-CNN", "Không có", "0,353", "0,354", "0,176", "xấp xỉ 140"],
        ],
        widths=[1.25, 0.95, 0.85, 0.85, 0.95, 0.9],
    )
    p(
        "DAWN val là phép đánh giá ngoài miền vì không tham gia Stage 1, Stage 2 hoặc Phase 2. Kết quả ở Bảng 4.5 tiếp tục "
        "cho thấy RT-DETR có mAP50-95 cao nhất trong benchmark ban đầu. YOLOv8n đạt precision cao nhất nhưng recall thấp hơn "
        "RT-DETR; điều này phù hợp với nhận xét rằng mô hình nano thận trọng hơn và bỏ sót nhiều đối tượng hơn. Precision của "
        "Faster R-CNN không có trong log tổng hợp, do đó được giữ ở trạng thái không khả dụng."
    )
    cap("Bảng 4.5. Kết quả benchmark trên DAWN val")
    table(
        ["Mô hình", "Precision", "Recall", "mAP50", "mAP50-95"],
        [
            ["RT-DETR", "0,627", "0,739", "0,747", "0,475"],
            ["YOLOv8n", "0,729", "0,513", "0,634", "0,415"],
            ["YOLO11n", "0,546", "0,567", "0,554", "0,337"],
            ["Faster R-CNN", "Không có", "0,377", "0,444", "0,240"],
        ],
        widths=[1.45, 1.1, 1.0, 1.0, 1.1],
    )

    h3("4.2.3. Lựa chọn mô hình tiếp tục Phase 2")
    p(
        "RT-DETR là mô hình có độ chính xác tuyệt đối cao nhất trong benchmark ban đầu. Tuy nhiên, mục tiêu của đề tài còn "
        "bao gồm khả năng triển khai, nên quyết định không chỉ dựa trên mAP. YOLOv8n có khoảng ba triệu tham số, chi phí tính "
        "toán thấp và hệ thống suy luận đơn giản hơn, trong khi vẫn đứng thứ hai trên các bảng tổng hợp. Vì lý do đó, YOLOv8n "
        "được chọn làm mô hình thực dụng để tiếp tục Phase 2; RT-DETR được giữ như mốc tham chiếu về độ chính xác."
    )
    p(
        "Hình 4.2 biểu diễn sự đánh đổi dựa trên số liệu XWOD test và thời gian suy luận trong bảng tổng hợp. Hình này chỉ có "
        "ý nghĩa trong điều kiện đo đã ghi nhận; không dùng để khẳng định một tốc độ cố định trên mọi GPU. Điểm của YOLOv8n "
        "nằm gần vùng tốc độ cao trong khi giữ mAP tốt hơn YOLO11n và Faster R-CNN, qua đó minh họa trực quan cơ sở của lựa chọn."
    )
    builder.picture(FIGURES / "hinh_4_1_speed_accuracy.png")
    cap("Hình 4.2. Đánh đổi tốc độ và độ chính xác trong benchmark XWOD test")

    h2("4.3. Kết quả khảo sát cơ chế chú ý SE")
    h3("4.3.1. Kết quả định lượng")
    p(
        "Ablation SE so sánh YOLOv8n chuẩn với biến thể có một khối SE sau SPPF. Trên XWOD test, mAP50-95 tăng rất nhẹ từ "
        "0,266 lên 0,268, trong khi recall thay đổi từ 0,448 xuống 0,446. Nếu chỉ nhìn vào XWOD, chênh lệch này gần như trung "
        "tính. Trên DAWN, mAP50-95 giảm từ 0,411 xuống 0,396 và recall giảm rõ từ 0,572 xuống 0,492."
    )
    cap("Bảng 4.6. So sánh YOLOv8n chuẩn và YOLOv8n có SE")
    table(
        ["Mô hình", "Tập đánh giá", "mAP50", "mAP50-95", "Recall"],
        [
            ["YOLOv8n chuẩn", "XWOD test", "0,472", "0,266", "0,448"],
            ["YOLOv8n + SE", "XWOD test", "0,479", "0,268", "0,446"],
            ["YOLOv8n chuẩn", "DAWN val", "0,634", "0,411", "0,572"],
            ["YOLOv8n + SE", "DAWN val", "0,612", "0,396", "0,492"],
        ],
        widths=[1.55, 1.45, 1.0, 1.1, 1.0],
    )

    h3("4.3.2. Diễn giải và quyết định cấu hình")
    p(
        "Kết quả không cho phép kết luận SE hoàn toàn không có tác dụng: trên XWOD, biến thể SE vẫn tạo một mức tăng nhỏ. "
        "Vấn đề nằm ở tính nhất quán. Mức tăng cùng miền không đi kèm cải thiện ngoài miền, trong khi recall DAWN suy giảm "
        "đáng kể. Với mục tiêu phát hiện trong điều kiện thay đổi, một mô-đun chỉ hữu ích khi lợi ích không phụ thuộc quá mạnh "
        "vào nguồn dữ liệu đã dùng để tinh chỉnh."
    )
    p(
        "SE vì vậy bị loại khỏi cấu hình cuối. Kết luận này còn được kiểm chứng ở cấp checkpoint: kiến trúc của best.pt final "
        "chỉ chứa các khối chuẩn của YOLOv8n và không có SEBlock, SqueezeExcitation hoặc lớp attention tương đương. CBAM "
        "không được triển khai trong đề tài; nếu được nhắc đến, mô-đun này chỉ thuộc hướng nghiên cứu tiếp theo."
    )

    h2("4.4. Phase 2 và kết quả của mô hình cuối cùng")
    h3("4.4.1. Ảnh hưởng của dữ liệu gộp, lấy mẫu và độ phân giải")
    p(
        "Bảng 4.7 tổng hợp các cấu hình Phase 2 theo mAP50-95. Các cột được hiểu là những cấu hình thực nghiệm khác nhau, "
        "không mặc nhiên là một chuỗi checkpoint nối tiếp. v2 sử dụng tập gộp; v3 bổ sung lấy mẫu tăng cường cho lớp hiếm; "
        "v3_960 giữ chiến lược dữ liệu v3 và tăng kích thước ảnh. Quan hệ khởi tạo trực tiếp giữa v3_rare @640 và v3_960 "
        "@960 chưa có metadata đủ mạnh để khẳng định."
    )
    cap("Bảng 4.7. mAP50-95 của các cấu hình trong quá trình mở rộng")
    table(
        ["Tập đánh giá", "Cơ sở chỉ XWOD", "Phase 2 v2", "Phase 2 v3 @640", "Phase 2 v3 @960"],
        [
            ["XWOD test", "0,266", "0,278", "0,274", "0,279"],
            ["DAWN val", "0,411", "0,435", "0,450", "0,541"],
            ["ACDC val", "0,150", "0,188", "0,190", "0,239"],
        ],
        widths=[1.45, 1.2, 1.1, 1.35, 1.35],
    )
    p(
        "Việc gộp dữ liệu tạo mức tăng trên DAWN và ACDC trong bảng tổng hợp, trong khi tác động trên XWOD nhỏ hơn. v3 @640 "
        "không vượt v2 trên XWOD nhưng cải thiện thêm DAWN và ACDC, cho thấy lấy mẫu lớp hiếm có thể tạo đánh đổi giữa các "
        "miền. Khi tăng imgsz lên 960, cả ba tập đều đạt mốc cao nhất trong nhóm YOLOv8n. Riêng ACDC cần được diễn giải thận "
        "trọng: giá trị 0,150 trước Phase 2 là đánh giá chéo dữ liệu, còn 0,239 sau Phase 2 là held-out sau khi mô hình đã học "
        "từ ACDC train. Chênh lệch 0,089 mô tả hiệu quả thích nghi, không phải cải thiện zero-shot thuần túy."
    )
    builder.picture(FIGURES / "hinh_4_2_evolution.png")
    cap("Hình 4.3. mAP50-95 của các cấu hình trong quá trình mở rộng mô hình")

    p(
        "Để trả lời trực tiếp câu hỏi mô hình thay đổi như thế nào khi gặp thời tiết bất lợi, Bảng 4.8 đối chiếu mô hình "
        "cơ sở và mô hình cuối trên cùng ba tập đánh giá. Trong phép đối chiếu này, mô hình cơ sở là YOLOv8n sau Stage 2, "
        "được tinh chỉnh bằng XWOD nhưng chưa trải qua Phase 2; mô hình cuối là YOLOv8n Phase 2 v3 @960. Đây là so sánh "
        "giữa hai đầu của toàn bộ quá trình phát triển, nên mức chênh lệch phản ánh tác động kết hợp của dữ liệu Phase 2, "
        "lấy mẫu lớp hiếm và độ phân giải 960, không được dùng để quy toàn bộ cải thiện cho riêng một thành phần."
    )
    cap("Bảng 4.8. So sánh mô hình cơ sở và mô hình cuối trong điều kiện thời tiết bất lợi")
    table(
        ["Tập đánh giá", "mAP50-95 cơ sở", "mAP50-95 cuối", "Chênh lệch", "Recall cơ sở", "Recall cuối"],
        [
            ["XWOD test", "0,266", "0,279", "+0,013", "0,448", "0,472"],
            ["DAWN val", "0,411", "0,541", "+0,130", "0,572", "0,731"],
            ["ACDC val", "0,150", "0,239", "+0,089", "Không có nhật ký", "0,390"],
        ],
        widths=[1.2, 1.05, 1.0, 0.9, 1.15, 0.9],
    )
    p(
        "Trên XWOD, mAP50-95 chỉ tăng 0,013 và recall tăng từ 0,448 lên 0,472. Điều này cho thấy mô hình cơ sở đã nhận biết "
        "được phần lớn các mẫu thời tiết xấu cùng miền sau Stage 2; Phase 2 chủ yếu duy trì năng lực đó và tạo thêm một mức "
        "cải thiện vừa phải, thay vì đánh đổi XWOD để lấy kết quả trên nguồn dữ liệu khác. Trên DAWN, recall tăng 0,159 và "
        "mAP50-95 tăng 0,130. Vì DAWN hoàn toàn không tham gia huấn luyện, đây là bằng chứng rõ nhất rằng mô hình cuối bỏ sót "
        "ít đối tượng hơn và tạo hộp dự đoán đúng, ổn định hơn khi chuyển sang một nguồn ảnh thời tiết xấu chưa quan sát."
    )
    p(
        "Trên ACDC, mAP50-95 tăng từ 0,150 lên 0,239, tương đương 0,089 điểm tuyệt đối. Mô hình cơ sở gặp khó khăn rõ hơn "
        "khi biên vật thể suy giảm do sương mù, mưa, tuyết hoặc thiếu sáng; mô hình cuối nhận biết và định vị tốt hơn sau khi "
        "được thích nghi bằng ACDC train. Tuy nhiên, recall cuối chỉ đạt 0,390 nên hiện tượng bỏ sót vẫn còn đáng kể, nhất là "
        "với vật thể nhỏ, xa, bị che khuất hoặc có độ tương phản thấp. Kết quả ACDC phản ánh hiệu quả thích nghi trên phần dữ "
        "liệu giữ lại, không phải khả năng suy luận khi chưa từng học từ ACDC."
    )
    p(
        "Từ góc độ hành vi phát hiện, mô hình cơ sở vẫn phát hiện được các phương tiện lớn và có đường biên rõ, nhưng kém ổn "
        "định hơn khi chi tiết bị thời tiết che lấp. Mô hình cuối giảm bỏ sót rõ nhất trên DAWN, giữ kết quả XWOD và cải thiện "
        "độ chính xác tổng hợp trên ACDC. Diễn giải này dựa trên các chỉ số đánh giá; repo hiện chưa lưu bộ ảnh dự đoán ghép "
        "cặp của hai checkpoint trên cùng khung hình, do đó báo cáo không trình bày nhận xét trực quan như một bằng chứng đã "
        "được quan sát trực tiếp."
    )

    h3("4.4.2. Kết quả tái đánh giá trực tiếp checkpoint final")
    p(
        "Mô hình thực dụng cuối cùng là YOLOv8n Phase 2 v3 với lấy mẫu lớp hiếm và imgsz 960. Checkpoint best.pt đã được "
        "tái đánh giá trực tiếp trên đúng ba tập dữ liệu. Đây là nguồn có provenance mạnh nhất của mô hình final, thay cho "
        "việc chỉ giữ số liệu trong báo cáo. Bảng 4.9 trình bày số ảnh hợp lệ, số đối tượng và bốn độ đo tổng thể."
    )
    cap("Bảng 4.9. Kết quả tái đánh giá mô hình YOLOv8n Phase 2 v3 @960")
    table(
        ["Tập đánh giá", "Ảnh", "Đối tượng", "Precision", "Recall", "mAP50", "mAP50-95"],
        [
            ["XWOD test", "2.320", "9.801", "0,649", "0,472", "0,504", "0,279"],
            ["DAWN val", "206", "1.806", "0,801", "0,731", "0,788", "0,541"],
            ["ACDC val", "398", "2.830", "0,541", "0,390", "0,399", "0,239"],
        ],
        widths=[1.25, 0.7, 0.85, 0.85, 0.75, 0.75, 0.9],
    )
    p(
        "XWOD test có 2.321 tệp nguồn nhưng chỉ 2.320 ảnh hợp lệ được đánh giá vì flooding_test_00217.jpg thực chất là tệp "
        "GIF89a. DAWN đạt kết quả cao nhất trong ba tập, với mAP50-95 bằng 0,541 và recall bằng 0,731. ACDC vẫn là tập khó "
        "nhất: recall chỉ 0,390 và mAP50-95 đạt 0,239. Kết quả này chỉ ra rằng chiến lược Phase 2 cải thiện đáng kể khả năng "
        "thích nghi nhưng chưa giải quyết triệt để hiện tượng bỏ sót trong các cảnh suy giảm mạnh."
    )

    h3("4.4.3. Phân tích theo lớp và dạng lỗi")
    cap("Bảng 4.10. mAP50-95 theo lớp của mô hình final")
    table(
        ["Lớp", "XWOD test", "DAWN val", "ACDC val"],
        [
            ["person", "0,273", "0,531", "0,205"],
            ["bicycle", "0,226", "0,733*", "0,071"],
            ["car", "0,499", "0,604", "0,533"],
            ["motorcycle", "0,136", "0,522", "0,174"],
            ["bus", "0,270", "0,406", "0,175"],
            ["truck", "0,272", "0,451", "0,275"],
        ],
        widths=[1.45, 1.55, 1.55, 1.55],
    )
    p(
        "Car là lớp ổn định nhất, đạt 0,499 trên XWOD và 0,533 trên ACDC. Bicycle và motorcycle là hai lớp yếu nhất ở "
        "ACDC, tương ứng 0,071 và 0,174; trên XWOD, motorcycle cũng thấp nhất với 0,136. Kết quả phù hợp với đặc điểm của "
        "hai lớp này: số mẫu ít hơn, kích thước thường nhỏ và biên dễ mất khi ảnh mờ hoặc thiếu sáng. Person, bus và truck "
        "nằm ở nhóm trung gian nhưng đều giảm đáng kể trên ACDC so với DAWN."
    )
    p(
        "Giá trị 0,733 của bicycle trên DAWN không nên được xem là bằng chứng rằng lớp này đã được giải quyết tốt. DAWN val "
        "chỉ có sáu đối tượng bicycle, nên AP rất nhạy với từng dự đoán. Khi đối chiếu với 0,226 trên XWOD và 0,071 trên ACDC, "
        "kết luận hợp lý hơn là hiệu năng bicycle chưa ổn định giữa các nguồn dữ liệu."
    )
    builder.picture(FIGURES / "hinh_4_3_confusion_xwod.png", width=6.1)
    cap("Hình 4.4. Ma trận nhầm lẫn chuẩn hóa của mô hình final trên XWOD test")
    p(
        "Ma trận nhầm lẫn cho thấy car có đường chéo nổi bật hơn các lớp còn lại. Bicycle và motorcycle có tỉ lệ dự đoán đúng "
        "thấp hơn, đồng thời xuất hiện nhầm lẫn giữa hai lớp phương tiện hai bánh. Bus và truck cũng có xu hướng nhầm lẫn qua "
        "lại hoặc bị gán thành car khi hình dạng bị che khuất. Cột và hàng background cho thấy bỏ sót và phát hiện nền vẫn là "
        "hai nguồn lỗi quan trọng, phù hợp với recall tổng thể còn thấp trên XWOD và ACDC."
    )

    h3("4.4.4. Quan sát định tính theo điều kiện thời tiết")
    p(
        "Quan sát các ảnh lỗi cho thấy sương mù làm giảm biên của vật thể ở xa; mưa và mặt đường ướt tạo phản xạ; tuyết làm "
        "nền sáng và che khuất kết cấu; bão cát làm giảm tương phản trên phần lớn khung hình. Các tác động thường mạnh hơn khi "
        "kết hợp với thiếu sáng, vật thể nhỏ hoặc che khuất. Đây là phân tích định tính từ ví dụ và ma trận lỗi, không phải "
        "bảng xếp hạng độ khó theo thời tiết vì log mAP đầy đủ cho từng điều kiện chưa được tái xác minh."
    )
    p(
        "Hình 4.5 minh họa kết quả trên ACDC. Mô hình vẫn phát hiện được các phương tiện có kích thước đủ lớn và biên tương "
        "đối rõ, nhưng các đối tượng nhỏ ở xa hoặc hòa vào vùng tối dễ bị bỏ sót. Những trường hợp này cho thấy tăng độ phân "
        "giải hỗ trợ giữ chi tiết nhưng không thay thế được dữ liệu đại diện cho từng lớp và từng điều kiện chiếu sáng."
    )
    builder.picture(FIGURES / "hinh_4_4_detect_acdc.jpg", width=6.2)
    cap("Hình 4.5. Ví dụ kết quả phát hiện của mô hình final trên ACDC val")

    h2("4.5. Các thực nghiệm bổ trợ")
    h3("4.5.1. YOLOv8s như mốc trần độ chính xác")
    p(
        "YOLOv8s được huấn luyện với cùng chiến lược Phase 2 v3 và imgsz 960 để kiểm tra ảnh hưởng của việc tăng năng lực mô "
        "hình. Cấu hình này cải thiện mAP50-95 trên cả ba tập: XWOD tăng từ 0,279 lên 0,305; DAWN tăng từ 0,541 lên 0,564; "
        "ACDC tăng từ 0,239 lên 0,269. Mức tăng nhất quán cho thấy năng lực biểu diễn vẫn là một yếu tố ảnh hưởng đến độ chính xác."
    )
    p(
        "Tuy nhiên, recall ACDC của YOLOv8s là 0,389, gần như không đổi so với 0,390 của YOLOv8n. Mô hình lớn hơn cải thiện "
        "chủ yếu khả năng phân loại và định vị các đối tượng đã phát hiện, nhưng không làm giảm rõ rệt số đối tượng bị bỏ sót "
        "trên ACDC. YOLOv8s vì thế được giữ làm mốc trần độ chính xác, không thay thế mô hình thực dụng cuối cùng."
    )

    h3("4.5.2. Kết quả copy-paste")
    p(
        "Run copy-paste @960 tạo kết quả thấp hơn YOLOv8n v3_960 trên cả ba tập. Recall giảm từ 0,472 xuống 0,446 trên "
        "XWOD, từ 0,731 xuống 0,704 trên DAWN và từ 0,390 xuống 0,361 trên ACDC. mAP50-95 cũng giảm tương ứng. Vì xu hướng "
        "suy giảm xuất hiện đồng thời ở ba miền, copy-paste không được chọn cho cấu hình cuối."
    )
    cap("Bảng 4.11. So sánh mô hình final, YOLOv8s và copy-paste @960")
    table(
        ["Cấu hình", "Tập", "Precision", "Recall", "mAP50", "mAP50-95"],
        [
            ["YOLOv8n v3_960", "XWOD", "0,6490", "0,4720", "0,5040", "0,2790"],
            ["YOLOv8n v3_960", "DAWN", "0,8010", "0,7310", "0,7880", "0,5410"],
            ["YOLOv8n v3_960", "ACDC", "0,5410", "0,3900", "0,3990", "0,2390"],
            ["YOLOv8s v3_960", "XWOD", "0,6584", "0,4940", "0,5311", "0,3049"],
            ["YOLOv8s v3_960", "DAWN", "0,7922", "0,7673", "0,8016", "0,5637"],
            ["YOLOv8s v3_960", "ACDC", "0,6197", "0,3893", "0,4314", "0,2688"],
            ["Copy-paste v5", "XWOD", "0,6317", "0,4463", "0,4891", "0,2777"],
            ["Copy-paste v5", "DAWN", "0,7221", "0,7042", "0,7362", "0,5100"],
            ["Copy-paste v5", "ACDC", "0,5368", "0,3606", "0,3841", "0,2308"],
        ],
        widths=[1.35, 0.8, 0.95, 0.85, 0.9, 1.0],
    )
    p(
        "Các ảnh ghép cứng có thể tạo biên cắt, tỉ lệ, ánh sáng hoặc bối cảnh không tự nhiên; đây là những giả thuyết hợp lý "
        "để giải thích xu hướng nhưng chưa được kiểm chứng bằng một thí nghiệm tách riêng từng yếu tố. Kết luận được giới hạn "
        "ở mức thực nghiệm: cấu hình copy-paste đã dùng không cải thiện mô hình và bị loại."
    )

    h2("4.6. Thảo luận")
    h3("4.6.1. Các phát hiện chính và lựa chọn mô hình")
    p(
        "Benchmark cho thấy RT-DETR dẫn đầu về độ chính xác, trong khi YOLOv8n tạo được sự cân bằng tốt hơn cho mục tiêu triển "
        "khai. Các bước Phase 2 sau đó cải thiện rõ DAWN và ACDC mà chỉ làm thay đổi nhỏ XWOD. Kết quả này cho thấy giới hạn "
        "ban đầu của YOLOv8n không chỉ nằm ở kiến trúc; thành phần dữ liệu, phân bố lớp và độ phân giải đầu vào có ảnh hưởng "
        "trực tiếp đến khả năng chuyển miền."
    )
    p(
        "SE không tạo cải thiện nhất quán và bị loại. Copy-paste cũng là một kết quả âm. Ngược lại, tăng độ phân giải và tăng "
        "năng lực mô hình đều đem lại lợi ích về mAP. Tuy vậy, YOLOv8s không cải thiện recall ACDC, cho thấy bài toán bỏ sót "
        "đối tượng nhỏ hoặc hiếm chưa thể giải quyết chỉ bằng cách tăng số tham số. Cấu hình cuối do đó vẫn là YOLOv8n Phase "
        "2 v3 @960; YOLOv8s được giữ như tham chiếu về độ chính xác."
    )

    h3("4.6.2. Hạn chế và hướng phát triển")
    p(
        "Hạn chế rõ nhất nằm ở bicycle và motorcycle trên ACDC, cùng recall tổng thể 0,390. Quy mô DAWN nhỏ, đặc biệt chỉ có "
        "sáu đối tượng bicycle, cũng làm tăng độ bất định của phân tích theo lớp. Bên cạnh đó, full text log của benchmark "
        "XWOD test bốn mô hình chưa đầy đủ đồng đều; quan hệ checkpoint giữa v3_rare và v3_960 chưa được xác nhận. Các khoảng "
        "trống này không làm mất hiệu lực của lần tái đánh giá final nhưng giới hạn mức độ truy vết của một số bảng trung gian."
    )
    p(
        "Hướng phát triển phù hợp là bổ sung dữ liệu thật cho lớp hiếm, xây dựng phép tăng cường có kiểm soát về chiếu sáng và "
        "bối cảnh, đồng thời đánh giá theo từng điều kiện thời tiết bằng một giao thức đã khóa trước. Có thể khảo sát cơ chế "
        "chú ý khác như CBAM hoặc các kiến trúc lớn hơn, nhưng cần xem đây là nghiên cứu mới thay vì kết quả đã có. Đối với "
        "triển khai, cần đo lại tốc độ của checkpoint final @960 trên phần cứng đích, vì số ms/ảnh ở benchmark 640 không đại "
        "diện trực tiếp cho cấu hình cuối."
    )
    p(
        "Tổng hợp các kết quả cho thấy đề tài đã xây dựng được một quy trình thực nghiệm có thể truy vết, xác định được mô "
        "hình thực dụng và ghi nhận cả những hướng cải tiến không thành công. Giá trị của kết quả không chỉ nằm ở mAP cuối, "
        "mà còn ở việc phân biệt rõ ảnh hưởng của kiến trúc, dữ liệu, độ phân giải và năng lực mô hình trong điều kiện thời "
        "tiết bất lợi."
    )


def main(path_str: str) -> None:
    path = Path(path_str)
    document = Document(path)
    _, references = remove_chapter4_body(document)
    builder = ChapterBuilder(document, references)
    add_chapter4(builder)
    builder.commit()
    update_manual_figure_table_lists(document)
    reset_toc_cache(document)
    mark_fields_dirty(document)
    document.save(path)
    print(f"Đã ghi bản Chương 4 học thuật: {path}")


if __name__ == "__main__":
    main(sys.argv[1])
