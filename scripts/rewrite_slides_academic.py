from pathlib import Path
import zipfile

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


SOURCE = Path("/Users/thanhtin/Documents/KLTN_LamThanhBao/Slide.pptx")
OUTPUT = Path("/Users/thanhtin/Documents/Object Detection /deliverables/Slide_REVISED_AI_STORY_FINAL.pptx")

FONT = "Arial"
TEXT = RGBColor(48, 54, 61)
BLUE = RGBColor(34, 112, 184)
GREEN = RGBColor(31, 122, 78)
ORANGE = RGBColor(221, 120, 43)
RED = RGBColor(184, 45, 45)


def body_shape(slide):
    for shape in slide.shapes:
        if shape.name.startswith("Text Placeholder"):
            return shape
    raise RuntimeError("Body text box not found")


def set_title(slide, text):
    shape = slide.shapes[0]
    shape.text = text
    for paragraph in shape.text_frame.paragraphs:
        paragraph.font.name = FONT
        paragraph.font.size = Pt(25)
        paragraph.font.bold = True
        paragraph.font.color.rgb = RGBColor(0, 0, 0)


def set_body(slide, items, *, font_size=18, top=None, height=None):
    """items: list[(text, kind)] where kind is normal/bold/green/orange/red."""
    shape = body_shape(slide)
    if top is not None:
        shape.top = Inches(top)
    if height is not None:
        shape.height = Inches(height)
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    colors = {
        "normal": TEXT,
        "bold": TEXT,
        "green": GREEN,
        "orange": ORANGE,
        "red": RED,
        "blue": BLUE,
    }
    for idx, (text, kind) in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = text
        p.font.name = FONT
        p.font.size = Pt(font_size)
        p.font.bold = kind in {"bold", "green", "orange", "red", "blue"}
        p.font.color.rgb = colors[kind]
        p.space_after = Pt(5 if kind == "normal" else 6)
        p.line_spacing = 1.05
        p.alignment = PP_ALIGN.LEFT


def style_table(table, font_size=13):
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            cell.margin_left = Inches(0.05)
            cell.margin_right = Inches(0.05)
            cell.margin_top = Inches(0.03)
            cell.margin_bottom = Inches(0.03)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.name = FONT
                paragraph.font.size = Pt(font_size)
                paragraph.font.bold = r_idx == 0
                paragraph.alignment = PP_ALIGN.CENTER
                if r_idx == 0:
                    paragraph.font.color.rgb = RGBColor(255, 255, 255)


def set_table(slide, rows, font_size=13):
    table_shape = next(s for s in slide.shapes if getattr(s, "has_table", False))
    table = table_shape.table
    if len(rows) != len(table.rows) or len(rows[0]) != len(table.columns):
        raise ValueError(f"Table size mismatch on {slide}: expected {len(table.rows)}x{len(table.columns)}")
    for r, values in enumerate(rows):
        for c, value in enumerate(values):
            table.cell(r, c).text = str(value)
    style_table(table, font_size)


def clear_body(slide):
    shape = body_shape(slide)
    shape.text = ""
    shape.height = Inches(0.1)


def remove_shape(shape):
    """Remove a shape from a slide while preserving the surrounding template."""
    element = shape._element
    element.getparent().remove(element)


def add_card(slide, left, top, width, height, color, title, lines):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(247, 250, 253)
    shape.line.color.rgb = color
    shape.line.width = Pt(2)
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.16)
    tf.margin_right = Inches(0.16)
    tf.margin_top = Inches(0.13)
    tf.margin_bottom = Inches(0.10)
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = color
    p.space_after = Pt(9)
    for line in lines:
        p = tf.add_paragraph()
        p.text = line
        p.font.name = FONT
        p.font.size = Pt(14)
        p.font.color.rgb = TEXT
        p.space_after = Pt(7)
        p.line_spacing = 1.05
    return shape


def revise(prs):
    # Slide 2
    set_body(prs.slides[1], [
        ("1.  Giới thiệu tổng quan đề tài", "bold"),
        ("2.  Cơ sở lý thuyết", "normal"),
        ("3.  Bộ dữ liệu sử dụng", "normal"),
        ("4.  Phân tích, thiết kế và giao thức đánh giá", "normal"),
        ("5.  Triển khai và thực nghiệm", "normal"),
        ("6.  Kết luận và hướng phát triển", "normal"),
        ("7.  Demo chương trình", "normal"),
    ], font_size=19)

    # Slide 3
    slide = prs.slides[2]
    set_title(slide, "Phạm vi và mục tiêu của đề tài")
    clear_body(slide)

    intro = slide.shapes.add_textbox(Inches(0.65), Inches(1.02), Inches(12.0), Inches(0.64))
    intro.text = "PHÁT HIỆN ĐỐI TƯỢNG GIAO THÔNG TRONG ĐIỀU KIỆN THỜI TIẾT KHẮC NGHIỆT"
    for p in intro.text_frame.paragraphs:
        p.font.name = FONT; p.font.size = Pt(21); p.font.bold = True; p.font.color.rgb = BLUE
        p.alignment = PP_ALIGN.CENTER

    add_card(slide, 0.65, 1.92, 3.75, 2.55, BLUE, "Đầu vào", [
        "Ảnh giao thông từ camera gắn trên xe và nhiều nguồn khác.",
        "Điều kiện clear, fog, rain, snow, night và thời tiết bất lợi khác.",
    ])
    add_card(slide, 4.78, 1.92, 3.75, 2.55, ORANGE, "Đối tượng", [
        "6 lớp: person, bicycle, car, motorcycle, bus và truck.",
        "Yêu cầu nhận diện đồng thời đối tượng lớn, nhỏ và lớp hiếm.",
    ])
    add_card(slide, 8.90, 1.92, 3.75, 2.55, GREEN, "Mục tiêu", [
        "So sánh 4 mô hình thuộc 3 họ kiến trúc.",
        "Xây dựng quy trình huấn luyện thích nghi thời tiết nhưng vẫn giữ miền thông thường.",
    ])

    problem = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.80), Inches(4.92), Inches(11.70), Inches(1.18)
    )
    problem.fill.solid(); problem.fill.fore_color.rgb = RGBColor(255, 242, 232)
    problem.line.color.rgb = ORANGE
    problem.text = "CÂU HỎI NGHIÊN CỨU: kiến trúc và quy trình huấn luyện nào cải thiện thời tiết bất lợi mà vẫn duy trì năng lực trên miền thông thường?"
    for p in problem.text_frame.paragraphs:
        p.font.name = FONT; p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = RED
        p.alignment = PP_ALIGN.CENTER

    # Slide 4
    slide = prs.slides[3]
    set_title(slide, "Một dẫn chứng thực tế: sương mù làm tăng số phương tiện bị bỏ sót")
    clear_body(slide)

    data = [
        ["Điều kiện", "Clear", "Light fog", "Moderate fog", "Heavy fog"],
        ["Recall Faster R-CNN", "91,55%", "85,21%", "72,54–64,79%", "≤ 57,75%"],
    ]
    table_shape = slide.shapes.add_table(2, 5, Inches(0.65), Inches(1.20), Inches(12.0), Inches(1.35))
    table = table_shape.table
    widths = [2.30, 1.65, 1.85, 3.10, 3.10]
    for col, width in zip(table.columns, widths): col.width = Inches(width)
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            cell = table.cell(r, c); cell.text = value
            cell.fill.solid(); cell.fill.fore_color.rgb = BLUE if r == 0 else RGBColor(241, 246, 250)
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT; p.font.size = Pt(14); p.font.bold = True
                p.font.color.rgb = RGBColor(255, 255, 255) if r == 0 else (RED if c == 4 else TEXT)
                p.alignment = PP_ALIGN.CENTER

    add_card(slide, 0.65, 2.88, 5.72, 2.22, ORANGE, "Ảnh sương mù thực tế từ BDD100K", [
        "Moderate fog: 172 xe → bỏ sót 55 xe (32,0%).",
        "Heavy fog: 113 xe → bỏ sót 47 xe (41,6%).",
    ])
    add_card(slide, 6.70, 2.88, 5.72, 2.22, RED, "Thông điệp", [
        "Trong sương mù dày, cứ 10 phương tiện xuất hiện thì detector bỏ sót khoảng 4.",
        "Kết quả thực cảnh này đánh giá vehicle, chưa đánh giá pedestrian.",
    ])

    bridge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(5.35), Inches(11.68), Inches(0.82)
    )
    bridge.fill.solid(); bridge.fill.fore_color.rgb = RGBColor(231, 242, 251)
    bridge.line.color.rgb = BLUE
    bridge.text = (
        "Từ đó, đề tài tìm kiến trúc và quy trình huấn luyện giúp phát hiện 6 lớp tốt hơn trong thời tiết bất lợi mà vẫn giữ năng lực trên miền thông thường."
    )
    for p in bridge.text_frame.paragraphs:
        p.font.name = FONT; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = BLUE
        p.alignment = PP_ALIGN.CENTER

    note = slide.shapes.add_textbox(Inches(0.65), Inches(6.30), Inches(12.0), Inches(0.42))
    note.text = (
        "Nguồn: Liu et al., Sensors 20(2):349, 2020, Bảng 3–4. Recall clear–fog từ mô phỏng; số xe bỏ sót từ ảnh BDD100K thực tế."
    )
    for p in note.text_frame.paragraphs:
        p.font.name = FONT; p.font.size = Pt(10); p.font.italic = True; p.font.color.rgb = RGBColor(90, 90, 90)

    # Slide 5
    set_title(prs.slides[4], "2. Cơ sở lý thuyết — ba hướng phát hiện đối tượng")
    set_body(prs.slides[4], [
        ("Cùng một bài toán 6 lớp, đề tài khảo sát ba họ kiến trúc có cơ chế và đánh đổi accuracy–speed khác nhau.", "bold"),
        ("Bốn mô hình đại diện được huấn luyện và đánh giá theo cùng không gian nhãn.", "normal"),
    ], font_size=17, height=1.65)
    set_table(prs.slides[4], [
        ["Hướng", "Mô hình trong đề tài", "Vai trò khi so sánh"],
        ["One-stage", "YOLOv8n, YOLO11n", "Nhẹ, nhanh; baseline cho yêu cầu gần thời gian thực"],
        ["Two-stage", "Faster R-CNN", "Sinh region proposal trước; mốc tham chiếu về độ chính xác"],
        ["Transformer", "RT-DETR-L", "Phát hiện end-to-end; ứng viên cho Phase 2"],
    ], font_size=13)

    # Slide 6
    set_title(prs.slides[5], "4. Phân tích và thiết kế hệ thống — pipeline tổng thể")
    set_body(prs.slides[5], [
        ("COCO pretrained → BDD100K → XWOD → Phase 2 đa miền → A0R/A1-DINO.", "bold"),
        ("Mọi giai đoạn dùng chung 6 lớp; validation chọn checkpoint, test chỉ báo cáo sau khi cấu hình đã khóa.", "green"),
    ], font_size=16, height=1.35)

    # Slide 7 — presented as Stage 2 after reordering
    set_title(prs.slides[6], "Stage 2 — thích nghi với thời tiết khắc nghiệt")
    set_body(prs.slides[6], [
        ("LÀM GÌ? Tiếp tục huấn luyện cả 4 mô hình trên XWOD — 6.006 ảnh thuộc 7 nhóm thời tiết khắc nghiệt.", "blue"),
        ("KẾT QUẢ RT-DETR-L: XWOD 0,291 → 0,500; DAWN 0,344 → 0,526 mAP50–95.", "green"),
        ("VẤN ĐỀ PHÁT SINH: BDD giảm 0,362 → 0,183 → mô hình thích nghi miền đích nhưng quên miền nguồn.", "red"),
    ], font_size=16, height=1.58)

    # Slide 8 — full model-selection evidence after Stage 2
    slide = prs.slides[7]
    set_title(slide, "Chọn best model sau Stage 2 — số liệu đầy đủ")
    set_body(prs.slides[7], [
        ("Tiêu chí quyết định là mAP50–95 trên XWOD validation; các test set chỉ dùng để kiểm chứng sau khi chọn.", "bold"),
    ], font_size=15, height=0.85)
    for shape in list(slide.shapes):
        if shape.shape_type == 13:  # picture from the original slide
            remove_shape(shape)
    values = [
        ["Mô hình", "XWOD val", "XWOD test", "DAWN test", "BDD test", "ACDC test", "FPS"],
        ["RT-DETR-L", "0,559", "0,500", "0,526", "0,183", "0,128", "33"],
        ["YOLO11n", "0,506", "0,454", "0,400", "0,198", "0,161", "96"],
        ["YOLOv8n", "0,503", "0,452", "0,414", "0,199", "0,155", "112"],
        ["Faster R-CNN", "0,421", "0,381", "0,449", "0,127", "0,107", "61"],
    ]
    table_shape = slide.shapes.add_table(5, 7, Inches(0.45), Inches(2.03), Inches(12.45), Inches(3.55))
    table = table_shape.table
    widths = [2.25, 1.70, 1.75, 1.75, 1.65, 1.65, 1.70]
    for col, width in zip(table.columns, widths): col.width = Inches(width)
    for r, row in enumerate(values):
        for c, value in enumerate(row):
            cell = table.cell(r, c); cell.text = value
            cell.fill.solid()
            cell.fill.fore_color.rgb = BLUE if r == 0 else (RGBColor(225, 244, 234) if r == 1 else RGBColor(248, 249, 250))
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT; p.font.size = Pt(12); p.font.bold = r in {0, 1}
                p.font.color.rgb = RGBColor(255, 255, 255) if r == 0 else (GREEN if r == 1 else TEXT)
                p.alignment = PP_ALIGN.CENTER
    decision = slide.shapes.add_textbox(Inches(0.65), Inches(5.85), Inches(12.0), Inches(0.60))
    decision.text = "KẾT LUẬN: RT-DETR-L dẫn đầu XWOD validation và hai test thời tiết XWOD, DAWN → được chọn cho Phase 2."
    for p in decision.text_frame.paragraphs:
        p.font.name = FONT; p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = GREEN
        p.alignment = PP_ALIGN.CENTER

    # Slide 9
    set_title(prs.slides[8], "Từ nguyên nhân đến các hướng thử nghiệm")
    set_body(prs.slides[8], [
        ("Chẩn đoán: bicycle/motorcycle/bus ít mẫu; oversampling chỉ lặp lại dữ liệu cũ; nhiều vật thể quá nhỏ sau resize 640.", "bold"),
        ("Hướng thử 1 — A0R: thêm ngẫu nhiên 5.000 ảnh BDD có lớp hiếm để đo lợi ích của lượng dữ liệu.", "normal"),
        ("Hướng thử 2 — A1-DINO: truy hồi 5.000 ảnh tương đồng với mẫu khó; giữ cùng checkpoint và training budget.", "green"),
    ], font_size=16, height=2.1)
    set_table(prs.slides[8], [
        ["Nhánh", "Cách chọn 5.000 ảnh", "Vai trò"],
        ["A0R", "Ngẫu nhiên sau khi lọc ảnh có bicycle/motorcycle/bus", "Đối chứng lượng dữ liệu"],
        ["A1-DINO", "Xếp hạng theo độ tương đồng DINOv2 với mẫu khó", "Đánh giá chiến lược truy hồi"],
    ], font_size=12)

    # Slide 10
    set_title(prs.slides[9], "3. Bộ dữ liệu sử dụng")
    set_body(prs.slides[9], [
        ("Tất cả được chuẩn hóa về cùng thứ tự class ID: person, bicycle, car, motorcycle, bus, truck.", "bold"),
        ("Nguồn ảnh: BDD100K dashcam; XWOD hỗn hợp dashcam + web; ACDC camera trên xe; DAWN web-sourced và chỉ dùng đánh giá.", "orange"),
    ], font_size=16, height=1.45)
    set_table(prs.slides[9], [
        ["Bộ dữ liệu", "Vai trò", "Điều kiện", "Train / Val / Test"],
        ["BDD100K", "Miền nguồn; Stage 1", "Giao thông đa dạng", "30.000 / 3.000 / 7.000"],
        ["XWOD", "Miền đích; Stage 2", "7 nhóm thời tiết", "6.006 / 1.001 / 3.003"],
        ["ACDC", "Thích nghi đa miền", "Fog, rain, snow, night", "1.182 / 197 / 591"],
        ["DAWN", "Đánh giá ngoài miền", "Fog, rain, sand, snow", "— / 308 / 718"],
    ], font_size=11)

    # Slide 11 — presented as the first experiment slide after reordering
    set_title(prs.slides[10], "5. Triển khai và thực nghiệm — Stage 1")
    set_body(prs.slides[10], [
        ("LÀM GÌ? Fine-tune cả 4 mô hình pretrained COCO trên 30.000 ảnh BDD100K.", "blue"),
        ("MỤC ĐÍCH: chuyển từ đặc trưng COCO sang bối cảnh giao thông và góc nhìn camera trên xe.", "normal"),
        ("KẾT QUẢ RT-DETR-L", "green"),
        ("BDD 0,362 · XWOD 0,291 · DAWN 0,344 · ACDC 0,174 mAP50–95.", "bold"),
        ("NHẬN XÉT: mô hình đã học miền giao thông nhưng hiệu năng trên thời tiết bất lợi còn thấp → cần Stage 2.", "orange"),
    ], font_size=18, top=1.20, height=4.85)

    # Slide 12
    set_title(prs.slides[11], "Độ đo sử dụng")
    set_body(prs.slides[11], [
        ("Độ đo chính: mAP50–95; bổ sung Precision, Recall, mAP50, kết quả theo lớp và FPS.", "bold"),
        ("Môi trường: RTX 5090 32 GB · ảnh 640×640 · tốc độ đo batch 1. Phase 2: AdamW, batch 16, lr0=5×10⁻⁵, seed 42.", "normal"),
    ], font_size=16, height=1.4)

    # Slide 13 — Phase 2 setup after model selection
    slide = prs.slides[12]
    set_title(slide, "Phase 2 — khắc phục quên miền bằng dữ liệu đa miền")
    clear_body(slide)
    for shape in list(slide.shapes):
        if getattr(shape, "has_table", False):
            remove_shape(shape)
    add_card(slide, 0.65, 1.42, 3.75, 3.15, BLUE, "XWOD — giữ miền đích", [
        "6.006 ảnh train.",
        "Giữ năng lực phát hiện trong 7 nhóm thời tiết khắc nghiệt.",
    ])
    add_card(slide, 4.78, 1.42, 3.75, 3.15, GREEN, "BDD100K — phục hồi miền nguồn", [
        "30.000 ảnh train.",
        "Giảm hiện tượng quên miền giao thông thông thường sau Stage 2.",
    ])
    add_card(slide, 8.90, 1.42, 3.75, 3.15, ORANGE, "ACDC — bổ sung điều kiện khó", [
        "1.182 ảnh train.",
        "Bổ sung fog, rain, snow và night từ camera trên xe.",
    ])
    phase_note = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.82), Inches(4.95), Inches(11.68), Inches(1.10)
    )
    phase_note.fill.solid(); phase_note.fill.fore_color.rgb = RGBColor(241, 246, 250)
    phase_note.line.color.rgb = BLUE
    phase_note.text = "Tổng: 37.188 ảnh gốc → oversampling lớp hiếm thành 50.985 ảnh. DAWN không tham gia train, chỉ dùng đánh giá ngoài miền."
    for p in phase_note.text_frame.paragraphs:
        p.font.name = FONT; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = TEXT
        p.alignment = PP_ALIGN.CENTER

    # Slide 14
    set_title(prs.slides[13], "Phase 2 — kết quả cân bằng đa miền")
    set_body(prs.slides[13], [
        ("So với Stage 2, Phase 2 giữ XWOD 0,500 → 0,510 và nâng ACDC 0,128 → 0,247.", "green"),
        ("BDD phục hồi 0,183 → 0,354, tương đương khoảng 96% phần hiệu năng đã mất; DAWN giảm nhẹ 0,526 → 0,515.", "normal"),
    ], font_size=16, height=1.3)

    # Slide 15
    set_title(prs.slides[14], "Phase 2 — số liệu tổng thể")
    set_body(prs.slides[14], [
        ("Mô hình cuối cải thiện mạnh trên ba miền bất lợi nhưng gần như duy trì toàn bộ hiệu năng BDD.", "bold"),
        ("Recall XWOD tăng 0,481 → 0,745; mAP50–95 là chỉ số chính của bảng.", "green"),
    ], font_size=16, height=1.4)

    # Slide 16
    set_title(prs.slides[15], "Sau Phase 2, metrics vẫn thấp ở đâu?")
    set_body(prs.slides[15], [
        ("Bicycle, motorcycle và bus là các lớp yếu; car ổn định nhất do có nhiều mẫu và kích thước lớn hơn.", "bold"),
        ("ACDC night chỉ đạt 0,121; recall vật thể nhỏ dao động 0,000–0,696 → nghi vấn chính: thiếu mẫu, thiếu sáng và mất chi tiết sau resize.", "orange"),
    ], font_size=15, height=0.75)

    # Slide 17
    set_title(prs.slides[16], "Kết quả các hướng thử nghiệm")
    set_body(prs.slides[16], [
        ("V1 — mAP50–95 trung bình: A0R 0,4098 > Phase 2 0,4062 > A1-DINO 0,4056.", "bold"),
        ("V2 (đánh giá lại cùng môi trường): A0R 0,4191 > A1-DINO v2 0,4180 > Phase 2 0,4162.", "blue"),
        ("Chênh lệch chỉ 0,0011–0,0042 và mỗi cấu hình một seed: chưa đủ kết luận ý nghĩa thống kê.", "red"),
        ("Trên ACDC, recall vật thể nhỏ 0,000–0,696 trong khi vật thể lớn đạt 1,000; kích thước tại đầu vào là ràng buộc nổi bật.", "green"),
    ], font_size=14, height=2.05)
    pic = prs.slides[16].shapes[3]
    pic.left = Inches(1.7)
    pic.top = Inches(3.0)
    pic.width = Inches(9.6)
    pic.height = Inches(3.9)

    # Slide 18
    set_title(prs.slides[17], "6. Kết luận và hướng phát triển")
    set_body(prs.slides[17], [
        ("1 · KHẢ NĂNG NHẬN DIỆN TRONG THỜI TIẾT XẤU — CẢI THIỆN RÕ", "blue"),
        ("XWOD 0,291 → 0,510; ACDC 0,174 → 0,247 mAP50–95.", "normal"),
        ("2 · KHẢ NĂNG TỔNG QUÁT TRÊN MIỀN THÔNG THƯỜNG — ĐƯỢC DUY TRÌ", "green"),
        ("BDD sau Stage 2 chỉ 0,183; Phase 2 phục hồi lên 0,354, tương đương khoảng 96% phần đã mất.", "normal"),
        ("3 · LỚP HIẾM/VẬT THỂ NHỎ — CHƯA GIẢI QUYẾT TRIỆT ĐỂ", "orange"),
        ("A0R và hai phiên bản DINOv2 chỉ tạo chênh lệch nhỏ; ACDC night và vật thể nhỏ vẫn là điểm yếu.", "normal"),
        ("Hạn chế: một seed · RT-DETR-L nặng/chậm hơn YOLO nano · khác biệt framework · A1-v2 khác phiên bản Ultralytics.", "red"),
        ("Hướng tiếp theo: ảnh gốc + độ phân giải hiệu dụng cao hơn/tiled inference · nhiều seed · balanced replay · tối ưu triển khai.", "bold"),
    ], font_size=16)

    # Slide 19
    set_title(prs.slides[18], "7. Demo chương trình")
    set_body(prs.slides[18], [
        ("Ảnh/video → letterbox 640×640 → RT-DETR-L Phase 2 → hộp, nhãn và độ tin cậy", "bold"),
        ("Thông lượng khoảng 32–35 FPS trên RTX 5090, batch 1; ngưỡng hiển thị không làm thay đổi số liệu mAP đã báo cáo.", "normal"),
        ("Demo trình bày cả trường hợp phát hiện tốt và trường hợp bỏ sót vật thể nhỏ/thiếu sáng.", "orange"),
    ], font_size=16, height=1.2)

    # New opening slide: explain the research problem before evidence and scope.
    problem_slide = prs.slides.add_slide(prs.slides[2].slide_layout)
    set_title(problem_slide, "1. Bài toán nghiên cứu — suy giảm độ chính xác dưới thời tiết bất lợi")
    content_placeholder = next(
        shape for shape in problem_slide.placeholders
        if shape.placeholder_format.idx == 1
    )
    content_placeholder.text = ""
    content_placeholder.height = Inches(0.1)

    # Causal chain: good-weather training -> visual degradation -> domain shift -> missed detections.
    flow = [
        (0.55, BLUE, "Điều kiện thuận lợi", "Ảnh rõ, đủ sáng\nđộ tương phản cao"),
        (3.72, ORANGE, "Thời tiết bất lợi", "Fog · rain · snow · night\nche mờ, nhiễu, thiếu sáng"),
        (6.89, RED, "Domain shift", "Đặc trưng đầu vào khác\nphân phối dữ liệu huấn luyện"),
        (10.06, GREEN, "Hệ quả", "Confidence và Recall giảm\ntăng đối tượng bị bỏ sót"),
    ]
    for idx, (left, color, heading, detail) in enumerate(flow):
        card = problem_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(1.42), Inches(2.72), Inches(1.72)
        )
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(247, 250, 253)
        card.line.color.rgb = color; card.line.width = Pt(2)
        tf = card.text_frame; tf.clear(); tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = heading
        p.font.name = FONT; p.font.size = Pt(17); p.font.bold = True; p.font.color.rgb = color
        p.alignment = PP_ALIGN.CENTER; p.space_after = Pt(8)
        p = tf.add_paragraph(); p.text = detail
        p.font.name = FONT; p.font.size = Pt(13); p.font.color.rgb = TEXT
        p.alignment = PP_ALIGN.CENTER
        if idx < len(flow) - 1:
            arrow = problem_slide.shapes.add_shape(
                MSO_SHAPE.CHEVRON, Inches(left + 2.77), Inches(1.98), Inches(0.34), Inches(0.52)
            )
            arrow.fill.solid(); arrow.fill.fore_color.rgb = RGBColor(130, 138, 145)
            arrow.line.fill.background()

    add_card(problem_slide, 0.70, 3.58, 5.82, 2.15, RED, "Bài toán là gì?", [
        "Xây dựng detector nhận diện ổn định người và phương tiện khi điều kiện quan sát thay đổi.",
        "Không chỉ tăng kết quả trên thời tiết xấu mà còn phải hạn chế suy giảm trên thời tiết thông thường.",
    ])
    add_card(problem_slide, 6.82, 3.58, 5.82, 2.15, BLUE, "Đề tài làm gì?", [
        "So sánh 3 hướng, 4 mô hình trên cùng 6 lớp và cùng giao thức đánh giá.",
        "Huấn luyện nhiều giai đoạn, phối hợp nhiều miền dữ liệu và phân tích nguyên nhân metrics còn thấp.",
    ])
    takeaway = problem_slide.shapes.add_textbox(Inches(0.75), Inches(6.00), Inches(11.85), Inches(0.48))
    takeaway.text = "Mục tiêu cốt lõi: tăng tính bền vững trước thời tiết bất lợi, không đánh đổi toàn bộ năng lực trên miền ban đầu."
    for p in takeaway.text_frame.paragraphs:
        p.font.name = FONT; p.font.size = Pt(15); p.font.bold = True; p.font.color.rgb = GREEN
        p.alignment = PP_ALIGN.CENTER

    # Reorder the existing slides into the narrative requested by the user:
    # problem -> model families -> datasets -> pipeline/metrics/protocol -> experiments.
    # Section 1: problem theory -> real-world evidence -> concrete project scope.
    order = [0, 1, 20, 3, 2, 4, 9, 5, 11, 10, 6, 7, 12, 13, 14, 15, 8, 16, 17, 18, 19]
    slide_ids = list(prs.slides._sldIdLst)
    for slide_id in slide_ids:
        prs.slides._sldIdLst.remove(slide_id)
    for idx in order:
        prs.slides._sldIdLst.append(slide_ids[idx])

    # The template uses small text boxes for page numbers; refresh them after reordering.
    for page_no, current_slide in enumerate(prs.slides, start=1):
        for shape in current_slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            if shape.top < Inches(6.75):
                continue
            value = shape.text.strip()
            if value.isdigit() and len(value) <= 2:
                shape.text = str(page_no)
                for p in shape.text_frame.paragraphs:
                    p.font.name = FONT
                    p.font.size = Pt(11)
                    p.font.color.rgb = TEXT


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation(SOURCE)
    revise(prs)
    prs.save(OUTPUT)
    # python-pptx may retain duplicate slide members from packages containing
    # embedded media. Repack once so PowerPoint/LibreOffice sees one member per
    # OOXML part while preserving the latest (edited) copy.
    repacked = OUTPUT.with_suffix(".repacked.pptx")
    with zipfile.ZipFile(OUTPUT, "r") as src:
        last_entry = {info.filename: info for info in src.infolist()}
        with zipfile.ZipFile(repacked, "w") as dst:
            for name, info in last_entry.items():
                dst.writestr(info, src.read(info), compress_type=zipfile.ZIP_DEFLATED)
    repacked.replace(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
