import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

def create_document():
    doc = docx.Document()
    
    # Page setup (A4 / Letter with 1 inch margins)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        
    # Styles
    COLOR_PRIMARY = RGBColor(31, 78, 121)    # #1F4E79 Dark Blue
    COLOR_SECONDARY = RGBColor(47, 85, 151) # #2F5597 Steel Blue
    COLOR_DARK = RGBColor(38, 38, 38)       # #262626 Charcoal
    COLOR_ACCENT = RGBColor(192, 0, 0)      # #C00000 Crimson Red
    
    # Set default font to Times New Roman
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(13)
    style_normal.font.color.rgb = COLOR_DARK
    style_normal.paragraph_format.line_spacing = 1.2
    style_normal.paragraph_format.space_after = Pt(4)
    style_normal.paragraph_format.space_before = Pt(0)
    
    # Helpers
    def add_p(text="", bold=False, italic=False, color=COLOR_DARK, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=4, font_size=13):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.2
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = color
        return p

    def add_h1(text):
        p = add_p(text, bold=True, color=COLOR_PRIMARY, space_before=16, space_after=8, font_size=16)
        return p

    def add_h2(text):
        p = add_p(text, bold=True, color=COLOR_SECONDARY, space_before=12, space_after=6, font_size=14)
        return p

    def add_h3(text):
        p = add_p(text, bold=True, color=RGBColor(32, 55, 100), space_before=8, space_after=4, font_size=13)
        return p

    def add_callout(title, text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F7"/>'))
        
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:left w:val="single" w:sz="24" w:color="1F4E79"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        r_t = p.add_run(f"📌 {title}: ")
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(12)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_PRIMARY
        
        r_c = p.add_run(text)
        r_c.font.name = 'Times New Roman'
        r_c.font.size = Pt(12)
        r_c.font.italic = True
        r_c.font.color.rgb = COLOR_DARK
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def format_table(tbl, col_widths, headers, data, align_center=True):
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblPr = tbl._tbl.tblPr
        tblPr.append(parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="6" w:color="2F5597"/>
                <w:bottom w:val="single" w:sz="6" w:color="2F5597"/>
                <w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/>
                <w:insideV w:val="none"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tblBorders>
        '''))
        
        # Header row
        hdr_row = tbl.rows[0]
        for idx, h_text in enumerate(headers):
            cell = hdr_row.cells[idx]
            cell.width = col_widths[idx]
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="2F5597"/>'))
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(h_text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(11.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            
        # Data rows
        for r_idx, row_data in enumerate(data):
            row = tbl.add_row()
            bg_color = "F9FAFB" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row_data):
                cell = row.cells[c_idx]
                cell.width = col_widths[c_idx]
                if bg_color != "FFFFFF":
                    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>'))
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align_center else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                r = p.add_run(str(val))
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)
                r.font.color.rgb = COLOR_DARK
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ==================== TRANG BÌA ====================
    add_p("TRƯỜNG ĐẠI HỌC SÀI GÒN", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20, space_after=2, font_size=14)
    add_p("KHOA CÔNG NGHỆ THÔNG TIN", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=40, font_size=14)
    
    add_p("BÁO CÁO BÀI TẬP TUẦN 05", bold=True, color=COLOR_PRIMARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=50, space_after=8, font_size=18)
    add_p("MÔN HỌC: HỌC SÂU (DEEP LEARNING)", bold=True, color=COLOR_SECONDARY, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=12, font_size=15)
    add_p("CHỦ ĐỀ: HỒI QUY TUYẾN TÍNH & HỒI QUY PHÂN LỚP\n(LINEAR REGRESSION & LOGISTIC REGRESSION)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=60, font_size=14)
    
    # Info Box Table
    info_table = doc.add_table(rows=4, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["Sinh viên thực hiện:", "Mã số sinh viên:", "Lớp:", "Giảng viên hướng dẫn:"]
    values = ["Bùi Đức Trường", "3123411319", "DCT123C4", "ThS. Đỗ Như Tài"]
    for i in range(4):
        c0 = info_table.rows[i].cells[0]
        c1 = info_table.rows[i].cells[1]
        c0.width = Inches(2.5)
        c1.width = Inches(3.5)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(4)
        p0.paragraph_format.space_after = Pt(4)
        r0 = p0.add_run(labels[i])
        r0.font.name = 'Times New Roman'
        r0.font.bold = True
        r0.font.size = Pt(13)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(4)
        p1.paragraph_format.space_after = Pt(4)
        r1 = p1.add_run(values[i])
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(13)
        
    add_p("Thành phố Hồ Chí Minh, tháng 10 năm 2026", bold=False, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=100, space_after=0, font_size=12)
    doc.add_page_break()

    # ==================== MỤC LỤC & TỔNG QUAN ====================
    add_h1("TỔNG QUAN NỘI DUNG BÁO CÁO")
    add_p("Báo cáo gồm đầy đủ các phần bài tập được giao theo yêu cầu của giảng viên:")
    add_p("• PHẦN I: Ví dụ mẫu trong slide bài giảng (Hồi quy tuyến tính dự đoán Cân nặng theo Chiều cao) giải chi tiết bằng cả 3 phương pháp: (1) Giải hệ phương trình đạo hàm, (2) Ma trận nghịch đảo giả Pseudo-Inverse, (3) Batch Gradient Descent.", bold=False)
    add_p("• PHẦN II: Bài tập trong file DL03_Regression.docx (10 câu về Hồi quy tuyến tính đơn, đa biến, Gradient Descent, Decision Boundary, Logistic Regression, Cross-Entropy Loss và Perceptron Learning).", bold=False)
    add_p("• PHẦN III: Bài tập trong file DL03_Regression_P2.docx (15 câu nâng cao về OLS, các chỉ số đánh giá MAE/MSE/RMSE, RelMSE, CV, Ma trận, Đa thức bậc 2, Logistic đa biến, R^2, và Regularization Ridge, LASSO, Elastic Net).", bold=False)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # ==================== PHẦN I: VÍ DỤ MẪU SLIDE ====================
    add_h1("PHẦN I: VÍ DỤ MẪU TRONG SLIDE BÀI GIẢNG")
    add_p("Đề bài: Cho tập dữ liệu mẫu gồm 5 người (N = 5) với thông tin Chiều cao (x - cm) và Cân nặng (y - kg):", bold=True)
    
    # Bảng dữ liệu slide
    h_slide = ["Điểm (i)", "Tên", "Chiều cao x (cm)", "Cân nặng y (kg)", "x^2", "x*y"]
    d_slide = [
        ["1", "A", "170", "65", "28,900", "11,050"],
        ["2", "B", "155", "50", "24,025", "7,750"],
        ["3", "C", "150", "45", "22,500", "6,750"],
        ["4", "D", "175", "70", "30,625", "12,250"],
        ["5", "E", "165", "55", "27,225", "9,075"],
        ["Tổng (Σ)", "N = 5", "Σx = 815", "Σy = 285", "Σx² = 133,275", "Σxy = 46,875"]
    ]
    w_slide = [Inches(0.8), Inches(0.8), Inches(1.3), Inches(1.3), Inches(1.1), Inches(1.1)]
    tbl_slide = doc.add_table(rows=1, cols=6)
    format_table(tbl_slide, w_slide, h_slide, d_slide)

    add_p("Mô hình hồi quy tuyến tính giả thiết: y_hat = a*x + b. Mục tiêu là xác định hai tham số a (hệ số dốc) và b (hệ số chặn / bias).")
    
    # Phương pháp 1
    add_h2("1. Phương pháp 1: Giải hệ phương trình đạo hàm (OLS / Residual)")
    add_p("Bước 1: Thiết lập hàm tổng sai số bình phương phần dư (Residual):")
    add_p("R(a, b) = Σ [y_hat_i - y_i]² = Σ (a*x_i + b - y_i)²", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("Khai triển cho 5 điểm dữ liệu:")
    add_p("R(a, b) = (170a + b - 65)² + (155a + b - 50)² + (150a + b - 45)² + (175a + b - 70)² + (165a + b - 55)²", italic=True)
    
    add_p("Bước 2: Tìm điều kiện cực trị bằng cách cho đạo hàm riêng triệt tiêu:")
    add_p("∂R/∂a = 2 * Σ x_i * (a*x_i + b - y_i) = 2 * [ a*Σ(x_i²) + b*Σx_i - Σ(x_i*y_i) ] = 0\n"
          "<=> a*Σ(x_i²) + b*Σx_i = Σ(x_i*y_i)", bold=True)
    add_p("∂R/∂b = 2 * Σ (a*x_i + b - y_i) = 2 * [ a*Σx_i + N*b - Σy_i ] = 0\n"
          "<=> a*Σx_i + N*b = Σy_i", bold=True)

    add_p("Bước 3: Thay số liệu tính được từ bảng:")
    add_p("• 133,275 a + 815 b = 46,875   (Phương trình 1)\n"
          "• 815 a + 5 b = 285            (Phương trình 2)", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_p("Bước 4: Giải hệ phương trình bậc nhất hai ẩn:")
    add_p("Từ phương trình (2), chia cả hai vế cho 5 ta được: 163 a + b = 57 => b = 57 - 163 a.")
    add_p("Thế b vào phương trình (1):")
    add_p("133,275 a + 815 * (57 - 163 a) = 46,875\n"
          "133,275 a + 46,455 - 132,845 a = 46,875\n"
          "430 a = 420 => a = 420 / 430 = 42 / 43 ≈ 0.976744", bold=True)
    add_p("Tính hệ số chặn b:")
    add_p("b = 57 - 163 * (42 / 43) = (2451 - 6846) / 43 = -4395 / 43 ≈ -102.209302", bold=True)
    add_callout("Kết luận Phương pháp 1", "Phương trình hồi quy thu được: y_hat = 0.9767 x - 102.2093")

    # Phương pháp 2
    add_h2("2. Phương pháp 2: Ma trận nghịch đảo giả (Pseudo-Inverse Moore-Penrose)")
    add_p("Biểu diễn hệ phương trình dưới dạng ma trận A*x = Y:")
    add_p("• Ma trận đặc trưng A (kích thước 5x2, bổ sung cột 1 biểu diễn bias):\n"
          "  A = [[170, 1], [155, 1], [150, 1], [175, 1], [165, 1]]\n"
          "• Vector trọng số x = [a, b]^T (kích thước 2x1)\n"
          "• Vector nhãn thực tế Y = [65, 50, 45, 70, 55]^T (kích thước 5x1)", italic=True)
    add_p("Vì A không vuông, ta dùng nghiệm chuẩn thông qua ma trận giả nghịch đảo:")
    add_p("x = (A^T * A)^(-1) * A^T * Y = A^+ * Y", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_p("Tính các thành phần ma trận:")
    add_p("• A^T * A = [[Σx², Σx], [Σx, N]] = [[133275, 815], [815, 5]]\n"
          "• det(A^T * A) = 133,275 * 5 - 815² = 666,375 - 664,225 = 2,150 ≠ 0\n"
          "• (A^T * A)^(-1) = (1 / 2150) * [[5, -815], [-815, 133275]]\n"
          "• A^T * Y = [Σxy, Σy]^T = [46875, 285]^T", bold=True)
    add_p("Thực hiện phép nhân ma trận tìm vector nghiệm x:")
    add_p("x = (1 / 2150) * [[5*46875 - 815*285], [-815*46875 + 133275*285]]\n"
          "x = (1 / 2150) * [[234375 - 232275], [-38203125 + 37983375]]\n"
          "x = (1 / 2150) * [[2100], [-219750]] = [[42/43], [-4395/43]] ≈ [[0.976744], [-102.209302]]", bold=True)
    add_callout("Kết luận Phương pháp 2", "Kết quả tìm được bằng Pseudo-Inverse trùng khớp hoàn toàn: a ≈ 0.9767, b ≈ -102.2093.")

    # Phương pháp 3
    add_h2("3. Phương pháp 3: Batch Gradient Descent (Hạ độ dốc)")
    add_p("Hàm mục tiêu Mean Squared Error (MSE): J(a, b) = (1 / N) * Σ (a*x_i + b - y_i)²")
    add_p("Công thức Gradient đạo hàm riêng:")
    add_p("∂J/∂a = (2 / N) * [ a*Σx² + b*Σx - Σxy ] = (2 / 5) * [ 133,275 a + 815 b - 46,875 ]\n"
          "∂J/∂b = (2 / N) * [ a*Σx + N*b - Σy ] = (2 / 5) * [ 815 a + 5 b - 285 ]", bold=True)
    add_p("Quy tắc cập nhật lặp: a^(t+1) = a^(t) - η * (∂J/∂a),   b^(t+1) = b^(t) - η * (∂J/∂b).")
    add_p("Do giá trị x lớn (150 ~ 175) nên thành phần Σx² rất lớn. Theo tài liệu slide và file Excel mẫu Regression_Demo.xlsx, ta chọn tốc độ học nhỏ: η = 0.00001 (10^-5) để tránh phân kỳ.", italic=True)

    add_p("Mô phỏng 2 bước cập nhật đầu tiên với khởi tạo a^(0) = 0, b^(0) = 0:")
    add_p("• Bước lặp 1 (t = 0 -> 1):\n"
          "  - ∂J/∂a |(0,0) = (2/5) * (-46,875) = -18,750\n"
          "  - ∂J/∂b |(0,0) = (2/5) * (-285) = -114\n"
          "  - a^(1) = 0 - 0.00001 * (-18,750) = 0.1875\n"
          "  - b^(1) = 0 - 0.00001 * (-114) = 0.00114", bold=True)
    add_p("• Bước lặp 2 (t = 1 -> 2):\n"
          "  - Sai số gradient: ∂J/∂a = -8,754.0034;   ∂J/∂b = -52.8727\n"
          "  - a^(2) = 0.1875 - 0.00001 * (-8,754.0034) = 0.27504\n"
          "  - b^(2) = 0.00114 - 0.00001 * (-52.8727) = 0.001669", bold=True)
    add_p("Khi số vòng lặp đủ lớn (khoảng 1000 - 5000 epochs như thể hiện trong file Excel đính kèm), giá trị hàm mất mát giảm dần về cực tiểu và cặp tham số hội tụ chính xác về a -> 0.9767, b -> -102.2093.")

    # Bảng so sánh 3 phương pháp
    add_h3("Bảng tổng kết so sánh 3 phương pháp:")
    h_cmp = ["Phương pháp", "Công thức cốt lõi", "Kết quả hệ số a", "Kết quả hệ số b", "Ưu điểm & Ứng dụng"]
    d_cmp = [
        ["1. Giải hệ PT đạo hàm", "Σx²*a + Σx*b = Σxy\nΣx*a + N*b = Σy", "≈ 0.9767", "≈ -102.2093", "Dễ giải tích, trực quan cho bài toán 1 chiều."],
        ["2. Pseudo-Inverse", "x = (A^T*A)^(-1) * A^T * Y", "≈ 0.9767", "≈ -102.2093", "Mã nguồn gọn gàng, tổng quát cho bài toán nhiều biến."],
        ["3. Gradient Descent", "θ^(t+1) = θ^(t) - η*∇J(θ)", "Hội tụ về 0.9767", "Hội tụ về -102.2093", "Phù hợp tập dữ liệu khổng lồ (Big Data) không thể nghịch đảo ma trận."]
    ]
    w_cmp = [Inches(1.3), Inches(1.8), Inches(0.9), Inches(1.0), Inches(1.5)]
    tbl_cmp = doc.add_table(rows=1, cols=5)
    format_table(tbl_cmp, w_cmp, h_cmp, d_cmp, align_center=False)
    doc.add_page_break()

    # ==================== PHẦN II: DL03_REGRESSION.DOCX ====================
    add_h1("PHẦN II: BÀI TẬP FILE DL03_REGRESSION.DOCX (10 CÂU)")
    
    # CÂU 1
    add_h2("CÂU 1. XÂY DỰNG PHƯƠNG TRÌNH HỒI QUY TUYẾN TÍNH ĐƠN (OLS)")
    add_p("Dữ liệu: x = [1, 2, 3, 4, 5], y = [3, 5, 7, 9, 11]. Mô hình: y = β0 + β1*x.")
    add_p("a) Tính giá trị trung bình:\n"
          "   x_bar = (1 + 2 + 3 + 4 + 5) / 5 = 3.0\n"
          "   y_bar = (3 + 5 + 7 + 9 + 11) / 5 = 7.0\n"
          "b) Tính SS_xy và SS_x:\n"
          "   • Độ lệch x: [-2, -1, 0, 1, 2];   Độ lệch y: [-4, -2, 0, 2, 4]\n"
          "   • SS_xy = Σ (x_i - x_bar)(y_i - y_bar) = (-2)(-4) + (-1)(-2) + 0 + (1)(2) + (2)(4) = 8 + 2 + 0 + 2 + 8 = 20\n"
          "   • SS_x = Σ (x_i - x_bar)² = (-2)² + (-1)² + 0² + 1² + 2² = 4 + 1 + 0 + 1 + 4 = 10\n"
          "c) Tính hệ số dốc: β1 = SS_xy / SS_x = 20 / 10 = 2.0\n"
          "d) Tính hệ số chặn: β0 = y_bar - β1 * x_bar = 7.0 - 2.0 * 3.0 = 1.0\n"
          "e) Phương trình hồi quy: y_hat = 2x + 1\n"
          "f) Dự đoán y khi x = 6: y_hat = 2(6) + 1 = 13", bold=False)

    # CÂU 2
    add_h2("CÂU 2. TÍNH MEAN SQUARED ERROR (MSE)")
    add_p("Cho mô hình y_hat = 2x + 1 và dữ liệu thực tế 4 mẫu: (1, 4), (2, 6), (3, 8), (4, 11).")
    add_p("a) Giá trị dự đoán từng mẫu:\n"
          "   y_hat_1 = 2(1) + 1 = 3;   y_hat_2 = 2(2) + 1 = 5;   y_hat_3 = 2(3) + 1 = 7;   y_hat_4 = 2(4) + 1 = 9\n"
          "b) Sai số phần dư e_i = y_i - y_hat_i:\n"
          "   e_1 = 4 - 3 = 1;   e_2 = 6 - 5 = 1;   e_3 = 8 - 7 = 1;   e_4 = 11 - 9 = 2\n"
          "c) Mean Squared Error (MSE):\n"
          "   MSE = (1 / n) * Σ e_i² = (1² + 1² + 1² + 2²) / 4 = (1 + 1 + 1 + 4) / 4 = 7 / 4 = 1.75\n"
          "d) Nhận xét độ phù hợp:\n"
          "   Mô hình có MSE = 1.75 khá nhỏ (RMSE ≈ 1.32), dự đoán gần với thực tế. Tuy nhiên toàn bộ e_i > 0 cho thấy mô hình luôn dự đoán hơi thấp hơn thực tế, có thể nâng bias để khớp tốt hơn.", bold=False)

    # CÂU 3
    add_h2("CÂU 3. GRADIENT DESCENT CHO LINEAR REGRESSION")
    add_p("Dữ liệu: (1, 3), (2, 5), (4, 9). Mô hình: y = wx + b. Khởi tạo: w = 0, b = 0, η = 0.1.")
    add_p("Hàm mất mát: J(w, b) = (1 / 2m) * Σ (y_hat_i - y_i)² với m = 3.")
    add_p("a) Dự đoán ban đầu: y_hat_1 = 0, y_hat_2 = 0, y_hat_3 = 0\n"
          "b) Tính gradient:\n"
          "   • ∂J/∂w = (1/m) * Σ (y_hat_i - y_i) * x_i = (1/3) * [(-3)(1) + (-5)(2) + (-9)(4)] = -49 / 3 ≈ -16.3333\n"
          "   • ∂J/∂b = (1/m) * Σ (y_hat_i - y_i) = (1/3) * [(-3) + (-5) + (-9)] = -17 / 3 ≈ -5.6667\n"
          "c) Cập nhật 1 bước Gradient Descent:\n"
          "   • w_new = w - η * (∂J/∂w) = 0 - 0.1 * (-49/3) = 4.9 / 3 ≈ 1.6333\n"
          "   • b_new = b - η * (∂J/∂b) = 0 - 0.1 * (-17/3) = 1.7 / 3 ≈ 0.5667\n"
          "d) Phương trình hồi quy mới: y_hat = 1.6333 x + 0.5667", bold=False)

    # CÂU 4
    add_h2("CÂU 4. HỒI QUY TUYẾN TÍNH ĐA BIẾN DẠNG MA TRẬN")
    add_p("Dữ liệu: (x1, x2, y) = (1, 2, 5), (2, 1, 6), (3, 2, 9), (4, 3, 12). Mô hình: y = β0 + β1*x1 + β2*x2.")
    add_p("a) Ma trận X và vector Y:\n"
          "   X = [[1, 1, 2], [1, 2, 1], [1, 3, 2], [1, 4, 3]] (kích thước 4x3)\n"
          "   Y = [5, 6, 9, 12]^T (kích thước 4x1)\n"
          "b) Công thức nghiệm chuẩn: β = (X^T * X)^(-1) * X^T * Y\n"
          "c) Tính X^T * X: [[4, 10, 8], [10, 30, 22], [8, 22, 18]]\n"
          "d) Tính X^T * Y: [32, 92, 70]^T\n"
          "e) Giải hệ tìm vector tham số β:\n"
          "   Hệ: 4β0 + 10β1 + 8β2 = 32; 10β0 + 30β1 + 22β2 = 92; 8β0 + 22β1 + 18β2 = 70\n"
          "   => β = [1.0, 2.0, 1.0]^T (tức β0 = 1, β1 = 2, β2 = 1)\n"
          "   Phương trình hồi quy: y_hat = 1 + 2 x1 + x2\n"
          "f) Dự đoán y khi x1 = 5, x2 = 2: y_hat = 1 + 2(5) + 1(2) = 13", bold=False)

    # CÂU 5
    add_h2("CÂU 5. BATCH GRADIENT DESCENT CHO 5 ĐIỂM DỮ LIỆU")
    add_p("Dữ liệu: (1, 2), (2, 3), (3, 5), (4, 4), (5, 6). Mô hình: y = wx + b. w0 = 0, b0 = 0, η = 0.1, m = 5.")
    add_p("a) Biểu thức đạo hàm và hàm mất mát:\n"
          "   • J(w, b) = (1 / 2m) * Σ (w*x_i + b - y_i)² = (1/10) * Σ (w*x_i + b - y_i)²\n"
          "   • ∂J/∂w = (1/m) * Σ (w*x_i + b - y_i) * x_i\n"
          "   • ∂J/∂b = (1/m) * Σ (w*x_i + b - y_i)\n"
          "b) Thực hiện 1 bước cập nhật (tại w0 = 0, b0 = 0):\n"
          "   • Sai số: y_hat_i - y_i = -y_i\n"
          "   • Σ (-y_i * x_i) = -(2*1 + 3*2 + 5*3 + 4*4 + 6*5) = -69 => ∂J/∂w = -69 / 5 = -13.8\n"
          "   • Σ (-y_i) = -(2 + 3 + 5 + 4 + 6) = -20 => ∂J/∂b = -20 / 5 = -4.0\n"
          "   • Cập nhật: w1 = 0 - 0.1 * (-13.8) = 1.38;   b1 = 0 - 0.1 * (-4.0) = 0.40\n"
          "c) Dự đoán khi x = 6: y_hat = 1.38 * 6 + 0.40 = 8.28 + 0.40 = 8.68", bold=False)

    # CÂU 6
    add_h2("CÂU 6. PHÂN LỚP BẰNG ĐƯỜNG QUYẾT ĐỊNH TUYẾN TÍNH")
    add_p("Cho w = (2, -1), b = 1. Quy tắc: w^T * x + b >= 0 => Class 1; < 0 => Class 0.")
    add_p("Các điểm dữ liệu: A(1, 1), B(2, 1), C(1, 3), D(4, 2).")
    add_p("a) & b) Tính giá trị z = 2*x1 - x2 + 1 và xác định lớp:\n"
          "   • Điểm A(1, 1): z_A = 2(1) - 1 + 1 = 2 >= 0 => Class 1\n"
          "   • Điểm B(2, 1): z_B = 2(2) - 1 + 1 = 4 >= 0 => Class 1\n"
          "   • Điểm C(1, 3): z_C = 2(1) - 3 + 1 = 0 >= 0 => Class 1 (nằm trên đường biên)\n"
          "   • Điểm D(4, 2): z_D = 2(4) - 2 + 1 = 7 >= 0 => Class 1\n"
          "c) Phương trình đường quyết định: 2*x1 - x2 + 1 = 0 <=> x2 = 2*x1 + 1\n"
          "d) Đồ thị: Đường thẳng x2 = 2*x1 + 1 chia mặt phẳng thành hai nửa. Cả 4 điểm A, B, C, D đều nằm về phía nửa mặt phẳng x2 <= 2*x1 + 1 (Class 1).", bold=False)

    # CÂU 7
    add_h2("CÂU 7. LOGISTIC REGRESSION – TÍNH XÁC SUẤT PHÂN LỚP")
    add_p("Mô hình: P(y=1|x) = σ(z) = 1 / (1 + e^-z) với z = w1*x1 + w2*x2 + b.")
    add_p("Cho w1 = 0.5, w2 = -1, b = 0.2 và điểm x = (2, 1).")
    add_p("a) Tính z: z = 0.5(2) + (-1)(1) + 0.2 = 1.0 - 1.0 + 0.2 = 0.2\n"
          "b) Tính xác suất: P(y=1|x) = 1 / (1 + e^-0.2) = 1 / (1 + 0.81873) ≈ 0.5498 (54.98%)\n"
          "c) Phân loại với ngưỡng 0.5: Do 0.5498 >= 0.5 => Mẫu thuộc Lớp 1.\n"
          "d) Phân loại với ngưỡng 0.8: Do 0.5498 < 0.8 => Mẫu thuộc Lớp 0.", bold=False)

    # CÂU 8
    add_h2("CÂU 8. TÍNH CROSS-ENTROPY LOSS")
    add_p("Bảng dữ liệu: Mẫu 1 (y=1, z=0.9); Mẫu 2 (y=0, z=0.2); Mẫu 3 (y=1, z=0.7).")
    add_p("Hàm mất mát: J = - Σ [ y_i * ln(z_i) + (1 - y_i) * ln(1 - z_i) ].")
    add_p("a) Loss từng mẫu (tính theo log tự nhiên ln):\n"
          "   • L1 = -ln(0.9) ≈ 0.1054\n"
          "   • L2 = -ln(1 - 0.2) = -ln(0.8) ≈ 0.2231\n"
          "   • L3 = -ln(0.7) ≈ 0.3567\n"
          "   (Nếu tính theo log10: L1 ≈ 0.0458, L2 ≈ 0.0969, L3 ≈ 0.1549)\n"
          "b) Tổng loss: J_total = 0.1054 + 0.2231 + 0.3567 = 0.6852 (hoặc 0.2976 theo log10)\n"
          "c) Average loss: J_avg = 0.6852 / 3 ≈ 0.2284 (hoặc 0.0992 theo log10)\n"
          "d) Nhận xét: Giá trị loss trung bình rất thấp, xác suất dự đoán đồng pha hoàn toàn với nhãn thực tế. Mô hình có chất lượng phân lớp rất tốt.", bold=False)

    # CÂU 9
    add_h2("CÂU 9. PERCEPTRON LEARNING")
    add_p("Dữ liệu: x1(1, 1, t=1), x2(2, 1, t=1), x3(1, 3, t=-1). w0 = (0, 0), b0 = 0, ρ = 1.")
    add_p("Quy tắc: Cập nhật khi t_k * (w^T * x_k + b) <= 0: w <- w + ρ*t_k*x_k; b <- b + ρ*t_k.")
    add_p("Quá trình duyệt 1 epoch:\n"
          "• Mẫu 1 (1, 1, t=1): f = 0*1 + 0*1 + 0 = 0 => t*f = 0 <= 0 (Sai!).\n"
          "  Cập nhật: w^(1) = (0, 0) + 1*1*(1, 1) = (1, 1);   b^(1) = 0 + 1*1 = 1.\n"
          "• Mẫu 2 (2, 1, t=1): f = 1*2 + 1*1 + 1 = 4 > 0 => t*f = 4 > 0 (Đúng!). Giữ nguyên w^(2) = (1, 1), b^(2) = 1.\n"
          "• Mẫu 3 (1, 3, t=-1): f = 1*1 + 1*3 + 1 = 5 > 0 => t*f = (-1)*5 = -5 <= 0 (Sai!).\n"
          "  Cập nhật: w^(3) = (1, 1) + 1*(-1)*(1, 3) = (0, -2);   b^(3) = 1 + 1*(-1) = 0.\n"
          "Kết quả sau 1 epoch: w = (0, -2), b = 0.\n"
          "Phương trình đường quyết định: 0*x1 - 2*x2 + 0 = 0 <=> x2 = 0 (trục hoành).", bold=False)

    # CÂU 10
    add_h2("CÂU 10. HỒI QUY PHÂN LỚP VỚI HÀM CHI PHÍ PERCEPTRON")
    add_p("Dữ liệu 4 mẫu: x1(1,1, +1), x2(2,1, +1), x3(1,2, -1), x4(2,2, -1). Tham số ban đầu: w0 = (1, 1)^T, b0 = 0.")
    add_p("a) Xác định các mẫu bị phân lớp sai:\n"
          "   • Mẫu 1: 1(1) + 1(1) + 0 = 2 >= 0 => sign = +1 (Đúng)\n"
          "   • Mẫu 2: 1(2) + 1(1) + 0 = 3 >= 0 => sign = +1 (Đúng)\n"
          "   • Mẫu 3: 1(1) + 1(2) + 0 = 3 >= 0 => sign = +1 != -1 (Sai)\n"
          "   • Mẫu 4: 1(2) + 1(2) + 0 = 4 >= 0 => sign = +1 != -1 (Sai)\n"
          "   => Tập mẫu bị phân lớp sai: Y = {x3, x4}.\n"
          "b) Đạo hàm hàm chi phí Perceptron: ∂J/∂w = - Σ_(k∈Y) t_k * x_k;   ∂J/∂b = - Σ_(k∈Y) t_k\n"
          "c) Một bước cập nhật theo Batch Perceptron (với η = 1):\n"
          "   • Σ t_k * x_k = (-1)*(1, 2) + (-1)*(2, 2) = (-3, -4) => ∂J/∂w = -(-3, -4) = (3, 4)\n"
          "   • Σ t_k = (-1) + (-1) = -2 => ∂J/∂b = -(-2) = 2\n"
          "   • Cập nhật: w1 = (1, 1) - 1*(3, 4) = (-2, -3);   b1 = 0 - 1*2 = -2.", bold=False)
    doc.add_page_break()

    # ==================== PHẦN III: DL03_REGRESSION_P2.DOCX ====================
    add_h1("PHẦN III: BÀI TẬP FILE DL03_REGRESSION_P2.DOCX (15 CÂU)")

    # P2 - CÂU 1
    add_h2("CÂU 1. TÍNH HỆ SỐ HỒI QUY TUYẾN TÍNH ĐƠN BẰNG OLS")
    add_p("Dữ liệu tuần x = [1, 2, 3, 4, 5], doanh số y = [1.2, 1.8, 2.6, 3.2, 3.8]. Mô hình: y = a0 + a1*x.")
    add_p("a) Tính các đại lượng thống kê trung bình:\n"
          "   x_bar = 3.0;   y_bar = 2.52;   (x²)_bar = 11.0;   (xy)_bar = 8.88\n"
          "b) Tính hệ số dốc: a1 = ((xy)_bar - x_bar*y_bar) / ((x²)_bar - (x_bar)²) = (8.88 - 3.0 * 2.52) / (11.0 - 9.0) = 1.32 / 2.0 = 0.66\n"
          "c) Tính hệ số chặn: a0 = y_bar - a1 * x_bar = 2.52 - 0.66 * 3.0 = 2.52 - 1.98 = 0.54\n"
          "d) Phương trình hồi quy: y_hat = 0.54 + 0.66 x\n"
          "e) Dự đoán doanh số tuần 7 và tuần 9:\n"
          "   • Tuần 7: y_hat(7) = 0.54 + 0.66(7) = 5.16\n"
          "   • Tuần 9: y_hat(9) = 0.54 + 0.66(9) = 6.48", bold=False)

    # P2 - CÂU 2
    add_h2("CÂU 2. TÍNH SAI SỐ DỰ ĐOÁN CỦA MÔ HÌNH HỒI QUY")
    add_p("Cho mô hình y_hat = 0.54 + 0.66x và dữ liệu y = [1.2, 1.8, 2.6, 3.2, 3.8].")
    add_p("a) Giá trị dự đoán y_hat từng mẫu:\n"
          "   y_hat = [0.54 + 0.66(1), ..., 0.54 + 0.66(5)] = [1.20, 1.86, 2.52, 3.18, 3.84]\n"
          "b) Residual e_i = y_i - y_hat_i:\n"
          "   e = [1.2 - 1.20, 1.8 - 1.86, 2.6 - 2.52, 3.2 - 3.18, 3.8 - 3.84] = [0.00, -0.06, +0.08, +0.02, -0.04]\n"
          "c) Tổng bình phương sai số SSE = Σ e_i²:\n"
          "   SSE = 0² + (-0.06)² + 0.08² + 0.02² + (-0.04)² = 0 + 0.0036 + 0.0064 + 0.0004 + 0.0016 = 0.0120\n"
          "d) Nhận xét: SSE = 0.0120 cực kỳ nhỏ, mô hình khớp xuất sắc với dữ liệu thực tế.", bold=False)

    # P2 - CÂU 3
    add_h2("CÂU 3. TÍNH MAE, MSE, RMSE")
    add_p("Dữ liệu: y_thực = [80, 75, 90, 100], y_dự đoán = [75, 85, 88, 95] (n = 4).")
    add_p("a) Sai số tuyệt đối |y_i - y_hat_i|: [|80 - 75|, |75 - 85|, |90 - 88|, |100 - 95|] = [5, 10, 2, 5]\n"
          "b) MAE = (1/n) * Σ |e_i| = (5 + 10 + 2 + 5) / 4 = 22 / 4 = 5.5\n"
          "c) MSE = (1/n) * Σ e_i² = (5² + 10² + 2² + 5²) / 4 = (25 + 100 + 4 + 25) / 4 = 154 / 4 = 38.5\n"
          "d) RMSE = sqrt(MSE) = sqrt(38.5) ≈ 6.2048\n"
          "e) Nhận xét: Sai số trung bình khoảng 5.5 đến 6.2 đơn vị so với thang đo doanh số ~80-100 (tỷ lệ sai số ~6%), đây là mức sai số nhỏ và mô hình dự đoán khá tin cậy.", bold=False)

    # P2 - CÂU 4
    add_h2("CÂU 4. TÍNH RELATIVE MSE VÀ HỆ SỐ BIẾN THIÊN (CV)")
    add_p("Tập huấn luyện: [80, 90, 100, 110, 120]. Tập kiểm tra: y_thực = [80, 75], y_dự đoán = [75, 85].")
    add_p("a) y_bar tập huấn luyện: y_bar = (80 + 90 + 100 + 110 + 120) / 5 = 100.0\n"
          "b) Relative MSE:\n"
          "   • Tử số = Σ (y_test - y_pred)² = (80 - 75)² + (75 - 85)² = 25 + 100 = 125\n"
          "   • Mẫu số = Σ (y_test - y_bar_train)² = (80 - 100)² + (75 - 100)² = (-20)² + (-25)² = 400 + 625 = 1025\n"
          "   => RelMSE = 125 / 1025 ≈ 0.1220 (12.2%)\n"
          "c) Hệ số biến thiên CV = RMSE / y_bar_train:\n"
          "   RMSE_test = sqrt(125 / 2) = sqrt(62.5) ≈ 7.9057 => CV = 7.9057 / 100.0 ≈ 0.0791 (7.91%)\n"
          "d) Nhận xét: RelMSE = 12.2% << 1 và CV = 7.91% < 10% chứng tỏ mô hình tổng quát hóa rất tốt trên tập kiểm tra.", bold=False)

    # P2 - CÂU 5
    add_h2("CÂU 5. HỒI QUY TUYẾN TÍNH DẠNG MA TRẬN")
    add_p("Dữ liệu: x = [1, 2, 3, 4], y = [1, 3, 4, 8]. Mô hình: y = a0 + a1*x.")
    add_p("a) Ma trận X: [[1, 1], [1, 2], [1, 3], [1, 4]]\n"
          "b) Vector Y: [1, 3, 4, 8]^T\n"
          "c) X^T * X = [[4, 10], [10, 30]]\n"
          "d) (X^T * X)^(-1): det = 4*30 - 10*10 = 20 => (X^T * X)^(-1) = [[1.5, -0.5], [-0.5, 0.2]]\n"
          "e) X^T * Y = [16, 51]^T => a = (X^T * X)^(-1) * X^T * Y = [[1.5*16 - 0.5*51], [-0.5*16 + 0.2*51]] = [-1.5, 2.2]^T\n"
          "f) Phương trình hồi quy: y_hat = -1.5 + 2.2 x", bold=False)

    # P2 - CÂU 6
    add_h2("CÂU 6. HỒI QUY TUYẾN TÍNH ĐA BIẾN DẠNG MA TRẬN")
    add_p("Dữ liệu: (x1, x2, y) = (1, 4, 1), (2, 5, 6), (3, 8, 8), (4, 2, 12). Mô hình: y = a0 + a1*x1 + a2*x2.")
    add_p("a) Ma trận X và vector Y:\n"
          "   X = [[1, 1, 4], [1, 2, 5], [1, 3, 8], [1, 4, 2]];   Y = [1, 6, 8, 12]^T\n"
          "b) Hệ số a = (X^T * X)^(-1) * X^T * Y:\n"
          "   • X^T * X = [[4, 10, 19], [10, 30, 46], [19, 46, 109]]\n"
          "   • X^T * Y = [27, 85, 122]^T\n"
          "   • Giải ra hệ số: a0 ≈ -1.6995, a1 ≈ 3.4836, a2 ≈ -0.0546\n"
          "c) Mô hình hồi quy: y_hat = -1.6995 + 3.4836 x1 - 0.0546 x2\n"
          "d) Dự đoán khi x1 = 5, x2 = 6: y_hat = -1.6995 + 3.4836(5) - 0.0546(6) ≈ 15.3907", bold=False)

    # P2 - CÂU 7
    add_h2("CÂU 7. HỒI QUY ĐA THỨC BẬC 2")
    add_p("Dữ liệu: (1, 1), (2, 4), (3, 9), (4, 15). Mô hình: y = a0 + a1*x + a2*x².")
    add_p("a) & b) Bảng tính các tổng (n = 4):\n"
          "   • Σx = 10;   Σy = 29;   Σx² = 30;   Σx³ = 100;   Σx⁴ = 354;   Σxy = 96;   Σx²y = 338\n"
          "c) Hệ phương trình chuẩn:\n"
          "   [[4, 10, 30], [10, 30, 100], [30, 100, 354]] * [a0, a1, a2]^T = [29, 96, 338]^T\n"
          "d) Giải hệ phương trình:\n"
          "   a0 = -0.75,   a1 = 0.95,   a2 = 0.75\n"
          "e) Phương trình hồi quy đa thức: y_hat = -0.75 + 0.95 x + 0.75 x²\n"
          "f) Dự đoán y khi x = 5: y_hat = -0.75 + 0.95(5) + 0.75(25) = -0.75 + 4.75 + 18.75 = 22.75", bold=False)

    # P2 - CÂU 8
    add_h2("CÂU 8. LOGISTIC REGRESSION – TÍNH XÁC SUẤT VÀ PHÂN LỚP")
    add_p("Mô hình: p(x) = 1 / (1 + exp(-(-4 + 0.08*x))). x là điểm thi đầu vào.")
    add_p("a) Khi x = 40: z = -4 + 0.08(40) = -0.80 => p = 1 / (1 + e^0.80) ≈ 0.3100 (31.00%)\n"
          "b) Khi x = 60: z = -4 + 0.08(60) = +0.80 => p = 1 / (1 + e^-0.80) ≈ 0.6900 (69.00%)\n"
          "c) Khi x = 80: z = -4 + 0.08(80) = +2.40 => p = 1 / (1 + e^-2.40) ≈ 0.9168 (91.68%)\n"
          "d) Phân lớp với ngưỡng 0.5:\n"
          "   • x = 40: p = 0.3100 < 0.5 => Rớt\n"
          "   • x = 60: p = 0.6900 >= 0.5 => Đậu\n"
          "   • x = 80: p = 0.9168 >= 0.5 => Đậu", bold=False)

    # P2 - CÂU 9
    add_h2("CÂU 9. TÍNH ODDS VÀ LOG-ODDS")
    add_p("Xác suất spam: Email A (p = 0.2), Email B (p = 0.5), Email C (p = 0.8).")
    add_p("a) Tính odds = p / (1 - p):\n"
          "   • Email A: odds = 0.2 / 0.8 = 0.2500\n"
          "   • Email B: odds = 0.5 / 0.5 = 1.0000\n"
          "   • Email C: odds = 0.8 / 0.2 = 4.0000\n"
          "b) Tính log-odds (logit = ln(odds)):\n"
          "   • Email A: logit = ln(0.25) ≈ -1.3863\n"
          "   • Email B: logit = ln(1.00) = 0.0000\n"
          "   • Email C: logit = ln(4.00) ≈ +1.3863\n"
          "c) Nhận xét: Khi xác suất p tăng từ 0 đến 1, odds tăng đơn điệu từ 0 đến +vô cùng và log-odds tăng đơn điệu từ -vô cùng đến +vô cùng. Log-odds giúp ánh xạ đoạn xác suất [0, 1] sang toàn bộ trục số thực (-∞, +∞).", bold=False)

    # P2 - CÂU 10
    add_h2("CÂU 10. LOGISTIC REGRESSION VỚI HAI BIẾN ĐẦU VÀO")
    add_p("Mô hình: z = -3 + 0.04*x1 + 0.06*x2, p = 1 / (1 + e^-z).")
    add_p("Dữ liệu: SV A(60, 50), SV B(70, 80), SV C(40, 45), SV D(90, 85).")
    add_p("a) & b) Tính z và xác suất trúng tuyển p:\n"
          "   • SV A: z = -3 + 0.04(60) + 0.06(50) = 2.40 => p ≈ 0.9168 (91.68%)\n"
          "   • SV B: z = -3 + 0.04(70) + 0.06(80) = 4.60 => p ≈ 0.9900 (99.00%)\n"
          "   • SV C: z = -3 + 0.04(40) + 0.06(45) = 1.30 => p ≈ 0.7858 (78.58%)\n"
          "   • SV D: z = -3 + 0.04(90) + 0.06(85) = 5.70 => p ≈ 0.9967 (99.67%)\n"
          "c) Phân loại với ngưỡng 0.5: Cả 4 sinh viên A, B, C, D đều có p >= 0.5 nên đều Trúng tuyển.\n"
          "d) Tuyển sinh với ngưỡng 0.8: Chỉ có các sinh viên có p >= 0.8 được nhận gồm: Sinh viên A, Sinh viên B, Sinh viên D. Sinh viên C (78.58% < 80%) không được nhận.", bold=False)

    # P2 - CÂU 11
    add_h2("CÂU 11. TÍNH HỆ SỐ XÁC ĐỊNH R²")
    add_p("Dữ liệu: y_thực = [10, 15, 20, 25, 30], y_dự đoán = [12, 14, 19, 27, 28] (n = 5).")
    add_p("a) Tính y_bar: y_bar = (10 + 15 + 20 + 25 + 30) / 5 = 20.0\n"
          "b) Tổng biến thiên toàn phần SST = Σ (y_i - y_bar)²:\n"
          "   SST = (-10)² + (-5)² + 0² + 5² + 10² = 100 + 25 + 0 + 25 + 100 = 250.0\n"
          "c) Tổng sai số chưa giải thích SSE = Σ (y_i - y_hat_i)²:\n"
          "   e = [10-12, 15-14, 20-19, 25-27, 30-28] = [-2, 1, 1, -2, 2]\n"
          "   SSE = (-2)² + 1² + 1² + (-2)² + 2² = 4 + 1 + 1 + 4 + 4 = 14.0\n"
          "d) Hệ số xác định R²: R² = 1 - (SSE / SST) = 1 - (14 / 250) = 1 - 0.056 = 0.9440 (94.4%)\n"
          "e) Ý nghĩa: Giá trị R² = 0.944 cho biết mô hình hồi quy giải thích được 94.4% sự biến thiên của biến phụ thuộc y, chỉ còn 5.6% sai số ngẫu nhiên chưa giải thích được.", bold=False)

    # P2 - CÂU 12
    add_h2("CÂU 12. RIDGE REGRESSION – TÍNH HÀM MẤT MÁT CÓ PHẠT L2")
    add_p("Cho SSE = 40, hệ số a1 = 3, λ = 2. Loss_Ridge = SSE + λ * a1².")
    add_p("a) Penalty Ridge: Penalty = λ * a1² = 2 * (3²) = 2 * 9 = 18\n"
          "b) Tổng loss: Loss_Ridge = 40 + 18 = 58\n"
          "c) Nếu λ tăng lên 5: Penalty = 5 * 9 = 45 => Loss_Ridge = 40 + 45 = 85 (loss tăng thêm 27 đơn vị)\n"
          "d) Giảm overfitting: Ridge phạt các trọng số có độ lớn cao bằng cách co (shrinkage) các hệ số về gần 0, làm mượt đường dự đoán và giảm phương sai (variance) của mô hình.", bold=False)

    # P2 - CÂU 13
    add_h2("CÂU 13. LASSO REGRESSION – TÍNH HÀM MẤT MÁT CÓ PHẠT L1")
    add_p("Cho SSE = 40, hệ số a1 = 3, a2 = -2, a3 = 0.5, λ = 2. Loss_LASSO = SSE + λ * Σ|ai|.")
    add_p("a) Tổng độ lớn tuyệt đối: Σ|ai| = |3| + |-2| + |0.5| = 3 + 2 + 0.5 = 5.5\n"
          "b) Penalty LASSO: Penalty = λ * Σ|ai| = 2 * 5.5 = 11.0\n"
          "c) Tổng loss: Loss_LASSO = 40 + 11.0 = 51.0\n"
          "d) Chọn lọc đặc trưng (Feature Selection): Chuẩn L1 có tính chất tạo nghiệm thưa (sparsity), có khả năng triệt tiêu hoàn toàn một số hệ số về đúng bằng 0, giúp loại bỏ các đặc trưng không quan trọng.", bold=False)

    # P2 - CÂU 14
    add_h2("CÂU 14. ELASTIC NET – TÍNH LOSS KẾT HỢP RIDGE VÀ LASSO")
    add_p("Cho SSE = 50, a1 = 2, a2 = -1, a3 = 3, λ1 = 1, λ2 = 0.5. Loss = SSE + λ1*Σ|ai| + λ2*Σai².")
    add_p("a) Thành phần LASSO: λ1 * Σ|ai| = 1 * (|2| + |-1| + |3|) = 1 * 6 = 6.0\n"
          "b) Thành phần Ridge: λ2 * Σai² = 0.5 * (2² + (-1)² + 3²) = 0.5 * (4 + 1 + 9) = 0.5 * 14 = 7.0\n"
          "c) Tổng Elastic Net loss: Loss = 50 + 6.0 + 7.0 = 63.0\n"
          "d) Khi nào nên dùng Elastic Net: Khi tập dữ liệu có nhiều biến độc lập tương quan cao (multicollinearity) hoặc số lượng thuộc tính p lớn hơn nhiều số mẫu n (p >> n), Elastic Net kết hợp ưu điểm chọn biến của LASSO và tính ổn định nhóm biến của Ridge.", bold=False)

    # P2 - CÂU 15
    add_h2("CÂU 15. BÀI TẬP TỔNG HỢP (QUẢNG CÁO & DOANH SỐ)")
    add_p("Dữ liệu 5 tuần: Chi phí quảng cáo x = [2, 3, 5, 7, 9], Doanh số y = [20, 25, 35, 45, 55].")
    add_p("a) Xây dựng mô hình OLS: y = a0 + a1*x\n"
          "   • x_bar = 5.2;   y_bar = 36.0\n"
          "   • SS_xy = Σ(x - x_bar)(y - y_bar) = 170.0;   SS_x = Σ(x - x_bar)² = 34.0\n"
          "   • a1 = 170.0 / 34.0 = 5.0\n"
          "   • a0 = 36.0 - 5.0 * 5.2 = 36.0 - 26.0 = 10.0\n"
          "   => Phương trình: y_hat = 10 + 5x\n"
          "b) Dự đoán doanh số khi x = 10: y_hat(10) = 10 + 5(10) = 60\n"
          "c) Tính MAE, MSE, RMSE:\n"
          "   Thay các giá trị x vào mô hình: y_hat = [10+5*2, 10+5*3, 10+5*5, 10+5*7, 10+5*9] = [20, 25, 35, 45, 55] = y_thực tế!\n"
          "   => Mọi sai số e_i = 0 => MAE = 0.0, MSE = 0.0, RMSE = 0.0\n"
          "d) Tính R²: Do SSE = 0 => R² = 1 - (0 / SST) = 1.0 (100%)\n"
          "e) Nhận xét: Mối quan hệ giữa chi phí quảng cáo và doanh số là tuyến tính hoàn hảo tuyệt đối (cứ mỗi triệu đồng chi thêm cho quảng cáo thì doanh số tăng đều đặn chính xác 5 triệu đồng).", bold=False)

    # Save document
    output_path = r"D:\UNIVERSITY\HocSau\SGU26K2_HocSau_S5_DCT123C4_0709\Tuan05\BT04_Regression_BuiDucTruong_3123411319_DapAn.docx"
    doc.save(output_path)
    print(f"File created successfully at: {output_path}")

if __name__ == "__main__":
    create_document()
