import io
import json
import re
import time
import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from deep_translator import GoogleTranslator
from PIL import Image
import numpy as np

# Thử import các thư viện OCR và đọc file bổ trợ
try:
    import easyocr
    # Khởi tạo sẵn reader cho tiếng Trung và tiếng Việt
    @st.cache_resource
    def get_ocr_reader():
        return easyocr.Reader(['ch_sim', 'vi', 'en'], gpu=False)
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

try:
    import fitz  # PyMuPDF để xử lý PDF sang ảnh
    PDF_IMAGE_AVAILABLE = True
except ImportError:
    PDF_IMAGE_AVAILABLE = False


# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Hai Chiều",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Dùng Thư Viện Chuyên Dụng)")
st.caption("Hỗ trợ chọn chế độ Trung ➔ Việt hoặc Việt ➔ Trung | Giữ nguyên 100% format Excel gốc hoặc chuyển từ Ảnh/PDF.")

# ============================================================
# 1. CẤU HÌNH BỘ LỌC HƯỚNG DỊCH & TẢI FILE
# ============================================================
col1, col2 = st.columns([1.2, 1.8])

with col1:
    translation_mode = st.radio(
        "Chế độ dịch:",
        options=["Trung ➔ Việt", "Việt ➔ Trung"],
        horizontal=True
    )

with col2:
    uploaded_file = st.file_uploader(
        "Tải lên Ảnh, PDF hoặc File Excel:",
        type=["png", "jpg", "jpeg", "pdf", "xlsx"]
    )

# Hàm kiểm tra chuỗi
def has_chinese(text):
    return bool(re.search(r'[\u4e00-\u9fff]', str(text))) if text else False

def has_vietnamese(text):
    if not isinstance(text, str):
        return False
    vietnamese_pattern = r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]'
    return bool(re.search(vietnamese_pattern, text, re.IGNORECASE))

# ============================================================
# HÀM DỊCH THUẬT DÙNG THƯ VIỆN CHUYÊN DỤNG (DEEP-TRANSLATOR)
# ============================================================
def translate_text(text, mode):
    if not text or not str(text).strip():
        return text
    
    # Xác định ngôn ngữ nguồn và đích cho deep-translator
    # 'zh-CN' cho tiếng Trung, 'vi' cho tiếng Việt
    source_lang = 'zh-CN' if mode == "Trung ➔ Việt" else 'vi'
    target_lang = 'vi' if mode == "Trung ➔ Việt" else 'zh-CN'
    
    try:
        translator = GoogleTranslator(source=source_lang, target=target_lang)
        result = translator.translate(str(text))
        return result if result else text
    except Exception as e:
        # Fallback thử lại nếu lỗi mạng tạm thời
        try:
            time.sleep(1)
            translator = GoogleTranslator(source=source_lang, target=target_lang)
            return translator.translate(str(text))
        except:
            return text

# Hàm dựng file Excel từ dữ liệu trích xuất (khi scan Ảnh/PDF)
def build_excel_from_json(data, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_name = "Microsoft YaHei"
    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    t_src = data.get("title_src", "")
    t_tgt = data.get("title_tgt", "")
    dt_str = data.get("date_str", "")
    rows = data.get("rows", [])

    top_title = t_src if mode == "Trung ➔ Việt" else t_tgt
    bot_title = t_tgt if mode == "Trung ➔ Việt" else t_src

    full_title = f"{dt_str} {top_title}\n{bot_title} ngày {dt_str}".strip()
    ws.merge_cells("A1:F1")
    ws["A1"] = full_title
    ws["A1"].font = Font(name=font_name, size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    headers = [("STT", "STT"), ("部门", "Bộ phận"), ("开几台机", "Số máy mở"), ("正式工", "Chính thức"), ("临时工", "Thời vụ"), ("备注", "Ghi chú")] if mode == "Trung ➔ Việt" else [("STT", "STT"), ("Bộ phận", "部门"), ("Số máy mở", "开几台机"), ("Chính thức", "正式工"), ("Thời vụ", "临时工"), ("Ghi chú", "备注")]

    for col_idx, (top_h, bot_h) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = f"{top_h}\n{bot_h}" if top_h != bot_h else top_h
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border
    ws.row_dimensions[2].height = 38

    current_row = 3
    total_workers = 0

    for row in rows:
        stt = row.get("stt", "")
        d_src = str(row.get("dept_src", "")) if row.get("dept_src") else ""
        d_tgt = str(row.get("dept_tgt", "")) if row.get("dept_tgt") else ""
        mac = row.get("machines", "") or ""
        fml = row.get("formal", "") or ""
        tmp = row.get("temp", "") or ""
        rmk = str(row.get("remark", "")) if row.get("remark") else ""

        try:
            if fml: total_workers += float(fml)
            if tmp: total_workers += float(tmp)
        except (ValueError, TypeError):
            pass

        ws.cell(row=current_row, column=1, value=stt)
        ws.cell(row=current_row, column=2, value=f"{d_src}\n{d_tgt}".strip())
        ws.cell(row=current_row, column=3, value=mac)
        ws.cell(row=current_row, column=4, value=fml)
        ws.cell(row=current_row, column=5, value=tmp)
        ws.cell(row=current_row, column=6, value=rmk)

        for col in range(1, 7):
            c = ws.cell(row=current_row, column=col)
            c.font = Font(name=font_name, size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border

        ws.row_dimensions[current_row].height = 32
        current_row += 1

    total_row = current_row
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    ws.cell(row=total_row, column=1, value="一共\nTổng cộng" if mode == "Trung ➔ Việt" else "Tổng cộng\n一共")
    ws.merge_cells(start_row=total_row, start_column=3, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=3, value=int(total_workers) if isinstance(total_workers, float) and total_workers.is_integer() else total_workers)

    for col in range(1, 7):
        c = ws.cell(row=total_row, column=col)
        c.font = Font(name=font_name, size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[total_row].height = 36

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 17
    ws.column_dimensions["E"].width = 17
    ws.column_dimensions["F"].width = 18

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ============================================================
# 2. XỬ LÝ DỊCH CHÍNH
# ============================================================
if uploaded_file is not None:
    is_excel = uploaded_file.name.lower().endswith('.xlsx')
    button_label = f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel" if is_excel else f"🚀 Quét OCR Ảnh/PDF & Dịch ({translation_mode})"

    if st.button(button_label, use_container_width=True):
        try:
            # TRƯỜNG HỢP 1: EXCEL FILE (.xlsx)
            if is_excel:
                with st.spinner(f"1️⃣ Đang quét các ô cần dịch theo chế độ [{translation_mode}]..."):
                    file_bytes = uploaded_file.read()
                    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

                    texts_to_translate = set()
                    for sheet in wb.worksheets:
                        for row in sheet.iter_rows():
                            for cell in row:
                                if cell.value and isinstance(cell.value, str):
                                    val = cell.value.strip()
                                    if val.startswith("="): # Bỏ qua ô chứa công thức
                                        continue
                                    if translation_mode == "Trung ➔ Việt" and has_chinese(val):
                                        texts_to_translate.add(val)
                                    elif translation_mode == "Việt ➔ Trung" and (has_vietnamese(val) or not has_chinese(val)):
                                        if len(val) > 1 and not val.isnumeric():
                                            texts_to_translate.add(val)

                    unique_texts = list(texts_to_translate)

                if not unique_texts:
                    st.warning("Không tìm thấy nội dung văn bản phù hợp với chế độ dịch đã chọn!")
                else:
                    with st.spinner(f"2️⃣ Đang dịch {len(unique_texts)} văn bản bằng thư viện chuyên dụng [{translation_mode}]..."):
                        translation_dict = {}
                        for idx, text in enumerate(unique_texts):
                            trans_text = translate_text(text, translation_mode)
                            translation_dict[text] = trans_text

                    with st.spinner("3️⃣ Đang chèn bản dịch & giữ nguyên 100% định dạng..."):
                        for sheet in wb.worksheets:
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str):
                                        orig = cell.value.strip()
                                        trans = translation_dict.get(orig, "")
                                        if trans:
                                            cell.value = f"{orig}\n{trans}"
                                            curr_align = cell.alignment
                                            cell.alignment = Alignment(
                                                horizontal=curr_align.horizontal or "center",
                                                vertical=curr_align.vertical or "center",
                                                wrap_text=True
                                            )

                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        st.success(f"✅ Đã dịch thành công ({translation_mode})!")
                        st.download_button(
                            label="⬇️ Tải File Excel Song Ngữ (.xlsx)",
                            data=output.getvalue(),
                            file_name=f"Translated_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            # TRƯỜNG HỢP 2: TẢI FILE ẢNH / PDF
            else:
                if not OCR_AVAILABLE:
                    st.error("❌ Thư viện 'easyocr' chưa được cài đặt để đọc ảnh/PDF. Vui lòng cài đặt thư viện phụ trợ.")
                else:
                    with st.spinner(f"1️⃣ Đang thực hiện OCR đọc văn bản từ hình ảnh/PDF và dịch [{translation_mode}]..."):
                        file_bytes = uploaded_file.read()
                        
                        # Chuyển đổi PDF sang danh sách ảnh nếu là file PDF
                        images = []
                        if uploaded_file.type == "application/pdf" or uploaded_file.name.lower().endswith('.pdf'):
                            if PDF_IMAGE_AVAILABLE:
                                pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
                                for page in pdf_doc:
                                    pix = page.get_pixmap()
                                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                                    images.append(img)
                            else:
                                st.error("❌ Thư viện 'PyMuPDF' (fitz) chưa được cài đặt để xử lý PDF.")
                                st.stop()
                        else:
                            images.append(Image.open(io.BytesIO(file_bytes)))

                        # Dùng EasyOCR đọc toàn bộ text tìm thấy
                        reader = get_ocr_reader()
                        extracted_lines = []
                        for img in images:
                            img_np = np.array(img)
                            results = reader.readtext(img_np)
                            for bbox, text, prob in results:
                                if prob > 0.2:  # Lọc độ tin cậy cơ bản
                                    extracted_lines.append(text.strip())

                        # Xây dựng cấu trúc dữ liệu JSON thô từ OCR đọc được để map vào bảng chấm công chuẩn
                        rows_data = []
                        for i, line in enumerate(extracted_lines, start=1):
                            translated_line = translate_text(line, translation_mode)
                            if translation_mode == "Trung ➔ Việt":
                                d_src, d_tgt = line, translated_line
                            else:
                                d_src, d_tgt = translated_line, line
                            
                            # Thử tách số lượng nếu có format số trong dòng
                            numbers = re.findall(r'\d+', line)
                            mach_val = int(numbers[0]) if len(numbers) > 0 else 0
                            form_val = int(numbers[1]) if len(numbers) > 1 else 0
                            
                            rows_data.append({
                                "stt": i,
                                "dept_src": d_src,
                                "dept_tgt": d_tgt,
                                "machines": mach_val,
                                "formal": form_val,
                                "temp": 0,
                                "remark": ""
                        })

                        parsed_data = {
                            "title_src": "Bảng Chấm Công (OCR)",
                            "title_tgt": translate_text("Bảng Chấm Công (OCR)", translation_mode),
                            "date_str": time.strftime("%Y-%m-%d"),
                            "rows": rows_data if rows_data else [{
                                "stt": 1, 
                                "dept_src": "Nội dung trích xuất", 
                                "dept_tgt": translate_text("Nội dung trích xuất", translation_mode),
                                "machines": 0, "formal": 0, "temp": 0, "remark": ""
                            }]
                        }

                    with st.spinner("2️⃣ Đang tạo bảng Excel định dạng chuẩn..."):
                        excel_bytes = build_excel_from_json(parsed_data, translation_mode)

                        st.success(f"✅ Đã trích xuất OCR và chuyển đổi sang Excel ({translation_mode}) thành công!")
                        st.download_button(
                            label="⬇️ Tải File Excel (.xlsx)",
                            data=excel_bytes.getvalue(),
                            file_name=f"Bang_cham_cong_{parsed_data.get('date_str', 'export')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

        except Exception as e:
            st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
