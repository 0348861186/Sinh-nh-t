import io
import json
import re
import time
import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from deep_translator import GoogleTranslator
import easyocr
from PIL import Image

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Hai Chiều (Thư Viện Chuyên Dụng)",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Thư Viện Chuyên Dụng)")
st.caption("Hỗ trợ chọn chế độ Trung ➔ Việt hoặc Việt ➔ Trung | Giữ nguyên 100% format Excel gốc hoặc chuyển từ Ảnh/PDF mà không tốn Quota AI.")

# Cache Reader EasyOCR để tránh load lại model nhiều lần gây chậm
@st.cache_resource
def get_ocr_reader():
    # Sử dụng 'ch_sim' (Tiếng Trung Giản Thể) và 'en' (Tiếng Anh/Latinh) để tránh lỗi 'zh is not supported'
    return easyocr.Reader(['ch_sim', 'en'], gpu=False)

# ============================================================
# 1. BỘ LỌC HƯỚNG DỊCH & TẢI FILE
# ============================================================
col1, col2 = st.columns([1, 2])

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
# HÀM DỊCH THUẬT DÙNG THƯ VIỆN DEEP-TRANSLATOR
# ============================================================
def translate_texts_batch(texts_list, mode):
    """
    Dịch danh sách văn bản dựa trên deep-translator (Google Translate Engine)
    """
    if not texts_list:
        return {}
    
    src_code = 'zh-CN' if mode == "Trung ➔ Việt" else 'vi'
    tgt_code = 'vi' if mode == "Trung ➔ Việt" else 'zh-CN'
    
    translator = GoogleTranslator(source=src_code, target=tgt_code)
    translation_dict = {}
    
    # Dịch từng từ/câu trong danh sách
    for text in texts_list:
        try:
            translated = translator.translate(text)
            translation_dict[text] = translated
        except Exception:
            translation_dict[text] = text # Giữ nguyên nếu dịch lỗi
            
    return translation_dict

# ============================================================
# HÀM XỬ LÝ QUÉT OCR TỪ ẢNH / PDF
# ============================================================
def process_image_or_pdf_to_data(file_bytes, file_type, mode):
    """
    Quét nội dung chữ từ Ảnh/PDF bằng EasyOCR và trích xuất dữ liệu
    """
    reader = get_ocr_reader()
    
    # Đọc ảnh từ Bytes
    image = Image.open(io.BytesIO(file_bytes)).convert("RGB")
    
    # Quét chữ từ ảnh
    results = reader.readtext(file_bytes, detail=0)
    
    src_lang = "zh-CN" if mode == "Trung ➔ Việt" else "vi"
    tgt_lang = "vi" if mode == "Trung ➔ Việt" else "zh-CN"
    translator = GoogleTranslator(source=src_lang, target=tgt_lang)
    
    # Bóc tách các chuỗi tìm thấy
    extracted_lines = [res.strip() for res in results if res.strip()]
    
    # Tìm ngày tháng nếu có (định dạng YYYY-MM-DD hoặc YYYY/MM/DD)
    date_str = ""
    for line in extracted_lines:
        match = re.search(r'\d{4}[-/.]\d{1,2}[-/.]\d{1,2}', line)
        if match:
            date_str = match.group(0)
            break
            
    # Xây dựng cấu trúc dữ liệu JSON giống logic code gốc
    title_src = extracted_lines[0] if len(extracted_lines) > 0 else "Bảng Chấm Công"
    try:
        title_tgt = translator.translate(title_src)
    except:
        title_tgt = title_src

    rows = []
    # Trích xuất các dòng thông tin bộ phận
    dept_lines = [l for l in extracted_lines[1:] if not re.search(r'^\d+$', l)]
    
    for idx, line in enumerate(dept_lines, start=1):
        if idx > 15: # Giới hạn tối đa 15 dòng mẫu
            break
        try:
            trans_line = translator.translate(line)
        except:
            trans_line = line
            
        rows.append({
            "stt": idx,
            "dept_src": line,
            "dept_tgt": trans_line,
            "machines": "",
            "formal": "",
            "temp": "",
            "remark": ""
        })

    parsed_data = {
        "title_src": title_src,
        "title_tgt": title_tgt,
        "date_str": date_str if date_str else time.strftime("%Y-%m-%d"),
        "rows": rows
    }
    
    return parsed_data

# ============================================================
# HÀM DỰNG FILE EXCEL TỪ DỮ LIỆU (GIỮ NGUYÊN CODE GỐC)
# ============================================================
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
    button_label = f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel" if is_excel else f"🚀 Quét Ảnh/PDF & Dịch ({translation_mode})"

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
                                    if val.startswith("="):  # Bỏ qua ô chứa công thức
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
                    with st.spinner(f"2️⃣ Đang dịch {len(unique_texts)} văn bản [{translation_mode}] qua Google Translator..."):
                        translation_dict = translate_texts_batch(unique_texts, translation_mode)

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
                with st.spinner(f"1️⃣ Đang quét OCR dữ liệu hình ảnh/PDF và dịch [{translation_mode}]..."):
                    file_bytes = uploaded_file.read()
                    parsed_data = process_image_or_pdf_to_data(file_bytes, uploaded_file.type, translation_mode)

                with st.spinner("2️⃣ Đang tạo bảng Excel định dạng chuẩn..."):
                    excel_bytes = build_excel_from_json(parsed_data, translation_mode)

                    st.success(f"✅ Đã trích xuất và chuyển đổi sang Excel ({translation_mode}) thành công!")
                    st.download_button(
                        label="⬇️ Tải File Excel (.xlsx)",
                        data=excel_bytes.getvalue(),
                        file_name=f"Bang_cham_cong_{parsed_data.get('date_str', 'export')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
