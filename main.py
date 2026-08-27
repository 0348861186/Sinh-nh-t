import io
import json
import re
import time
import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image

# Thư viện dịch thuật chuyên dụng
from deep_translator import GoogleTranslator
# Thư viện OCR đọc chữ từ ảnh
import easyocr
import numpy as np

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

# Cache EasyOCR Reader linh hoạt theo danh sách ngôn ngữ
@st.cache_resource
def get_ocr_reader(lang_tuple):
    return easyocr.Reader(list(lang_tuple))

# ============================================================
# 1. CẤU HÌNH BỘ LỌC HƯỚNG DỊCH & TẢI FILE
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
        type=["png", "jpg", "jpeg", "xlsx"]
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
# HÀM DỊCH CHUYÊN DỤNG BẰNG DEEP-TRANSLATOR
# ============================================================
def translate_texts_batch(text_list, src_lang, tgt_lang):
    translator = GoogleTranslator(source=src_lang, target=tgt_lang)
    translation_dict = {}
    
    for text in text_list:
        if not text.strip():
            continue
        try:
            translated = translator.translate(text)
            translation_dict[text] = translated
        except Exception:
            translation_dict[text] = text  
            
    return translation_dict

# Hàm dựng file Excel chuẩn format
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

    full_title = f"{dt_str} {top_title}\n{bot_title} {dt_str}".strip()
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
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=Thread if 'Thread' in globals() else Alignment(horizontal="center", vertical="center", wrap_text=True))
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
            src_code = 'zh-CN' if translation_mode == "Trung ➔ Việt" else 'vi'
            tgt_code = 'vi' if translation_mode == "Trung ➔ Việt" else 'zh-CN'

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
                                    if val.startswith("="):  
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
                        translation_dict = translate_texts_batch(unique_texts, src_code, tgt_code)

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

            # TRƯỜNG HỢP 2: TẢI FILE ẢNH / PDF (Xử lý thông minh phân tách cấu trúc bảng theo tọa độ)
            else:
                with st.spinner(f"1️⃣ Đang OCR nhận diện văn bản và phân tích cấu trúc bảng từ hình ảnh..."):
                    image = Image.open(uploaded_file).convert('RGB')
                    image_np = np.array(image)

                    if translation_mode == "Trung ➔ Việt":
                        ocr_langs = ('ch_sim', 'en')
                    else:
                        ocr_langs = ('vi', 'en')

                    reader = get_ocr_reader(ocr_langs)
                    # Lấy chi tiết bounding box (tọa độ): [ [ [[x1,y1], [x2,y2], [x3,y3], [x4,y4]], text, confidence ], ... ]
                    ocr_results = reader.readtext(image_np, detail=1)

                    # Lọc và sắp xếp các ô theo tọa độ hàng (y trung bình)
                    items = []
                    for bbox, text, prob in ocr_results:
                        text = text.strip()
                        if not text:
                            continue
                        y_center = sum([pt[1] for pt in bbox]) / 4
                        x_center = sum([pt[0] for pt in bbox]) / 4
                        items.append({'text': text, 'y': y_center, 'x': x_center})

                    # Sắp xếp theo chiều dọc (trên xuống dưới)
                    items.sort(key=lambda item: item['y'])

                    # Gom nhóm các phần tử thuộc cùng một hàng (ngưỡng chênh lệch y khoảng 15 pixels)
                    rows_grouped = []
                    current_row = []
                    last_y = -1

                    for item in items:
                        if last_y == -1 or abs(item['y'] - last_y) < 18:
                            current_row.append(item)
                            last_y = item['y'] if last_y == -1 else (last_y + item['y']) / 2
                        else:
                            # Sắp xếp các cột trong hàng theo chiều ngang (trái sang phải)
                            current_row.sort(key=lambda col: col['x'])
                            rows_grouped.append(current_row)
                            current_row = [item]
                            last_y = item['y']

                    if current_row:
                        current_row.sort(key=lambda col: col['x'])
                        rows_grouped.append(current_row)

                    # Phân tích trích xuất dữ liệu bảng chấm công chuẩn
                    title_text = ""
                    table_rows = []
                    
                    # Các từ khóa tiêu đề cột cần bỏ qua ở thân bảng
                    header_keywords = ["stt", "部分", "部门", "开几台机", "正式工", "临时工", "备注", "số máy", "chính thức", "thời vụ", "ghi chú"]

                    for r_idx, r_items in enumerate(rows_grouped):
                        row_texts = [i['text'] for i in r_items]
                        full_row_str = " ".join(row_texts)

                        # Dòng đầu tiên thường là tiêu đề bảng
                        if r_idx == 0 or ("上班" in full_row_str) or ("考勤" in full_row_str):
                            if not title_text:
                                title_text = full_row_str
                            continue

                        # Bỏ qua các dòng tiêu đề cột
                        is_header = False
                        for kw in header_keywords:
                            if kw.lower() in full_row_str.lower():
                                is_header = True
                                break
                        if is_header:
                            continue

                        # Phân tách các cột dữ liệu theo thứ tự x từ trái sang phải
                        # Cột 1: STT (thường là số thứ tự ngắn)
                        # Cột 2: Tên bộ phận
                        # Cột 3: Số máy mở
                        # Cột 4: Chính thức
                        # Cột 5: Thời vụ
                        # Cột 6: Ghi chú
                        if len(row_texts) >= 2:
                            # Đoán STT ở cột đầu tiên nếu là số
                            stt_val = row_texts[0] if row_texts[0].isdigit() else len(table_rows) + 1
                            dept_val = row_texts[1] if not row_texts[0].isdigit() else (row_texts[1] if len(row_texts) > 1 else "")
                            
                            # Lấy các cột số liệu phía sau nếu có
                            nums = row_texts[2:] if not row_texts[0].isdigit() else row_texts[2:]
                            
                            machines = nums[0] if len(nums) > 0 and nums[0].isdigit() else ""
                            formal = nums[1] if len(nums) > 1 and nums[1].isdigit() else ""
                            temp = nums[2] if len(nums) > 2 and nums[2].isdigit() else ""
                            remark = nums[3] if len(nums) > 3 else ""

                            table_rows.append({
                                "stt": stt_val,
                                "dept_src": dept_val,
                                "machines": machines,
                                "formal": formal,
                                "temp": temp,
                                "remark": remark
                            })

                # Trích xuất toàn bộ text cần dịch (Tiêu đề + Tên các bộ phận)
                texts_to_translate = [title_text] if title_text else []
                for r in table_rows:
                    if r["dept_src"]:
                        texts_to_translate.append(r["dept_src"])

                with st.spinner(f"2️⃣ Đang dịch nội dung sang [{translation_mode}]..."):
                    translation_dict = translate_texts_batch(texts_to_translate, src_code, tgt_code)

                    translated_title = translation_dict.get(title_text, title_text)

                    final_rows = []
                    for r in table_rows:
                        d_src = r["dept_src"]
                        d_tgt = translation_dict.get(d_src, "")
                        final_rows.append({
                            "stt": r["stt"],
                            "dept_src": d_src,
                            "dept_tgt": d_tgt,
                            "machines": r["machines"],
                            "formal": r["formal"],
                            "temp": r["temp"],
                            "remark": r["remark"]
                        })

                    # Tách ngày tháng từ tiêu đề nếu có (ví dụ: 2026年08月26日)
                    date_match = re.search(r'\d{4}[^\d]+\d{1,2}[^\d]+\d{1,2}', title_text)
                    date_str = date_match.group(0) if date_match else ""

                    parsed_data = {
                        "title_src": title_text,
                        "title_tgt": translated_title,
                        "date_str": date_str,
                        "rows": final_rows
                    }

                with st.spinner("3️⃣ Đang tạo bảng Excel định dạng chuẩn..."):
                    excel_bytes = build_excel_from_json(parsed_data, translation_mode)

                    st.success(f"✅ Đã trích xuất và chuyển đổi sang Excel ({translation_mode}) thành công!")
                    st.download_button(
                        label="⬇️ Tải File Excel (.xlsx)",
                        data=excel_bytes.getvalue(),
                        file_name=f"Bang_cham_cong_exported.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
