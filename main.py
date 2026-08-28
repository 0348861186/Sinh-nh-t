import io
import json
import re
import datetime
import time
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from google import genai
from google.genai import types

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Hai Chiều",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Tự Động Retry chống lỗi 503)")
st.caption("Hỗ trợ chọn chế độ Trung -> Việt hoặc Việt -> Trung | Giữ nguyên 100% format Excel gốc hoặc chuyển từ Ảnh/PDF.")

# ============================================================
# 1. CẤU HÌNH API KEY, BỘ LỌC HƯỚNG DỊCH & TẢI FILE
# ============================================================
col1, col2, col3 = st.columns([1, 1.2, 1.8])

with col1:
    api_key = st.text_input("Nhập GEMINI_API_KEY:", type="password")

with col2:
    translation_mode = st.radio(
        "Chế độ dịch:",
        options=["Trung ➔ Việt", "Việt ➔ Trung"],
        horizontal=True
    )

with col3:
    uploaded_file = st.file_uploader(
        "Tải lên Ảnh, PDF hoặc File Excel:", 
        type=["png", "jpg", "jpeg", "pdf", "xlsx"]
    )

# Hàm kiểm tra chuỗi có chứa chữ Hán / Tiếng Trung không
def has_chinese(text):
    if not isinstance(text, str):
        return False
    return bool(re.search(r'[\u4e00-\u9fff]', text))

# Hàm kiểm tra chuỗi có chứa tiếng Việt không
def has_vietnamese(text):
    if not isinstance(text, str):
        return False
    vietnamese_pattern = r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]'
    return bool(re.search(vietnamese_pattern, text, re.IGNORECASE))

# ============================================================
# HÀM GỌI GEMINI API CÓ CƠ CHẾ CHỐNG LỖI 503 (RETRY & FALLBACK)
# ============================================================
def generate_content_with_retry(client, contents, max_retries=3):
    """
    Hàm gọi Gemini API tự động retry khi gặp lỗi 503 và tự động chuyển Model nếu quá tải
    """
    models_to_try = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    
    for model_name in models_to_try:
        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents
                )
                return response
            except Exception as e:
                err_msg = str(e)
                if "503" in err_msg or "UNAVAILABLE" in err_msg or "429" in err_msg:
                    wait_time = (attempt + 1) * 2  # Chờ 2s, 4s, 6s...
                    st.warning(f"⚠️ Model {model_name} đang bận (Lỗi 503/429). Đang thử lại lần {attempt + 1}/{max_retries} sau {wait_time}s...")
                    time.sleep(wait_time)
                else:
                    # Nếu là lỗi khác (ví dụ sai API Key) thì throw lỗi ngay
                    raise e
        st.info(f"🔄 Chuyển sang model dự phòng tiếp theo...")
    
    raise Exception("Tất cả các mô hình Gemini hiện đang bị quá tải. Vui lòng thử lại sau ít phút!")

# Hàm tự tạo file Excel chuẩn đẹp nếu đọc từ Ảnh / PDF
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

    if mode == "Trung ➔ Việt":
        headers = [("STT", "STT"), ("部门", "Bộ phận"), ("开几台机", "Số máy mở"), ("正式工", "Chính thức"), ("临时工", "Thời vụ"), ("备注", "Ghi chú")]
    else:
        headers = [("STT", "STT"), ("Bộ phận", "部门"), ("Số máy mở", "开几台机"), ("Chính thức", "正式工"), ("Thời vụ", "临时工"), ("Ghi chú", "备注")]

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
        except:
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
# 2. XỬ LÝ DỊCH CHÍNH (PHÂN NHÁNH ẢNH / EXCEL + HƯỚNG DỊCH)
# ============================================================
if uploaded_file is not None:
    is_excel = uploaded_file.name.lower().endswith('.xlsx')
    
    button_label = f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel" if is_excel else f"🚀 AI Quét Ảnh/PDF & Dịch ({translation_mode})"
    
    if st.button(button_label, use_container_width=True):
        if not api_key:
            st.error("Vui lòng nhập GEMINI_API_KEY!")
        else:
            try:
                client = genai.Client(api_key=api_key)

                # ----------------------------------------------------
                # TRƯỜNG HỢP 1: TẢI UP FILE EXCEL (.xlsx)
                # ----------------------------------------------------
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
                                        if translation_mode == "Trung ➔ Việt" and has_chinese(val):
                                            texts_to_translate.add(val)
                                        elif translation_mode == "Việt ➔ Trung" and (has_vietnamese(val) or not has_chinese(val)):
                                            if len(val) > 1 and not val.isnumeric():
                                                texts_to_translate.add(val)

                        unique_texts = list(texts_to_translate)

                    if not unique_texts:
                        st.warning("Không tìm thấy nội dung phù hợp với chế độ dịch đã chọn!")
                    else:
                        with st.spinner(f"2️⃣ AI đang dịch {len(unique_texts)} từ/câu [{translation_mode}]..."):
                            src_lang = "tiếng Trung" if translation_mode == "Trung ➔ Việt" else "tiếng Việt"
                            tgt_lang = "tiếng Việt" if translation_mode == "Trung ➔ Việt" else "tiếng Trung"

                            prompt = f"""
                            Bạn là chuyên gia dịch thuật chuyên nghiệp trong lĩnh vực nhân sự, nhà xưởng và bảng chấm công.
                            Hãy dịch danh sách {src_lang} sau đây sang {tgt_lang}.
                            
                            Danh sách nguồn ({src_lang}):
                            {json.dumps(unique_texts, ensure_ascii=False, indent=2)}

                            Trả về kết quả dưới dạng MỘT JSON OBJECT duy nhất (không dùng markdown code blocks).
                            Key là văn bản gốc, Value là bản dịch tương ứng.
                            """

                            # Gọi qua hàm retry an toàn
                            response = generate_content_with_retry(client, prompt)

                            clean_json = response.text.replace("```json", "").replace("```", "").strip()
                            translation_dict = json.loads(clean_json)

                        with st.spinner("3️⃣ Đang chèn dịch & giữ nguyên 100% định dạng gốc..."):
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

                            st.success(f"✅ Đã dịch thành công ({translation_mode})! File Excel giữ nguyên 100% màu sắc & font gốc.")
                            st.download_button(
                                label="⬇️ Tải File Excel Song Ngữ (.xlsx)",
                                data=output.getvalue(),
                                file_name=f"Translated_{uploaded_file.name}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

                # ----------------------------------------------------
                # TRƯỜNG HỢP 2: TẢI UP FILE ẢNH / PDF
                # ----------------------------------------------------
                else:
                    with st.spinner(f"1️⃣ AI đang quét hình ảnh/PDF và dịch theo hướng [{translation_mode}]..."):
                        file_bytes = uploaded_file.read()
                        file_part = types.Part.from_bytes(data=file_bytes, mime_type=uploaded_file.type)

                        src_lang = "tiếng Trung" if translation_mode == "Trung ➔ Việt" else "tiếng Việt"
                        tgt_lang = "tiếng Việt" if translation_mode == "Trung ➔ Việt" else "tiếng Trung"

                        prompt = f"""
                        Hãy phân tích hình ảnh/file bảng chấm công này và trích xuất dữ liệu dưới dạng JSON thuần túy (không dùng markdown backticks).
                        Dịch các nội dung từ {src_lang} sang {tgt_lang}.
                        
                        Cấu trúc JSON yêu cầu:
                        {{
                            "title_src": "Tiêu đề ngôn ngữ gốc ({src_lang})",
                            "title_tgt": "Tiêu đề dịch ({tgt_lang})",
                            "date_str": "YYYY-MM-DD",
                            "rows": [
                                {{
                                    "stt": 1,
                                    "dept_src": "Tên bộ phận ngôn ngữ gốc ({src_lang})",
                                    "dept_tgt": "Dịch tên bộ phận sang ({tgt_lang})",
                                    "machines": 5,
                                    "formal": 3,
                                    "temp": 2,
                                    "remark": "Ghi chú nếu có"
                                }}
                            ]
                        }}
                        Lưu ý: Dịch chính xác nghĩa cho từng bộ phận/công việc.
                        """

                        # Gọi qua hàm retry an toàn
                        response = generate_content_with_retry(client, [file_part, prompt])

                        clean_json_str = response.text.replace("```json", "").replace("```", "").strip()
                        parsed_data = json.loads(clean_json_str)

                    with st.spinner("2️⃣ Đang dựng bảng Excel chuẩn đẹp..."):
                        excel_bytes = build_excel_from_json(parsed_data, translation_mode)

                        st.success(f"✅ Đã quét ảnh và chuyển đổi sang Excel ({translation_mode}) thành công!")
                        st.download_button(
                            label="⬇️ Tải Xuất File Excel (.xlsx)",
                            data=excel_bytes.getvalue(),
                            file_name=f"Bang_cham_cong_{parsed_data.get('date_str', 'export')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

            except Exception as e:
                st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")

