import io
import json
import re
import time
import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 1. AI CHỈ DÙNG ĐỂ ĐỌC CHỮ TỪ ẢNH (OCR)
try:
    from google import genai
    from google.genai import types
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

# 2. THƯ VIỆN CHUYÊN DỤNG DÙNG ĐỂ DỊCH 100%
try:
    from deep_translator import GoogleTranslator
    HAS_DEEP_TRANSLATOR = True
except ImportError:
    HAS_DEEP_TRANSLATOR = False

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Hai Chiều",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Excel + Ảnh/PDF)")
st.caption("AI chỉ làm nhiệm vụ đọc chữ từ Ảnh/PDF ➔ 100% Việc Dịch Thuật do THƯ VIỆN CHUYÊN DỤNG thực hiện.")

# ============================================================
# 1. CẤU HÌNH BỘ LỌC HƯỚNG DỊCH & TẢI FILE
# ============================================================
col1, col2, col3 = st.columns([1.2, 1, 1.8])

with col1:
    api_key = st.text_input("Nhập GEMINI_API_KEY (Chỉ cần khi tải File Ảnh/PDF để đọc chữ):", type="password")

with col2:
    translation_mode = st.radio(
        "Chế độ dịch:",
        options=["Trung ➔ Việt", "Việt ➔ Trung"],
        horizontal=True
    )

with col3:
    uploaded_file = st.file_uploader(
        "Tải lên File Excel (.xlsx) hoặc File Ảnh/PDF:", 
        type=["png", "jpg", "jpeg", "pdf", "xlsx"]
    )

def has_chinese(text):
    return bool(re.search(r'[\u4e00-\u9fff]', str(text))) if text else False

def has_vietnamese(text):
    if not isinstance(text, str):
        return False
    vietnamese_pattern = r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]'
    return bool(re.search(vietnamese_pattern, text, re.IGNORECASE))

# ============================================================
# HÀM DỊCH CHÍNH: DÙNG 100% THƯ VIỆN DEEP-TRANSLATOR
# ============================================================
def translate_texts_with_library(text_list, mode):
    """
    Toàn bộ việc dịch cho cả File Excel và File Ảnh đều đi qua hàm này.
    Dùng thư viện chuyên dụng, không tốn Quota AI.
    """
    if not HAS_DEEP_TRANSLATOR:
        st.error("❌ Thư viện 'deep-translator' chưa được cài đặt!")
        return {item: item for item in text_list}

    src_code = "zh-CN" if mode == "Trung ➔ Việt" else "vi"
    tgt_code = "vi" if mode == "Trung ➔ Việt" else "zh-CN"

    translator = GoogleTranslator(source=src_code, target=tgt_code)
    translated_dict = {}

    progress_bar = st.progress(0)
    total = len(text_list)

    for idx, item in enumerate(text_list):
        try:
            if str(item).strip():
                res = translator.translate(str(item))
                translated_dict[item] = res
            else:
                translated_dict[item] = item
        except Exception:
            translated_dict[item] = item
        
        progress_bar.progress((idx + 1) / total)
        time.sleep(0.02)

    progress_bar.empty()
    return translated_dict

# ============================================================
# HÀM OCR: AI CHỈ TRÍCH XUẤT CHỮ THÔ (KHÔNG DỊCH)
# ============================================================
def ocr_extract_text_only(api_key, file_bytes, mime_type):
    """Chỉ dùng AI để đọc văn bản gốc từ Ảnh/PDF ra JSON, KHÔNG DỊCH"""
    if not HAS_GENAI:
        raise Exception("Chưa cài 'google-genai' để đọc ảnh.")
    
    client = genai.Client(api_key=api_key)
    file_part = types.Part.from_bytes(data=file_bytes, mime_type=mime_type)

    prompt = """
    Hãy đọc và trích xuất TOÀN BỘ VĂN BẢN GỐC trong bảng chấm công này thành dạng JSON.
    LƯU Ý: GIỮ NGUYÊN VĂN BẢN GỐC, KHÔNG DỊCH BẤT KỲ TỪ NÀO.

    Cấu trúc JSON:
    {
        "title_src": "Tiêu đề gốc trên ảnh",
        "date_str": "YYYY-MM-DD",
        "rows": [
            {
                "stt": 1,
                "dept_src": "Tên bộ phận gốc",
                "machines": 5,
                "formal": 3,
                "temp": 2,
                "remark": "Ghi chú gốc"
            }
        ]
    }
    """

    config = types.GenerateContentConfig(response_mime_type="application/json")
    
    for model_name in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[file_part, prompt],
                config=config
            )
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            return json.loads(match.group(0)) if match else json.loads(response.text)
        except Exception:
            time.sleep(1)
            
    raise Exception("Lỗi khi đọc file ảnh/PDF.")

# Dựng Excel từ JSON sau khi đã dùng Thư viện dịch xong
def build_excel_from_translated_json(data, mode):
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
        rmk_src = str(row.get("remark_src", "")) if row.get("remark_src") else ""
        rmk_tgt = str(row.get("remark_tgt", "")) if row.get("remark_tgt") else ""

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
        ws.cell(row=current_row, column=6, value=f"{rmk_src}\n{rmk_tgt}".strip() if rmk_tgt else rmk_src)

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
# 2. XỬ LÝ CHÍNH
# ============================================================
if uploaded_file is not None:
    is_excel = uploaded_file.name.lower().endswith('.xlsx')

    if st.button("🚀 Bắt Đầu Xử Lý & Dịch (Bằng Thư Viện)", use_container_width=True):
        try:
            # TRƯỜNG HỢP 1: FILE EXCEL (.xlsx)
            if is_excel:
                file_bytes = uploaded_file.read()
                wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

                texts_to_translate = set()
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows():
                        for cell in row:
                            if cell.value and isinstance(cell.value, str):
                                val = cell.value.strip()
                                if val.startswith("="): continue
                                if translation_mode == "Trung ➔ Việt" and has_chinese(val):
                                    texts_to_translate.add(val)
                                elif translation_mode == "Việt ➔ Trung" and (has_vietnamese(val) or not has_chinese(val)):
                                    if len(val) > 1 and not val.isnumeric():
                                        texts_to_translate.add(val)

                unique_texts = list(texts_to_translate)

                if unique_texts:
                    st.info(f"Đang dùng THƯ VIỆN dịch {len(unique_texts)} đoạn văn bản...")
                    translation_dict = translate_texts_with_library(unique_texts, translation_mode)

                    for sheet in wb.worksheets:
                        for row in sheet.iter_rows():
                            for cell in row:
                                if cell.value and isinstance(cell.value, str):
                                    orig = cell.value.strip()
                                    trans = translation_dict.get(orig, "")
                                    if trans and trans != orig:
                                        cell.value = f"{orig}\n{trans}"
                                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)

                    st.success("✅ Dịch File Excel bằng Thư viện thành công!")
                    st.download_button("⬇️ Tải File Excel Song Ngữ", output.getvalue(), file_name=f"Translated_{uploaded_file.name}")

            # TRƯỜNG HỢP 2: FILE ẢNH / PDF
            else:
                if not api_key:
                    st.error("❌ Cần GEMINI_API_KEY để AI đọc chữ từ ảnh (OCR).")
                else:
                    # Bước 1: AI chỉ quét đọc chữ
                    with st.spinner("1️⃣ AI đang quét chữ từ Ảnh/PDF (Chưa dịch)..."):
                        file_bytes = uploaded_file.read()
                        raw_data = ocr_extract_text_only(api_key, file_bytes, uploaded_file.type)

                    # Bước 2: Thu thập chữ thô ➔ Đưa qua THƯ VIỆN dịch
                    with st.spinner("2️⃣ Đưa toàn bộ chữ vừa quét qua THƯ VIỆN để dịch..."):
                        texts_to_trans = []
                        if raw_data.get("title_src"): texts_to_trans.append(raw_data["title_src"])
                        for r in raw_data.get("rows", []):
                            if r.get("dept_src"): texts_to_trans.append(r["dept_src"])
                            if r.get("remark"): texts_to_trans.append(r["remark"])

                        # GỌI THƯ VIỆN DỊCH Ở ĐÂY
                        trans_dict = translate_texts_with_library(texts_to_trans, translation_mode)

                        # Gán bản dịch từ Thư viện vào dữ liệu
                        raw_data["title_tgt"] = trans_dict.get(raw_data.get("title_src"), "")
                        for r in raw_data.get("rows", []):
                            r["dept_tgt"] = trans_dict.get(r.get("dept_src"), "")
                            r["remark_src"] = r.get("remark", "")
                            r["remark_tgt"] = trans_dict.get(r.get("remark"), "")

                    # Bước 3: Dựng File Excel
                    excel_bytes = build_excel_from_translated_json(raw_data, translation_mode)
                    st.success("✅ Đã quét chữ từ Ảnh và Dịch bằng THƯ VIỆN thành công!")
                    st.download_button("⬇️ Tải File Excel Song Ngữ", excel_bytes.getvalue(), file_name="Bang_cham_cong.xlsx")

        except Exception as e:
            st.error(f"❌ Xảy ra lỗi: {e}")
