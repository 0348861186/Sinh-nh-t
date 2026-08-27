import io
import json
import re
import time
import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from google import genai
from google.genai import types

# Import thư viện dịch dự phòng chuyên dụng
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

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Tự Động Retry & Fallback Thư Viện)")
st.caption("Hỗ trợ chọn chế độ Trung ➔ Việt hoặc Việt ➔ Trung | Tự động chuyển sang thư viện dịch miễn phí khi Gemini hết Quota.")

# Cache Google GenAI Client để tránh khởi tạo nhiều lần
@st.cache_resource
def get_genai_client(key):
    if not key:
        return None
    return genai.Client(api_key=key)

# ============================================================
# 1. CẤU HÌNH API KEY, BỘ LỌC HƯỚNG DỊCH & TẢI FILE
# ============================================================
col1, col2, col3 = st.columns([1, 1.2, 1.8])

with col1:
    api_key = st.text_input("Nhập GEMINI_API_KEY (Không bắt buộc nếu dùng Fallback):", type="password")

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

# Hàm kiểm tra chuỗi
def has_chinese(text):
    return bool(re.search(r'[\u4e00-\u9fff]', str(text))) if text else False

def has_vietnamese(text):
    if not isinstance(text, str):
        return False
    vietnamese_pattern = r'[àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđĐ]'
    return bool(re.search(vietnamese_pattern, text, re.IGNORECASE))

def safe_extract_json(text_content):
    """Trích xuất JSON an toàn bằng Regex để tránh lỗi Markdown"""
    try:
        match = re.search(r'\{.*\}', text_content, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        return json.loads(text_content)
    except Exception as e:
        raise ValueError(f"Dữ liệu trả về không đúng định dạng JSON: {e}")

# ============================================================
# HÀM DỊCH DỰ PHÒNG BẰNG THƯ VIỆN CHUYÊN DỤNG (DEEP-TRANSLATOR)
# ============================================================
def translate_via_library(text_list, src_lang_code, tgt_lang_code):
    """
    Dịch danh sách văn bản bằng thư viện deep-translator khi AI gặp lỗi / hết Quota.
    src_lang_code: 'zh-CN' hoặc 'vi'
    tgt_lang_code: 'vi' hoặc 'zh-CN'
    """
    translated_dict = {}
    if not HAS_DEEP_TRANSLATOR:
        st.error("❌ Thư viện 'deep-translator' chưa được cài đặt. Vui lòng chạy: pip install deep-translator")
        return {item: item for item in text_list}

    st.info("🔄 Đang chuyển sang thư viện dịch chuyên dụng Google Translator (Free & Unlimited)...")
    translator = GoogleTranslator(source=src_lang_code, target=tgt_lang_code)

    for item in text_list:
        try:
            if item.strip():
                res = translator.translate(item)
                translated_dict[item] = res
            else:
                translated_dict[item] = item
        except Exception as e:
            st.warning(f"⚠️ Lỗi dịch từ '{item}': {e}")
            translated_dict[item] = item
        time.sleep(0.1) # Tránh bị rate limit

    return translated_dict

# ============================================================
# HÀM GỌI GEMINI API CÓ RE-TRY VÀ TỰ ĐỘNG FALLBACK SANG THƯ VIỆN
# ============================================================
def translate_texts_with_fallback(client, unique_texts, mode):
    """
    Ưu tiên dịch bằng Gemini. Nếu hết Quota (429/503/Exception) -> Tự động chuyển sang thư viện chuyên dụng.
    """
    src_lang = "tiếng Trung" if mode == "Trung ➔ Việt" else "tiếng Việt"
    tgt_lang = "tiếng Việt" if mode == "Trung ➔ Việt" else "tiếng Trung"
    src_code = "zh-CN" if mode == "Trung ➔ Việt" else "vi"
    tgt_code = "vi" if mode == "Trung ➔ Việt" else "zh-CN"

    # Nếu không cung cấp API Key, chuyển trực tiếp sang thư viện
    if not client:
        st.warning("⚠️ Không tìm thấy Gemini API Key. Đang sử dụng thư viện dịch miễn phí...")
        return translate_via_library(unique_texts, src_code, tgt_code)

    models_to_try = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    prompt = f"""
    Bạn là chuyên gia dịch thuật chuyên nghiệp trong lĩnh vực nhân sự, nhà xưởng và bảng chấm công.
    Hãy dịch danh sách các từ/câu {src_lang} sau đây sang {tgt_lang}.
    Dữ liệu nguồn:
    {json.dumps(unique_texts, ensure_ascii=False, indent=2)}

    Yêu cầu: Trả về một Key-Value JSON Object duy nhất với Key là văn bản gốc và Value là bản dịch tương ứng.
    """
    config = types.GenerateContentConfig(response_mime_type="application/json")

    for model_name in models_to_try:
        for attempt in range(2): # Retry 2 lần cho mỗi model
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=config
                )
                return json.loads(response.text)
            except Exception as e:
                err_msg = str(e)
                st.warning(f"⚠️ Lỗi kết nối Gemini ({model_name}): {err_msg[:100]}...")
                time.sleep(1)

    # Nếu tất cả các model Gemini đều thất bại (ví dụ: Hết Quota 429) -> Dùng Fallback Library
    st.error("🚨 Gemini API đã hết Quota hoặc bận. Đang tự động chuyển sang thư viện dịch chuyên dụng dự phòng!")
    return translate_via_library(unique_texts, src_code, tgt_code)

def generate_ocr_with_fallback(client, contents, mode):
    """
    Gọi AI để OCR Ảnh/PDF. Nếu Gemini hết quota khi dịch, hỗ trợ dịch bổ sung bằng thư viện.
    """
    models_to_try = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    config = types.GenerateContentConfig(response_mime_type="application/json")

    for model_name in models_to_try:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=contents,
                config=config
            )
            return safe_extract_json(response.text)
        except Exception as e:
            st.warning(f"⚠️ Thử lại với model tiếp theo do lỗi: {e}")
            time.sleep(1)

    raise Exception("Không thể thực hiện OCR với Gemini. Vui lòng kiểm tra lại API Key hoặc quota của bạn.")

def fallback_translate_json_data(parsed_data, mode):
    """Điền các bản dịch còn thiếu trong JSON bằng thư viện nếu Gemini chưa dịch xong"""
    src_code = "zh-CN" if mode == "Trung ➔ Việt" else "vi"
    tgt_code = "vi" if mode == "Trung ➔ Việt" else "zh-CN"
    
    if not parsed_data.get("title_tgt") and parsed_data.get("title_src"):
        parsed_data["title_tgt"] = translate_via_library([parsed_data["title_src"]], src_code, tgt_code).get(parsed_data["title_src"], "")
        
    for row in parsed_data.get("rows", []):
        if not row.get("dept_tgt") and row.get("dept_src"):
            row["dept_tgt"] = translate_via_library([row["dept_src"]], src_code, tgt_code).get(row["dept_src"], "")
            
    return parsed_data

# Hàm dựng file Excel từ JSON (khi scan Ảnh/PDF)
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
    button_label = f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel" if is_excel else f"🚀 AI Quét Ảnh/PDF & Dịch ({translation_mode})"

    if st.button(button_label, use_container_width=True):
        try:
            client = get_genai_client(api_key) if api_key else None

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
                    with st.spinner(f"2️⃣ Đang dịch {len(unique_texts)} văn bản [{translation_mode}]..."):
                        # Gọi dịch với cơ chế tự động chuyển sang Thư viện khi gặp lỗi/Hết Quota
                        translation_dict = translate_texts_with_fallback(client, unique_texts, translation_mode)

                    with st.spinner("3️⃣ Đang chèn bản dịch & giữ nguyên 100% định dạng..."):
                        for sheet in wb.worksheets:
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(cell.value, str):
                                        orig = cell.value.strip()
                                        trans = translation_dict.get(orig, "")
                                        if trans and trans != orig:
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
                if not client:
                    st.error("❌ Xử lý Quét Ảnh/PDF bắt buộc cần có GEMINI_API_KEY để OCR!")
                else:
                    with st.spinner(f"1️⃣ AI đang đọc dữ liệu hình ảnh/PDF và dịch [{translation_mode}]..."):
                        file_bytes = uploaded_file.read()
                        file_part = types.Part.from_bytes(data=file_bytes, mime_type=uploaded_file.type)

                        src_lang = "tiếng Trung" if translation_mode == "Trung ➔ Việt" else "tiếng Việt"
                        tgt_lang = "tiếng Việt" if translation_mode == "Trung ➔ Việt" else "tiếng Trung"

                        prompt = f"""
                        Hãy phân tích hình ảnh/PDF bảng chấm công này và trích xuất toàn bộ dữ liệu dưới dạng JSON.
                        Dịch các nội dung từ {src_lang} sang {tgt_lang}.

                        Định dạng JSON yêu cầu:
                        {{
                            "title_src": "Tiêu đề gốc ({src_lang})",
                            "title_tgt": "Tiêu đề dịch ({tgt_lang})",
                            "date_str": "YYYY-MM-DD",
                            "rows": [
                                {{
                                    "stt": 1,
                                    "dept_src": "Bộ phận gốc ({src_lang})",
                                    "dept_tgt": "Bộ phận dịch ({tgt_lang})",
                                    "machines": 5,
                                    "formal": 3,
                                    "temp": 2,
                                    "remark": "Ghi chú"
                                }}
                            ]
                        }}
                        """

                        parsed_data = generate_ocr_with_fallback(client, [file_part, prompt], translation_mode)
                        
                        # Kiểm tra xem có trường nào chưa dịch do đứt quãng không, nếu có dùng thư viện bổ sung
                        parsed_data = fallback_translate_json_data(parsed_data, translation_mode)

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
