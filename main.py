import io
import json
import re
import time
import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Thư viện dịch chuyên dụng
try:
    from deep_translator import GoogleTranslator
    HAS_TRANSLATOR = True
except ImportError:
    HAS_TRANSLATOR = False

# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Hai Chiều",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Dịch & Xuất Bảng Chấm Công Song Ngữ (Dùng Thư Viện Chuyên Dụng)")
st.caption("Hỗ trợ chọn chế độ Trung ➔ Việt hoặc Việt ➔ Trung | Giữ nguyên 100% format Excel gốc (Không lo hết Quota AI).")

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
        "Tải lên File Excel (.xlsx):", 
        type=["xlsx"]
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
# HÀM DỊCH BẰNG THƯ VIỆN DEEP-TRANSLATOR
# ============================================================
def translate_texts_with_library(text_list, mode):
    """
    Dịch toàn bộ danh sách chuỗi bằng GoogleTranslator trong deep-translator
    """
    if not HAS_TRANSLATOR:
        st.error("❌ Chưa cài đặt 'deep-translator'. Vui lòng thêm 'deep-translator' vào file requirements.txt!")
        return {item: item for item in text_list}

    src_code = "zh-CN" if mode == "Trung ➔ Việt" else "vi"
    tgt_code = "vi" if mode == "Trung ➔ Việt" else "zh-CN"

    translator = GoogleTranslator(source=src_code, target=tgt_code)
    translated_dict = {}

    progress_bar = st.progress(0)
    total = len(text_list)

    for idx, item in enumerate(text_list):
        try:
            if item.strip():
                res = translator.translate(item)
                translated_dict[item] = res
            else:
                translated_dict[item] = item
        except Exception as e:
            translated_dict[item] = item
        
        # Cập nhật tiến trình
        progress_bar.progress((idx + 1) / total)
        time.sleep(0.05) # Tránh bị rate-limit

    progress_bar.empty()
    return translated_dict

# ============================================================
# 2. XỬ LÝ DỊCH CHÍNH
# ============================================================
if uploaded_file is not None:
    if st.button(f"🚀 Dịch ({translation_mode}) & Bảo Toàn Format Excel", use_container_width=True):
        try:
            with st.spinner(f"1️⃣ Đang quét các ô cần dịch theo chế độ [{translation_mode}]..."):
                file_bytes = uploaded_file.read()
                wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

                texts_to_translate = set()
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows():
                        for cell in row:
                            if cell.value and isinstance(cell.value, str):
                                val = cell.value.strip()
                                if val.startswith("="): # Bỏ qua công thức
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
                st.info(f"2️⃣ Đang dịch {len(unique_texts)} đoạn văn bản bằng thư viện dịch chuyên dụng...")
                translation_dict = translate_texts_with_library(unique_texts, translation_mode)

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

        except Exception as e:
            st.error(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
