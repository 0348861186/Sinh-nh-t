import io
import os
import copy
from pathlib import Path
from PIL import Image
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# ============================================================
# CẤU HÌNH TRANG STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Phần mềm Dịch Song Ngữ Trung - Việt",
    layout="wide"
)

st.title("🈲 🇻🇳 Phần mềm Dịch Song Ngữ Trung - Việt")

st.markdown("""
- **Excel:** Giữ nguyên workbook gốc, số sheet, số dòng, số cột, merge cell, độ rộng cột, font, màu, border. Chỉ thay nội dung các ô được dịch.
- **Ảnh:** OCR và tạo Excel chuẩn mẫu.
""")

# ============================================================
# KIỂM TRA EASYOCR
# ============================================================
try:
    import easyocr
    @st.cache_resource
    def load_reader():
        return easyocr.Reader(["ch_sim", "en"])
    reader = load_reader()
    has_easyocr = True
except Exception:
    has_easyocr = False

# ============================================================
# KHỞI TẠO BỘ DỊCH
# ============================================================
translator = GoogleTranslator(source="zh-CN", target="vi")

def translate_text(text):
    if text is None:
        return ""
    text = str(text)
    if text.strip() == "":
        return ""
    try:
        if text.strip().isdigit():
            return text
        return translator.translate(text)
    except Exception:
        return text

def should_translate(value):
    if value is None or not isinstance(value, str):
        return False
    text = value.strip()
    if not text or text.startswith("=") or text.isdigit():
        return False
    
    has_chinese = any(
        "\u3400" <= ch <= "\u4dbf"
        or "\u4e00" <= ch <= "\u9fff"
        or "\uf900" <= ch <= "\ufaff"
        for ch in text
    )
    return has_chinese

def translate_cell_value(cell):
    value = cell.value
    if not should_translate(value):
        return

    original = value
    if "\n" in original:
        first_line = original.splitlines()[0].strip()
        if any(
            "\u3400" <= ch <= "\u4dbf"
            or "\u4e00" <= ch <= "\u9fff"
            or "\uf900" <= ch <= "\ufaff"
            for ch in first_line
        ):
            return

    translated = translate_text(original)
    if not translated or translated == original:
        return

    cell.value = f"{original}\n{translated}"
    old_alignment = copy.copy(cell.alignment)
    cell.alignment = Alignment(
        horizontal=old_alignment.horizontal,
        vertical=old_alignment.vertical,
        text_rotation=old_alignment.text_rotation,
        wrap_text=True,
        shrink_to_fit=old_alignment.shrink_to_fit,
        indent=old_alignment.indent,
        relativeIndent=old_alignment.relativeIndent,
        justifyLastLine=old_alignment.justifyLastLine,
        readingOrder=old_alignment.readingOrder
    )

def process_excel(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=False)
    original_sheetnames = list(wb.sheetnames)
    original_dimensions = {}
    original_column_widths = {}
    original_row_heights = {}
    original_merges = {}

    for ws in wb.worksheets:
        original_dimensions[ws.title] = (ws.max_row, ws.max_column)
        original_merges[ws.title] = [str(rng) for rng in ws.merged_cells.ranges]
        original_column_widths[ws.title] = {key: dim.width for key, dim in ws.column_dimensions.items()}
        original_row_heights[ws.title] = {key: dim.height for key, dim in ws.row_dimensions.items()}

    translated_count = 0
    for ws in wb.worksheets:
        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.__class__.__name__ == "MergedCell":
                    continue
                before = cell.value
                translate_cell_value(cell)
                if cell.value != before:
                    translated_count += 1

    if wb.sheetnames != original_sheetnames:
        raise RuntimeError("Cấu trúc sheet bị thay đổi ngoài ý muốn.")

    for ws in wb.worksheets:
        old_rows, old_cols = original_dimensions[ws.title]
        if ws.max_row != old_rows or ws.max_column != old_cols:
            raise RuntimeError(f"Kích thước sheet '{ws.title}' bị thay đổi.")

        for key, old_width in original_column_widths[ws.title].items():
            if ws.column_dimensions[key].width != old_width:
                ws.column_dimensions[key].width = old_width

        for key, old_height in original_row_heights[ws.title].items():
            if ws.row_dimensions[key].height != old_height:
                ws.row_dimensions[key].height = old_height

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, wb, translated_count

# ============================================================
# GIAO DIỆN CHÍNH (UPLOAD FILE)
# ============================================================
st.divider()
file_type = st.radio("📂 Chọn định dạng file đầu vào:", ("File Excel (.xlsx)", "File Hình Ảnh (.png, .jpg, .jpeg)"))

uploaded_file = st.file_uploader(
    "📥 Tải file lên tại đây (Kéo thả hoặc bấm nút Browse files):", 
    type=["xlsx", "png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    if "Excel" in file_type:
        st.success(f"Đã nhận file Excel: **{uploaded_file.name}**")
        try:
            preview_wb = openpyxl.load_workbook(uploaded_file, data_only=False)
            preview_ws = preview_wb.active
            
            with st.expander("👁️ Xem trước dữ liệu Excel gốc"):
                df_preview = pd.DataFrame(preview_ws.values)
                st.dataframe(df_preview.head(10), use_container_width=True)

            if st.button("🚀 Bắt đầu dịch và giữ nguyên 100% cấu trúc Excel"):
                with st.spinner("Đang dịch toàn bộ workbook..."):
                    uploaded_file.seek(0)
                    output, result_wb, translated_count = process_excel(uploaded_file)
                    st.success("✅ Dịch hoàn tất thành công!")
                    st.info(f"🔤 Số ô đã dịch: **{translated_count}**")

                    st.download_button(
                        label="📥 Tải xuống File Excel đã dịch",
                        data=output.getvalue(),
                        file_name="translated_" + uploaded_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Lỗi xử lý file Excel: {e}")
