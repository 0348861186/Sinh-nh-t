from pathlib import Path

code = r'''import io
import os
import copy
from PIL import Image
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


# ============================================================
# EASYOCR
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
# STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Phần mềm Dịch Song Ngữ Trung - Việt",
    layout="wide"
)

st.title("🈲 🇻🇳 Phần mềm Dịch Song Ngữ Trung - Việt")

st.markdown("""
**Excel:** Giữ nguyên workbook gốc, số sheet, số dòng, số cột,
merge cell, độ rộng cột, chiều cao dòng, font, màu, border,
alignment và các thuộc tính bố cục có trong file. Chỉ thay nội dung
các ô được dịch.

**Ảnh:** OCR và tạo Excel theo logic ảnh hiện tại.
""")


# ============================================================
# TRANSLATOR
# ============================================================
translator = GoogleTranslator(
    source="zh-CN",
    target="vi"
)


def translate_text(text):
    if text is None:
        return ""

    text = str(text)

    if text.strip() == "":
        return ""

    try:
        # Giữ nguyên số
        if text.strip().isdigit():
            return text

        return translator.translate(text)

    except Exception:
        # Không làm mất dữ liệu nếu dịch lỗi
        return text


# ============================================================
# HÀM KIỂM TRA CÓ NÊN DỊCH KHÔNG
# ============================================================
def should_translate(value):
    if value is None:
        return False

    if not isinstance(value, str):
        return False

    text = value.strip()

    if not text:
        return False

    # Không dịch công thức Excel
    if text.startswith("="):
        return False

    # Nếu chỉ toàn số thì không dịch
    if text.isdigit():
        return False

    # Kiểm tra có ký tự Trung hay không.
    # Nếu không có tiếng Trung thì giữ nguyên.
    has_chinese = any(
        "\u3400" <= ch <= "\u4dbf"
        or "\u4e00" <= ch <= "\u9fff"
        or "\uf900" <= ch <= "\ufaff"
        for ch in text
    )

    return has_chinese


# ============================================================
# DỊCH 1 Ô - GIỮ NGUYÊN CẤU TRÚC
# ============================================================
def translate_cell_value(cell):
    value = cell.value

    if not should_translate(value):
        return

    original = value

    # Không dịch lại nếu ô đã có Trung + Việt bằng xuống dòng
    # và dòng đầu tiên vẫn là tiếng Trung.
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

    if not translated:
        return

    if translated == original:
        return

    # Chỉ thay VALUE của ô.
    # Không thay font, border, fill, width, height, merge...
    cell.value = f"{original}\n{translated}"

    # Sao lưu alignment hiện tại và chỉ bật wrap_text.
    old_alignment = copy.copy(cell.alignment)

    cell.alignment = copy.copy(old_alignment)
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


# ============================================================
# XỬ LÝ EXCEL
# ============================================================
def process_excel(uploaded_file):
    """
    QUAN TRỌNG:
    Không tạo workbook mới.
    Mở trực tiếp workbook gốc và sửa trên workbook đó.

    Không:
      - append()
      - insert_rows()
      - delete_rows()
      - insert_cols()
      - delete_cols()
      - tự resize row
      - tự resize column
      - tạo lại header
      - tạo lại bảng
      - bỏ sheet

    Mục tiêu:
      File kết quả = file gốc + nội dung dịch.
    """

    # --------------------------------------------------------
    # Mở workbook gốc
    # --------------------------------------------------------
    wb = openpyxl.load_workbook(
        uploaded_file,
        data_only=False
    )

    # --------------------------------------------------------
    # Lưu snapshot cấu trúc để kiểm tra sau xử lý
    # --------------------------------------------------------
    original_sheetnames = list(wb.sheetnames)

    original_dimensions = {}
    original_column_widths = {}
    original_row_heights = {}
    original_merges = {}

    for ws in wb.worksheets:

        original_dimensions[ws.title] = (
            ws.max_row,
            ws.max_column
        )

        original_merges[ws.title] = list(
            str(rng)
            for rng in ws.merged_cells.ranges
        )

        original_column_widths[ws.title] = {
            key: dim.width
            for key, dim in ws.column_dimensions.items()
        }

        original_row_heights[ws.title] = {
            key: dim.height
            for key, dim in ws.row_dimensions.items()
        }

    # --------------------------------------------------------
    # Duyệt TOÀN BỘ workbook, không chỉ active sheet
    # --------------------------------------------------------
    translated_count = 0

    for ws in wb.worksheets:

        # Duyệt đúng vùng hiện hữu của sheet.
        # Không tạo thêm dòng/cột.
        max_row = ws.max_row
        max_column = ws.max_column

        for row_idx in range(1, max_row + 1):

            for col_idx in range(1, max_column + 1):

                cell = ws.cell(
                    row=row_idx,
                    column=col_idx
                )

                # Bỏ qua MergedCell
                # Chỉ ô neo của merged range mới chứa value.
                if cell.__class__.__name__ == "MergedCell":
                    continue

                before = cell.value

                translate_cell_value(cell)

                if cell.value != before:
                    translated_count += 1

    # --------------------------------------------------------
    # KHÔNG chỉnh row_dimensions / column_dimensions
    # --------------------------------------------------------
    # Cố ý không có code kiểu:
    #
    # ws.row_dimensions[x].height = ...
    # ws.column_dimensions[x].width = ...
    #
    # vì phải giữ nguyên file gốc.

    # --------------------------------------------------------
    # Kiểm tra workbook sau khi dịch
    # --------------------------------------------------------
    if wb.sheetnames != original_sheetnames:
        raise RuntimeError(
            "Cấu trúc sheet bị thay đổi ngoài ý muốn."
        )

    for ws in wb.worksheets:

        old_rows, old_cols = original_dimensions[ws.title]

        if ws.max_row != old_rows:
            raise RuntimeError(
                f"Sheet '{ws.title}' bị thay đổi số dòng: "
                f"{old_rows} -> {ws.max_row}"
            )

        if ws.max_column != old_cols:
            raise RuntimeError(
                f"Sheet '{ws.title}' bị thay đổi số cột: "
                f"{old_cols} -> {ws.max_column}"
            )

        new_merges = [
            str(rng)
            for rng in ws.merged_cells.ranges
        ]

        if new_merges != original_merges[ws.title]:
            raise RuntimeError(
                f"Merge cell của sheet '{ws.title}' bị thay đổi."
            )

        # Kiểm tra width
        for key, old_width in original_column_widths[
            ws.title
        ].items():

            new_width = ws.column_dimensions[key].width

            if new_width != old_width:
                ws.column_dimensions[key].width = old_width

        # Kiểm tra height
        for key, old_height in original_row_heights[
            ws.title
        ].items():

            new_height = ws.row_dimensions[key].height

            if new_height != old_height:
                ws.row_dimensions[key].height = old_height

    # --------------------------------------------------------
    # Xuất workbook
    # --------------------------------------------------------
    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output, wb, translated_count


# ============================================================
# CHỌN LOẠI FILE
# ============================================================
file_type = st.radio(
    "Chọn định dạng file đầu vào:",
    (
        "File Excel (.xlsx)",
        "File Hình Ảnh (.png, .jpg, .jpeg)"
    )
)


uploaded_file = st.file_uploader(
    "Tải file lên tại đây:",
    type=[
        "xlsx",
        "png",
        "jpg",
        "jpeg"
    ]
)


# ============================================================
# FILE ĐÃ UPLOAD
# ============================================================
if uploaded_file is not None:

    # ========================================================
    # EXCEL
    # ========================================================
    if "Excel" in file_type:

        st.success(
            "Đã tải lên file Excel thành công!"
        )

        try:

            # Chỉ đọc để preview.
            # Workbook xử lý chính sẽ được mở lại khi bấm dịch.
            preview_wb = openpyxl.load_workbook(
                uploaded_file,
                data_only=False
            )

            preview_ws = preview_wb.active

            st.write("Xem trước dữ liệu Excel gốc:")

            df_preview = pd.DataFrame(
                preview_ws.values
            )

            st.dataframe(
                df_preview.head(10),
                use_container_width=True
            )

            # ------------------------------------------------
            # Hiển thị thông tin workbook
            # ------------------------------------------------
            st.info(
                f"📁 File: **{uploaded_file.name}**  \n"
                f"📄 Số sheet: **{len(preview_wb.worksheets)}**  \n"
                f"📐 Sheet đang xem: "
                f"**{preview_ws.max_row} dòng × "
                f"{preview_ws.max_column} cột**"
            )

            st.write(
                "**Các sheet:** "
                + ", ".join(preview_wb.sheetnames)
            )

            # ------------------------------------------------
            # NÚT DỊCH
            # ------------------------------------------------
            if st.button(
                "🚀 Bắt đầu dịch và giữ nguyên 100% cấu trúc Excel"
            ):

                with st.spinner(
                    "Đang dịch toàn bộ workbook, giữ nguyên cấu trúc gốc..."
                ):

                    # Reset con trỏ file
                    uploaded_file.seek(0)

                    output, result_wb, translated_count = process_excel(
                        uploaded_file
                    )

                    st.success(
                        "✅ Dịch hoàn tất - đã giữ nguyên cấu trúc workbook!"
                    )

                    st.info(
                        f"🔤 Số ô đã dịch: **{translated_count}**"
                    )

                    st.info(
                        f"📄 Số sheet: **{len(result_wb.worksheets)}**"
                    )

                    # Hiển thị kiểm tra từng sheet
                    check_text = []

                    for ws in result_wb.worksheets:
                        check_text.append(
                            f"• **{ws.title}**: "
                            f"{ws.max_row} dòng × "
                            f"{ws.max_column} cột"
                        )

                    st.markdown(
                        "\n".join(check_text)
                    )

                    st.download_button(
                        label="📥 Tải xuống File Excel đã dịch",
                        data=output.getvalue(),
                        file_name=(
                            "translated_"
                            + os.path.splitext(
                                uploaded_file.name
                            )[0]
                            + ".xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        )
                    )

        except Exception as e:

            st.error(
                f"Lỗi xử lý file Excel: {e}"
            )


    # ========================================================
    # HÌNH ẢNH
    # ========================================================
    else:

        st.success(
            "Đã tải lên hình ảnh thành công!"
        )

        image = Image.open(
            uploaded_file
        )

        st.image(
            image,
            caption="Ảnh gốc tải lên",
            use_container_width=True
        )

        if not has_easyocr:

            st.error(
                "Thư viện EasyOCR chưa được cấu hình."
            )

        else:

            if st.button(
                "🚀 Bắt đầu dịch ảnh sang Excel chuẩn mẫu"
            ):

                with st.spinner(
                    "Đang xử lý ảnh, nhận diện bảng và dịch thuật..."
                ):

                    img_np = np.array(image)

                    results = reader.readtext(
                        img_np
                    )

                    # ------------------------------------------------
                    # Tạo workbook ảnh
                    # ------------------------------------------------
                    wb_img = openpyxl.Workbook()

                    ws_img = wb_img.active

                    ws_img.title = "Translated Table"

                    # ------------------------------------------------
                    # Header mẫu - GIỮ NGUYÊN LOGIC CŨ
                    # ------------------------------------------------
                    headers = [
                        "STT",
                        "部分\nBộ phận",
                        "开几台机\nSố máy mở",
                        "正式工\nChính thức",
                        "临时工\nThời vụ",
                        "备注\nGhi chú"
                    ]

                    ws_img.append(headers)

                    header_fill = PatternFill(
                        start_color="ED7D31",
                        end_color="ED7D31",
                        fill_type="solid"
                    )

                    header_font = Font(
                        bold=True,
                        color="FFFFFF",
                        size=11
                    )

                    thin_border = Border(
                        left=Side(
                            style="thin",
                            color="BFBFBF"
                        ),
                        right=Side(
                            style="thin",
                            color="BFBFBF"
                        ),
                        top=Side(
                            style="thin",
                            color="BFBFBF"
                        ),
                        bottom=Side(
                            style="thin",
                            color="BFBFBF"
                        )
                    )

                    ws_img.row_dimensions[1].height = 30

                    for col_num in range(
                        1,
                        len(headers) + 1
                    ):

                        cell = ws_img.cell(
                            row=1,
                            column=col_num
                        )

                        cell.fill = header_fill
                        cell.font = header_font

                        cell.alignment = Alignment(
                            wrap_text=True,
                            vertical="center",
                            horizontal="center"
                        )

                        cell.border = thin_border

                    # ------------------------------------------------
                    # Dữ liệu mẫu - GIỮ NGUYÊN LOGIC CŨ
                    # ------------------------------------------------
                    sample_rows = [
                        ("1", "连机", "5", "3", "2", ""),
                        ("2", "制袋机", "6", "3", "2", ""),
                        ("3", "连机吹膜", "5", "4", "", ""),
                        ("4", "制袋机吹膜", "4", "2", "1", "")
                    ]

                    for r_idx, r_data in enumerate(
                        sample_rows,
                        start=2
                    ):

                        ws_img.row_dimensions[
                            r_idx
                        ].height = 40

                        for c_idx, val in enumerate(
                            r_data,
                            start=1
                        ):

                            cell = ws_img.cell(
                                row=r_idx,
                                column=c_idx
                            )

                            if (
                                c_idx == 2
                                and val.strip() != ""
                            ):

                                translated_val = translate_text(
                                    val
                                )

                                cell.value = (
                                    f"{val}\n"
                                    f"{translated_val}"
                                )

                            else:
                                cell.value = val

                            cell.alignment = Alignment(
                                wrap_text=True,
                                vertical="center",
                                horizontal="center"
                            )

                            cell.border = thin_border

                            cell.font = Font(
                                size=11
                            )

                    # ------------------------------------------------
                    # Độ rộng cột ảnh - GIỮ NGUYÊN LOGIC CŨ
                    # ------------------------------------------------
                    for col in ws_img.columns:

                        col_letter = (
                            openpyxl.utils.get_column_letter(
                                col[0].column
                            )
                        )

                        ws_img.column_dimensions[
                            col_letter
                        ].width = 18

                    # ------------------------------------------------
                    # Xuất
                    # ------------------------------------------------
                    output_img = io.BytesIO()

                    wb_img.save(
                        output_img
                    )

                    output_img.seek(0)

                    st.success(
                        "Đã chuyển đổi ảnh thành công sang Excel!"
                    )

                    st.download_button(
                        label="📥 Tải xuống File Excel kết quả",
                        data=output_img.getvalue(),
                        file_name="translated_table_from_image.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        )
                    )
'''

path = Path("/mnt/data/main_fixed.py")
path.write_text(code, encoding="utf-8")
print(f"Đã tạo: {path}")
print(f"Số dòng code: {len(code.splitlines())}")
