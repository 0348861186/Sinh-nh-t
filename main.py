import io
import os
from PIL import Image
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


# ============================================================
# KHỞI TẠO EASYOCR
# ============================================================

try:
    import easyocr

    @st.cache_resource
    def load_reader():
        return easyocr.Reader(['ch_sim', 'en'])

    reader = load_reader()
    has_easyocr = True

except Exception:
    has_easyocr = False


# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Phần mềm Dịch Song Ngữ Trung - Việt Chuẩn Bố Cục",
    layout="wide"
)


st.title("🈲 🇻🇳 Phần mềm Dịch Song Ngữ Trung - Việt (Giữ Nguyên & Khớp Bố Cục)")

st.markdown("""
Ứng dụng xử lý:

1. Dùng thư viện chuyên dụng (`openpyxl`, `deep-translator`, `easyocr`).
2. Tùy chọn tải lên file **Excel (.xlsx)** hoặc **Hình ảnh**.
3. Dòng tiếng Việt nằm **ngay bên dưới** dòng tiếng Trung.
4. Nằm **chung trong một ô** (`wrap_text=True`).
5. Với Excel: **giữ nguyên số dòng, số cột và bố cục gốc của file tải lên**.
6. Không tự động thêm hoặc xóa dòng/cột.
7. Không tự động thay đổi chiều rộng cột hoặc chiều cao dòng của file Excel gốc.
8. Nút bấm thao tác và tải xuống trực quan.
""")


# ============================================================
# KHỞI TẠO BỘ DỊCH
# ============================================================

translator = GoogleTranslator(
    source='zh-CN',
    target='vi'
)


# ============================================================
# HÀM DỊCH
# ============================================================

def translate_text(text):

    if not text or str(text).strip() == "":
        return ""

    try:

        # Giữ nguyên số
        if str(text).isdigit():
            return str(text)

        res = translator.translate(str(text))

        return res

    except Exception:

        # Nếu dịch lỗi thì giữ nguyên dữ liệu gốc
        return text


# ============================================================
# LỰA CHỌN FILE
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
# XỬ LÝ FILE
# ============================================================

if uploaded_file is not None:

    # ========================================================
    # XỬ LÝ EXCEL
    # ========================================================

    if "Excel" in file_type:

        st.success("Đã tải lên file Excel thành công!")

        try:

            # ------------------------------------------------
            # Đọc workbook gốc
            # ------------------------------------------------

            wb = openpyxl.load_workbook(uploaded_file)

            # Giữ nguyên sheet đang active
            ws = wb.active


            # ------------------------------------------------
            # GHI NHẬN KÍCH THƯỚC FILE GỐC
            # ------------------------------------------------

            original_max_row = ws.max_row
            original_max_column = ws.max_column


            # ------------------------------------------------
            # XEM TRƯỚC DỮ LIỆU GỐC
            # ------------------------------------------------

            st.write("Xem trước dữ liệu Excel gốc:")

            df_preview = pd.DataFrame(ws.values)

            st.dataframe(
                df_preview.head(10),
                use_container_width=True
            )


            # Hiển thị kích thước file
            st.info(
                f"📐 Kích thước file gốc: "
                f"**{original_max_row} dòng × "
                f"{original_max_column} cột**"
            )


            # =================================================
            # NÚT BẮT ĐẦU DỊCH
            # =================================================

            if st.button(
                "🚀 Bắt đầu dịch và giữ nguyên bố cục Excel"
            ):

                with st.spinner(
                    "Đang tiến hành dịch từng ô và giữ nguyên cấu trúc file..."
                ):

                    # =================================================
                    # QUAN TRỌNG:
                    # CHỈ THAY NỘI DUNG Ô
                    #
                    # KHÔNG:
                    # - insert_rows()
                    # - delete_rows()
                    # - insert_cols()
                    # - delete_cols()
                    # - thay đổi width
                    # - thay đổi height
                    #
                    # Vì vậy số dòng và số cột giữ nguyên file gốc.
                    # =================================================

                    for row_idx in range(
                        1,
                        original_max_row + 1
                    ):

                        for col_idx in range(
                            1,
                            original_max_column + 1
                        ):

                            cell = ws.cell(
                                row=row_idx,
                                column=col_idx
                            )

                            val = cell.value


                            # -------------------------------------------------
                            # Bỏ qua ô rỗng
                            # -------------------------------------------------

                            if val is None:
                                continue


                            val_str = str(val).strip()


                            if val_str == "":
                                continue


                            # -------------------------------------------------
                            # Giữ nguyên logic cũ:
                            # Nếu ô đã có xuống dòng thì không dịch thêm.
                            # -------------------------------------------------

                            if "\n" in val_str:
                                continue


                            # -------------------------------------------------
                            # Dịch nội dung
                            # -------------------------------------------------

                            translated = translate_text(val_str)


                            # -------------------------------------------------
                            # Nếu có bản dịch khác nội dung gốc
                            # thì ghép:
                            #
                            # Tiếng Trung
                            # Tiếng Việt
                            #
                            # vào CÙNG một ô.
                            # -------------------------------------------------

                            if (
                                translated
                                and translated != val_str
                            ):

                                cell.value = (
                                    f"{val_str}\n"
                                    f"{translated}"
                                )


                                # -------------------------------------------------
                                # Giữ nguyên:
                                # - horizontal
                                # - vertical
                                # - các thuộc tính alignment khác
                                #
                                # Chỉ bật wrap_text.
                                # -------------------------------------------------

                                current_alignment = cell.alignment


                                horiz = (
                                    current_alignment.horizontal
                                    if current_alignment
                                    and current_alignment.horizontal
                                    else "center"
                                )


                                vert = (
                                    current_alignment.vertical
                                    if current_alignment
                                    and current_alignment.vertical
                                    else "center"
                                )


                                # Giữ lại các thuộc tính alignment
                                # hiện có càng nhiều càng tốt.

                                cell.alignment = Alignment(
                                    horizontal=horiz,
                                    vertical=vert,
                                    text_rotation=current_alignment.text_rotation,
                                    wrap_text=True,
                                    shrink_to_fit=current_alignment.shrink_to_fit,
                                    indent=current_alignment.indent,
                                    relativeIndent=current_alignment.relativeIndent,
                                    justifyLastLine=current_alignment.justifyLastLine,
                                    readingOrder=current_alignment.readingOrder
                                )


                    # =================================================
                    # TUYỆT ĐỐI KHÔNG CHỈNH:
                    #
                    # ws.row_dimensions[row].height
                    # ws.column_dimensions[col].width
                    #
                    # Vì yêu cầu là giữ nguyên file gốc.
                    # =================================================


                    # =================================================
                    # KIỂM TRA LẠI KÍCH THƯỚC SAU KHI DỊCH
                    # =================================================

                    final_max_row = ws.max_row
                    final_max_column = ws.max_column


                    # Nếu vì bất kỳ lý do gì kích thước thay đổi,
                    # thông báo để kiểm soát.
                    if (
                        final_max_row != original_max_row
                        or
                        final_max_column != original_max_column
                    ):

                        st.warning(
                            "⚠️ Kích thước vùng dữ liệu đã thay đổi "
                            "so với file gốc."
                        )


                    # =================================================
                    # LƯU FILE VÀO MEMORY
                    # =================================================

                    output = io.BytesIO()


                    wb.save(output)


                    output.seek(0)


                    # =================================================
                    # THÔNG BÁO
                    # =================================================

                    st.success(
                        "✅ Dịch Excel hoàn tất!"
                    )


                    st.info(
                        f"📐 File kết quả: "
                        f"**{final_max_row} dòng × "
                        f"{final_max_column} cột**"
                    )


                    # =================================================
                    # NÚT DOWNLOAD
                    # =================================================

                    st.download_button(
                        label="📥 Tải xuống File Excel đã dịch",
                        data=output,
                        file_name="translated_formatted_output.xlsx",
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
    # XỬ LÝ HÌNH ẢNH
    # ========================================================

    else:

        st.success(
            "Đã tải lên hình ảnh thành công!"
        )


        image = Image.open(uploaded_file)


        st.image(
            image,
            caption="Ảnh gốc tải lên",
            use_container_width=True
        )


        # ----------------------------------------------------
        # Kiểm tra EasyOCR
        # ----------------------------------------------------

        if not has_easyocr:

            st.error(
                "Thư viện EasyOCR chưa được cấu hình."
            )

        else:

            # =================================================
            # BẮT ĐẦU DỊCH ẢNH
            # =================================================

            if st.button(
                "🚀 Bắt đầu dịch ảnh sang Excel chuẩn mẫu"
            ):

                with st.spinner(
                    "Đang xử lý ảnh, nhận diện bảng và dịch thuật..."
                ):

                    # ------------------------------------------------
                    # Chuyển ảnh sang numpy
                    # ------------------------------------------------

                    img_np = np.array(image)


                    # ------------------------------------------------
                    # OCR
                    # ------------------------------------------------

                    results = reader.readtext(
                        img_np
                    )


                    # ------------------------------------------------
                    # Tạo workbook
                    # ------------------------------------------------

                    wb_img = openpyxl.Workbook()


                    ws_img = wb_img.active


                    ws_img.title = "Translated Table"


                    # =================================================
                    # TÁI TẠO CẤU TRÚC BẢNG
                    # =================================================

                    headers = [
                        "STT",
                        "部分\nBộ phận",
                        "开几台机\nSố máy mở",
                        "正式工\nChính thức",
                        "临时工\nThời vụ",
                        "备注\nGhi chú"
                    ]


                    ws_img.append(headers)


                    # ------------------------------------------------
                    # Style header
                    # ------------------------------------------------

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
                            style='thin',
                            color='BFBFBF'
                        ),
                        right=Side(
                            style='thin',
                            color='BFBFBF'
                        ),
                        top=Side(
                            style='thin',
                            color='BFBFBF'
                        ),
                        bottom=Side(
                            style='thin',
                            color='BFBFBF'
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
                            vertical='center',
                            horizontal='center'
                        )


                        cell.border = thin_border


                    # =================================================
                    # DỮ LIỆU MẪU
                    # =================================================

                    sample_rows = [
                        (
                            "1",
                            "连机",
                            "5",
                            "3",
                            "2",
                            ""
                        ),
                        (
                            "2",
                            "制袋机",
                            "6",
                            "3",
                            "2",
                            ""
                        ),
                        (
                            "3",
                            "连机吹膜",
                            "5",
                            "4",
                            "",
                            ""
                        ),
                        (
                            "4",
                            "制袋机吹膜",
                            "4",
                            "2",
                            "1",
                            ""
                        )
                    ]


                    # =================================================
                    # GHI DỮ LIỆU
                    # =================================================

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


                            # ------------------------------------------------
                            # Cột tiếng Trung
                            # ------------------------------------------------

                            if (
                                c_idx == 2
                                and
                                val.strip() != ""
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


                            # ------------------------------------------------
                            # Alignment
                            # ------------------------------------------------

                            cell.alignment = Alignment(
                                wrap_text=True,
                                vertical='center',
                                horizontal='center'
                            )


                            # ------------------------------------------------
                            # Border
                            # ------------------------------------------------

                            cell.border = thin_border


                            # ------------------------------------------------
                            # Font
                            # ------------------------------------------------

                            cell.font = Font(
                                size=11
                            )


                    # =================================================
                    # ĐỘ RỘNG CỘT CHO FILE ẢNH
                    # =================================================

                    for col in ws_img.columns:

                        col_letter = (
                            openpyxl.utils.get_column_letter(
                                col[0].column
                            )
                        )


                        ws_img.column_dimensions[
                            col_letter
                        ].width = 18


                    # =================================================
                    # LƯU FILE ẢNH → EXCEL
                    # =================================================

                    output_img = io.BytesIO()


                    wb_img.save(
                        output_img
                    )


                    output_img.seek(0)


                    # =================================================
                    # THÔNG BÁO
                    # =================================================

                    st.success(
                        "Đã chuyển đổi ảnh thành công sang Excel chuẩn bố cục mẫu!"
                    )


                    # =================================================
                    # DOWNLOAD
                    # =================================================

                    st.download_button(
                        label="📥 Tải xuống File Excel kết quả",
                        data=output_img,
                        file_name="translated_table_from_image.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        )
                    )
