import io
import os
import tempfile
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
from openpyxl import load_workbook

# Kiểm tra thư viện OCR cho ảnh
try:
    import easyocr
    from PIL import Image

    HAS_OCR = True
except ImportError:
    HAS_OCR = False

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Dịch Song Ngữ Trung - Việt / Việt - Trung",
    page_icon="🌐",
    layout="centered",
)

st.title("🌐 Ứng Dụng Dịch Song Ngữ Trung - Việt & Việt - Trung")
st.write(
    "Hỗ trợ dịch file Excel giữ nguyên định dạng hoặc dịch văn bản/hình ảnh có sẵn."
)

# 5) Dashboard có nút chọn chế độ dịch
translation_mode = st.radio(
    "Chọn chế độ dịch:", ("Trung ➔ Việt", "Việt ➔ Trung"), horizontal=True
)

# Khởi tạo bộ dịch chuyên dụng
if "Trung ➔ Việt" in translation_mode:
    source_lang, target_lang = "zh-CN", "vi"
else:
    source_lang, target_lang = "vi", "zh-CN"

translator = GoogleTranslator(source=source_lang, target=target_lang)


def translate_text(text):
    """Hàm dịch văn bản an toàn, giữ lại định dạng cơ bản"""
    if not text or not isinstance(text, str):
        return text
    if text.strip().isdigit() or len(text.strip()) <= 1:
        return text

    try:
        translated = translator.translate(text)
        if translated:
            # 2) Dòng tiếng Việt nằm ngay bên dưới dòng tiếng Trung (\n)
            # 3) Nội dung dịch nằm chung với ô được dịch
            return f"{text}\n{translated}"
    except Exception as e:
        print(f"Lỗi dịch: {e}")
    return text


# Khởi tạo state lưu trữ file Excel đã dịch trong bộ nhớ phiên làm việc
if "translated_file_bytes" not in st.session_state:
    st.session_state.translated_file_bytes = None
if "translated_file_name" not in st.session_state:
    st.session_state.translated_file_name = None

# 4) Dashboard có nút chọn file load lên (Ảnh hoặc Excel)
uploaded_file = st.file_uploader(
    "Tải lên file Excel (.xlsx) hoặc Hình ảnh (.png, .jpg, .jpeg)",
    type=["xlsx", "xls", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    file_extension = uploaded_file.name.split(".")[-1].lower()

    # XỬ LÝ FILE EXCEL
    if file_extension in ["xlsx", "xls"]:
        st.success(
            "Đã tải lên file Excel thành công! Vui lòng nhấn nút dịch bên dưới."
        )

        # 6) Dashboard có nút nhấn dịch và nút download file excel sau dịch (Độc lập rõ ràng)
        col1, col2 = st.columns(2)

        with col1:
            if st.button("🚀 Bắt đầu dịch file Excel"):
                with st.spinner("Đang xử lý và dịch tài liệu..."):
                    try:
                        bytes_data = uploaded_file.getvalue()
                        wb = load_workbook(io.BytesIO(bytes_data))

                        for sheet in wb.worksheets:
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value and isinstance(
                                        cell.value, str
                                    ):
                                        if not cell.value.startswith("="):
                                            cell.value = translate_text(
                                                cell.value
                                            )

                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        # Lưu vào session_state để nút download bên cạnh nhận diện được
                        st.session_state.translated_file_bytes = output.getvalue()
                        st.session_state.translated_file_name = (
                            f"translated_{uploaded_file.name}"
                        )

                        st.success(
                            "Dịch hoàn tất! Bạn có thể nhấn nút tải xuống bên cạnh."
                        )
                    except Exception as e:
                        st.error(f"Đã xảy ra lỗi khi xử lý file Excel: {e}")

        with col2:
            # Nút download file excel luôn hiển thị sẵn (sẽ kích hoạt khi đã dịch xong)
            if st.session_state.translated_file_bytes is not None:
                st.download_button(
                    label="📥 Tải xuống file Excel sau dịch",
                    data=st.session_state.translated_file_bytes,
                    file_name=st.session_state.translated_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("Nhấn 'Bắt đầu dịch' để mở khóa nút tải.")

    # XỬ LÝ FILE ẢNH
    elif file_extension in ["png", "jpg", "jpeg"]:
        st.image(uploaded_file, caption="Ảnh gốc đã tải lên")

        if not HAS_OCR:
            st.warning(
                "Thư viện OCR chưa được cài đặt. Vui lòng kiểm tra lại file requirements.txt."
            )
        else:
            if st.button("🚀 Nhận diện và dịch văn bản từ ảnh"):
                with st.spinner("Đang đọc chữ từ ảnh và dịch..."):
                    try:
                        image = Image.open(uploaded_file)
                        with tempfile.NamedTemporaryFile(
                            delete=False, suffix=".png"
                        ) as tmp:
                            image.save(tmp.name)
                            tmp_path = tmp.name

                        if "Trung ➔ Việt" in translation_mode:
                            ocr_langs = ["ch_sim", "en"]
                        else:
                            ocr_langs = ["vi", "en"]

                        reader = easyocr.Reader(ocr_langs)
                        results = reader.readtext(tmp_path)

                        st.write("### Kết quả dịch từ ảnh:")
                        for bbox, text, prob in results:
                            translated_line = translator.translate(text)
                            st.text(f"Gốc: {text}\nBản dịch: {translated_line}")

                        os.unlink(tmp_path)
                    except Exception as e:
                        st.error(f"Lỗi khi xử lý hình ảnh: {e}")
