import io
import cv2
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image
from openpyxl import load_workbook
from deep_translator import GoogleTranslator


# Khởi tạo bộ dịch tiếng Trung -> tiếng Việt chuyên dụng
@st.cache_resource
def get_translator():
    return GoogleTranslator(source="zh", target="vi")


# Khởi tạo OCR để đọc chữ từ ảnh (Hỗ trợ Tiếng Trung giản thể/phồn thể)
@st.cache_resource
def get_ocr_reader():
    import easyocr

    # 'ch_sim' dùng cho tiếng Trung giản thể, có thể đổi thành 'ch_tra' nếu cần
    return easyocr.Reader(["ch_sim", "en"])


def translate_text(text):
    if not text or not str(text).strip():
        return ""
    try:
        translator = get_translator()
        # Dịch đoạn văn bản
        translated = translator.translate(str(text))
        return translated if translated else text
    except Exception as e:
        return text


def process_excel(file_bytes):
    """Xử lý file Excel: Giữ nguyên định dạng, thêm dòng tiếng Việt ngay dưới ô tiếng Trung."""
    # Đọc workbook bằng openpyxl để giữ nguyên định dạng (màu sắc, font, border...)
    wb = load_workbook(io.BytesIO(file_bytes))

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    original_text = str(cell.value)
                    # Thực hiện dịch
                    translated_text = translate_text(original_text)

                    # Yêu cầu 3 & 4: Dòng tiếng Việt nằm ngay bên dưới dòng tiếng Trung trong cùng 1 ô
                    cell.value = f"{original_text}\n{translated_text}"

                    # Bật chế độ xuống dòng (Wrap text) để hiển thị đẹp mắt
                    cell.alignment = cell.alignment.copy(wrap_text=True)

    # Lưu ra buffer để tải về
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def process_image(uploaded_file):
    """Xử lý file ảnh: Nhận diện chữ tiếng Trung và dịch sang tiếng Việt."""
    image = Image.open(uploaded_file)
    img_np = np.array(image)

    reader = get_ocr_reader()
    results = reader.readtext(img_np)

    st.image(image, caption="Ảnh gốc", use_column_width=True)
    st.subheader("Kết quả dịch nội dung từ ảnh:")

    # Hiển thị kết quả dạng song ngữ
    translated_lines = []
    for bbox, text, prob in results:
        trans = translate_text(text)
        st.markdown(
            f"🇨🇳 **Trung:** {text}  \n🇻🇳 **Việt:** {trans}  \n---",
            unsafe_allow_html=True,
        )
        translated_lines.append(f"{text} -> {trans}")

    return translated_lines


# Giao diện Streamlit
st.set_page_config(
    page_title="Dịch Song Ngữ Trung - Việt", page_icon="🇨🇳🇻🇳", layout="centered"
)

st.title("🇨🇳 ➔ 🇻🇳 Phần Mềm Dịch Song Ngữ Trung - Việt")
st.write(
    "Hỗ trợ tải lên file Excel (giữ nguyên định dạng gốc) hoặc Hình ảnh để dịch."
)

# Chọn loại file tải lên
file_type = st.radio(
    "Chọn định dạng file bạn muốn xử lý:", ("File Excel (.xlsx)", "File Ảnh (Image)")
)

uploaded_file = st.file_uploader(
    "Tải file lên tại đây:",
    type=["xlsx"] if "Excel" in file_type else ["png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    if "Excel" in file_type:
        st.info("Đang xử lý file Excel, vui lòng đợi trong giây lát...")
        try:
            processed_excel = process_excel(uploaded_file.getvalue())
            st.success("Dịch thành công!")

            # Nút tải file Excel kết quả
            st.download_button(
                label="📥 Tải xuống File Excel Song Ngữ",
                data=processed_excel,
                file_name="song_ngu_trung_viet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý file Excel: {e}")

    else:
        st.info("Đang nhận diện chữ và dịch ảnh...")
        try:
            process_image(uploaded_file)
        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý ảnh: {e}")
