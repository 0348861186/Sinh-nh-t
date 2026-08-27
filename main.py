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

# Khởi tạo OCR để đọc chữ từ ảnh
@st.cache_resource
def get_ocr_reader():
    import easyocr
    return easyocr.Reader(['ch_sim', 'en'])

def translate_text(text):
    if not text or not str(text).strip():
        return ""
    try:
        translator = get_translator()
        translated = translator.translate(str(text))
        return translated if translated else text
    except Exception as e:
        return text

def process_excel(file_bytes):
    """Xử lý file Excel: Giữ nguyên định dạng, thêm dòng tiếng Việt ngay dưới ô tiếng Trung."""
    wb = load_workbook(io.BytesIO(file_bytes))
    
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    original_text = str(cell.value)
                    translated_text = translate_text(original_text)
                    # Gộp chung vào 1 ô, tiếng Việt nằm dưới tiếng Trung
                    cell.value = f"{original_text}\n{translated_text}"
                    cell.alignment = cell.alignment.copy(wrap_text=True)
                    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Giao diện Streamlit
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", page_icon="🇨🇳🇻🇳", layout="centered")

st.title("🇨🇳 ➔ 🇻🇳 Phần Mềm Dịch Song Ngữ Trung - Việt")
st.write("Hỗ trợ tải lên file Excel (giữ nguyên định dạng gốc) hoặc Hình ảnh để dịch.")

# Chọn loại file tải lên
file_type = st.radio("Chọn định dạng file bạn muốn xử lý:", ("File Excel (.xlsx)", "File Ảnh (Image)"))

uploaded_file = st.file_uploader(
    "Tải file lên tại đây:", 
    type=["xlsx"] if "Excel" in file_type else ["png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    if "Excel" in file_type:
        st.info("Đã nhận file Excel. Nhấn nút bên dưới để tiến hành dịch:")
        
        # NÚT BẤM BẮT ĐẦU DỊCH
        if st.button("🚀 Bắt đầu dịch file Excel", type="primary"):
            with st.spinner("Đang dịch toàn bộ file Excel, vui lòng đợi..."):
                try:
                    processed_excel = process_excel(uploaded_file.getvalue())
                    st.success("Dịch thành công!")
                    
                    # NÚT TẢI XUỐNG FILE KẾT QUẢ
                    st.download_button(
                        label="📥 Tải xuống File Excel Song Ngữ",
                        data=processed_excel,
                        file_name="song_ngu_trung_viet.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý file Excel: {e}")
                    
    else:
        st.info("Đã nhận file ảnh. Đang tiến hành đọc chữ và dịch...")
        try:
            image = Image.open(uploaded_file)
            st.image(image, caption="Ảnh gốc", use_container_width=True)
            
            with st.spinner("Đang quét OCR và dịch nội dung ảnh..."):
                img_np = np.array(image)
                reader = get_ocr_reader()
                results = reader.readtext(img_np)
                
                st.subheader("Kết quả dịch nội dung từ ảnh:")
                for bbox, text, prob in results:
                    trans = translate_text(text)
                    st.markdown(f"🇨🇳 **Trung:** {text}  \n🇻🇳 **Việt:** {trans}  \n---", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý ảnh: {e}")
