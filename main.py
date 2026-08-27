import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from googletrans import Translator
import easyocr
import cv2
import numpy as np
import io

# Cấu hình trang Streamlit
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🇨🇳 ĐỒNG BỘ DỊCH SONG NGỮ TRUNG - VIỆT 🇻🇳")

# Khởi tạo Translator và OCR Reader (hỗ trợ cả tiếng Trung và tiếng Việt)
translator = Translator()

@st.cache_resource
def load_ocr():
    # Khởi tạo EasyOCR cho tiếng Trung giản thể, phồn thể và tiếng Việt
    return easyocr.Reader(['ch_sim', 'en', 'vi'])

reader = load_ocr()

# --- GIAO DIỆN ĐIỀU KHIỂN (DASHBOARD) ---
st.sidebar.header("CÀI ĐẶT CẤU HÌNH")

# Yêu cầu 4: Chọn định dạng file load lên
file_type = st.sidebar.radio(
    "1. Chọn định dạng file đầu vào:",
    options=["File Excel (.xlsx)", "File Ảnh (.jpg, .png)"]
)

# Yêu cầu 5: Chọn chế độ dịch
translation_mode = st.sidebar.selectbox(
    "2. Chọn chế độ dịch:",
    options=["Trung -> Việt", "Việt -> Trung"]
)

src_lang = 'zh-cn' if translation_mode == "Trung -> Việt" else 'vi'
dest_lang = 'vi' if translation_mode == "Trung -> Việt" else 'zh-cn'

# Upload file dựa theo định dạng đã chọn
uploaded_file = None
if file_type == "File Excel (.xlsx)":
    uploaded_file = st.file_uploader("Tải lên file Excel gốc", type=["xlsx"])
else:
    uploaded_file = st.file_uploader("Tải lên file Ảnh gốc", type=["jpg", "png", "jpeg"])

# --- HÀM XỬ LÝ CHÍNH ---

def translate_text(text, src, dest):
    """Hàm dịch chuyên dụng sử dụng Googletrans"""
    if not text or str(text).strip() == "" or isinstance(text, (int, float)):
        return text
    try:
        translated = translator.translate(str(text), src=src, dest=dest)
        return translated.text
    except Exception as e:
        return f"[Lỗi dịch: {str(e)}]"

def process_excel(file_bytes, src, dest):
    """Xử lý dịch Excel: Giữ nguyên định dạng, dịch chung ô, xuống dòng (Yêu cầu 1, 2, 3)"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    
    for sheet in wb.worksheets:
        # Lặp qua từng ô có dữ liệu trong sheet
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and isinstance(cell.value, str):
                    original_text = cell.value.strip()
                    if original_text:
                        # Dịch nội dung
                        translated_text = translate_text(original_text, src, dest)
                        
                        # Yêu cầu 2 & 3: Nội dung dịch nằm chung ô, dòng đích nằm NGAY DƯỚI dòng gốc
                        # Cần kích hoạt chế độ wrap_text để Excel hiển thị xuống dòng trong ô
                        cell.value = f"{original_text}\n{translated_text}"
                        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
                        
                        # Yêu cầu 1: Định dạng (Font, Màu, Border...) được openpyxl giữ nguyên tự động từ file gốc
    
    # Xuất file ra bộ nhớ tạm
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

def process_image_to_excel(image_bytes, src, dest):
    """Xử lý Ảnh: Đọc chữ bằng OCR và xuất ra file Excel định dạng song ngữ"""
    # Đọc ảnh từ bytes
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    
    # Nhận diện chữ bằng EasyOCR
    results = reader.readtext(img)
    
    # Tạo một file Excel mới để lưu kết quả cấu trúc trực quan
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Result"
    
    # Thiết lập tiêu đề cột
    ws['A1'] = "Nội dung song ngữ (Gốc \n Dịch)"
    ws['A1'].font = openpyxl.styles.Font(bold=True)
    ws.column_dimensions['A'].width = 50
    
    row_idx = 2
    for (bbox, text, prob) in results:
        if text.strip():
            translated_text = translate_text(text, src, dest)
            # Chèn song ngữ vào chung ô, xuống dòng
            cell = ws.cell(row=row_idx, column=1)
            cell.value = f"{text}\n{translated_text}"
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
            row_idx += 1
            
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

# --- NÚT NHẤN ĐIỀU KHIỂN & KẾT QUẢ ---

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    st.success("Đã tải file lên thành công! Sẵn sàng dịch.")
    
    # Yêu cầu 6: Nút nhấn dịch
    if st.button("🚀 BẮT ĐẦU DỊCH"):
        with st.spinner("Hệ thống AI đang dịch thuật... Vui lòng đợi trong giây lát..."):
            try:
                if file_type == "File Excel (.xlsx)":
                    output_data = process_excel(file_bytes, src_lang, dest_lang)
                else:
                    output_data = process_image_to_excel(file_bytes, src_lang, dest_lang)
                
                st.session_state['translated_output'] = output_data
                st.success("🎉 Đã dịch xong toàn bộ nội dung!")
            except Exception as e:
                st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")

    # Yêu cầu 6: Nút download file excel sau khi dịch xong
    if 'translated_output' in st.session_state:
        st.download_button(
            label="📥 TẢI FILE EXCEL SAU DỊCH",
            data=st.session_state['translated_output'],
            file_name=f"SongNgu_{uploaded_file.name.split('.')[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Vui lòng upload file từ thanh menu bên trái để bắt đầu.")
