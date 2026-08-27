import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from deep_translator import GoogleTranslator
import easyocr
import cv2
import numpy as np
import io

# Cấu hình trang Streamlit
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🇨🇳 ĐỒNG BỘ DỊCH SONG NGỮ TRUNG - VIỆT 🇻🇳")

# Khởi tạo EasyOCR (Sử dụng bộ nhớ cache để không bị load lại mỗi lần nhấn nút)
@st.cache_resource
def load_ocr():
    # 'ch_sim': Tiếng Trung giản thể, 'vi': Tiếng Việt, 'en': Tiếng Anh phụ trợ
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

# Thiết lập mã ngôn ngữ chuẩn cho deep-translator
src_lang = 'zh-CN' if translation_mode == "Trung -> Việt" else 'vi'
dest_lang = 'vi' if translation_mode == "Trung -> Việt" else 'zh-CN'

# Cổng upload file tương ứng
uploaded_file = None
if file_type == "File Excel (.xlsx)":
    uploaded_file = st.file_uploader("Tải lên file Excel gốc", type=["xlsx"])
else:
    uploaded_file = st.file_uploader("Tải lên file Ảnh gốc", type=["jpg", "png", "jpeg"])

# --- HÀM XỬ LÝ CHỨC NĂNG CHUYÊN DỤNG ---

def translate_text(text, src, dest):
    """Hàm dịch chuyên dụng độ chính xác cao sử dụng Deep Translator"""
    if not text or str(text).strip() == "":
        return text
    # Bỏ qua nếu dữ liệu chỉ là số (int, float) nhằm giữ nguyên định dạng dữ liệu số gốc
    if isinstance(text, (int, float)):
        return text
        
    try:
        translated = GoogleTranslator(source=src, target=dest).translate(str(text))
        return translated
    except Exception as e:
        return f"[Lỗi dịch: {str(e)}]"

def process_excel(file_bytes, src, dest):
    """Xử lý Excel: Đọc từng ô, giữ nguyên định dạng, dịch chung ô, xuống dòng (Yêu cầu 1, 2, 3)"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                # Chỉ xử lý các ô chứa văn bản (chuỗi ký tự)
                if cell.value is not None and isinstance(cell.value, str):
                    original_text = cell.value.strip()
                    if original_text:
                        # Thực hiện dịch nội dung văn bản
                        translated_text = translate_text(original_text, src, dest)
                        
                        # Yêu cầu 2 & 3: Nội dung dịch nằm chung ô, dòng đích nằm NGAY DƯỚI dòng gốc
                        cell.value = f"{original_text}\n{translated_text}"
                        
                        # Kích hoạt wrap_text để Excel hiển thị xuống dòng đẹp mắt ngay trong ô đó
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
                        
                        # Yêu cầu 1: Mọi thuộc tính font, background, border, độ rộng cột cũ đều được tự động giữ lại
    
    # Xuất file ra bộ nhớ tạm dưới dạng bytes
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

def process_image_to_excel(image_bytes, src, dest):
    """Xử lý Ảnh: Đọc chữ qua OCR, dịch và xuất ra file Excel định dạng cấu trúc song ngữ"""
    # Giải mã dữ liệu ảnh bằng OpenCV
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    
    # Nhận diện chữ bằng EasyOCR
    results = reader.readtext(img)
    
    # Tạo một file Excel mới để nhận cấu trúc kết quả trực quan
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KetQuaDich"
    
    # Thiết lập cột tiêu đề ban đầu
    ws['A1'] = "Nội dung song ngữ (Gốc \n Dịch)"
    ws['A1'].font = openpyxl.styles.Font(bold=True)
    ws.column_dimensions['A'].width = 60
    
    row_idx = 2
    for (bbox, text, prob) in results:
        clean_text = text.strip()
        if clean_text:
            translated_text = translate_text(clean_text, src, dest)
            
            # Ghi dữ liệu song ngữ vào chung ô
            cell = ws.cell(row=row_idx, column=1)
            cell.value = f"{clean_text}\n{translated_text}"
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            row_idx += 1
            
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

# --- KHỞI CHẠY TIẾN TRÌNH DỊCH VÀ TẢI FILE ---

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    st.success("Tải file lên thành công! Sẵn sàng để dịch.")
    
    # Tạo 2 cột để đặt nút bấm trực quan
    col1, col2 = st.columns(2)
    
    with col1:
        # Yêu cầu 6: Nút nhấn dịch
        if st.button("🚀 BẮT ĐẦU DỊCH", use_container_width=True):
            with st.spinner("Hệ thống AI đang dịch thuật... Vui lòng đợi trong giây lát..."):
                try:
                    if file_type == "File Excel (.xlsx)":
                        output_data = process_excel(file_bytes, src_lang, dest_lang)
                    else:
                        output_data = process_image_to_excel(file_bytes, src_lang, dest_lang)
                    
                    # Lưu kết quả dịch vào bộ nhớ trạng thái session
                    st.session_state['translated_output'] = output_data
                    st.success("🎉 Đã dịch xong toàn bộ nội dung!")
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")

    with col2:
        # Yêu cầu 6: Nút download file excel sau dịch
        if 'translated_output' in st.session_state:
            # Định dạng lại tên file xuất ra dựa trên file gốc
            original_name = uploaded_file.name.split('.')[0]
            st.download_button(
                label="📥 TẢI FILE EXCEL SONG NGỮ",
                data=st.session_state['translated_output'],
                file_name=f"SongNgu_{original_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
else:
    st.info("Vui lòng tải file của bạn lên ở thanh menu bên trái để bắt đầu thực
