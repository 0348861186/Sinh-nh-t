import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from deep_translator import GoogleTranslator
import easyocr
from PIL import Image
import numpy as np
import io
import copy

# Cấu hình giao diện Streamlit Cloud
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🇨🇳 ĐỒNG BỘ DỊCH SONG NGỮ TRUNG - VIỆT 🇻🇳")

# --- DASHBOARD ĐIỀU KHIỂN (YÊU CẦU 4 & 5) ---
st.sidebar.header("CÀI ĐẶT CẤU HÌNH")

# Yêu cầu 4: Chọn định dạng file load lên là file ảnh hoặc file excel
file_type = st.sidebar.radio(
    "1. Chọn định dạng file đầu vào:",
    options=["File Excel (.xlsx)", "File Ảnh (.jpg, .png)"]
)

# Yêu cầu 5: Chọn chế độ dịch trung việt hoặc việt trung
translation_mode = st.sidebar.selectbox(
    "2. Chọn chế độ dịch:",
    options=["Trung -> Việt", "Việt -> Trung"]
)

# Thiết lập ngôn ngữ
if translation_mode == "Trung -> Việt":
    src_lang = 'zh-CN'
    dest_lang = 'vi'
    ocr_langs = ['ch_sim', 'en']
else:
    src_lang = 'vi'
    dest_lang = 'zh-CN'
    ocr_langs = ['vi', 'en']

@st.cache_resource
def load_ocr(langs):
    return easyocr.Reader(list(langs))

try:
    reader = load_ocr(tuple(ocr_langs))
except Exception as e:
    st.sidebar.error(f"Lỗi khởi tạo OCR: {str(e)}")
    reader = None

uploaded_file = None
if file_type == "File Excel (.xlsx)":
    uploaded_file = st.file_uploader("Tải lên file Excel gốc", type=["xlsx"])
else:
    uploaded_file = st.file_uploader("Tải lên file Ảnh gốc", type=["jpg", "png", "jpeg"])

# --- HÀM DỊCH THUẬT CHUYÊN DỤNG CHÍNH XÁC (YÊU CẦU 7) ---

def translate_text(text, src, dest):
    """Hàm dịch chuyên dụng, giữ nguyên cấu trúc dòng để tránh dịch sai bét"""
    if not text or str(text).strip() == "":
        return text
    if isinstance(text, (int, float)):
        return text
    try:
        # Làm sạch chuỗi nhưng giữ ngữ cảnh để dịch chính xác nhất
        query_text = str(text).strip()
        translated = GoogleTranslator(source=src, target=dest).translate(query_text)
        return translated
    except Exception as e:
        return f"[Lỗi dịch: {str(e)}]"

# --- XỬ LÝ QUY TRÌNH FILE (YÊU CẦU 1, 2, 3, 6) ---

def process_excel_pure_format(file_bytes, src, dest):
    """
    XỬ LÝ EXCEL GIỮ NGUYÊN ĐỊNH DẠNG TUYỆT ĐỐI (YÊU CẦU 1)
    Đọc trực tiếp từ file gốc, chỉ ghi đè text song ngữ vào ô chứa chữ
    """
    # Load workbook gốc đầy đủ định dạng (bật keep_vba nếu có macro, giữ nguyên mọi thứ)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False, keep_vba=True)
    
    for sheet in wb.worksheets:
        # Giữ nguyên độ rộng cột và độ cao dòng của sheet gốc
        
        for row in sheet.iter_rows():
            for cell in row:
                # Chỉ dịch các ô có giá trị và giá trị đó là chuỗi ký tự (Văn bản)
                if cell.value is not None and isinstance(cell.value, str):
                    original_text = cell.value.strip()
                    
                    # Bỏ qua nếu ô là công thức (bắt đầu bằng dấu =)
                    if original_text.startswith('='):
                        continue
                        
                    if original_text:
                        # Tiến hành dịch chuyên dụng
                        translated_text = translate_text(original_text, src, dest)
                        
                        # YÊU CẦU 2 & 3: Dòng dịch nằm ngay dưới dòng gốc, chung một ô
                        cell.value = f"{original_text}\n{translated_text}"
                        
                        # Bảo toàn định dạng căn lề cũ, chỉ bật thêm wrap_text để xuống dòng đẹp
                        if cell.alignment:
                            cell.alignment = Alignment(
                                horizontal=cell.alignment.horizontal,
                                vertical="top", # Ép lên top để text không bị che khi dòng cao
                                wrap_text=True,
                                shrink_to_fit=cell.alignment.shrink_to_fit,
                                indent=cell.alignment.indent,
                                text_rotation=cell.alignment.text_rotation
                            )
                        else:
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
                        
                        # TẤT CẢ các thuộc tính font, màu nền (fill), viền (border), số định dạng (number_format) 
                        # của ô tính KHÔNG HỀ BỊ THAY ĐỔI vì chúng ta không khởi tạo lại ô.
                        
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

def process_image_to_excel(image_bytes, src, dest):
    """
    XỬ LÝ ẢNH -> XUẤT RA FILE EXCEL SONG NGỮ
    """
    if reader is None:
        return None
        
    image = Image.open(io.BytesIO(image_bytes)).convert('RGB')
    img_np = np.array(image)
    results = reader.readtext(img_np)
    
    # Tạo file Excel mới cho đầu ra của ảnh
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KetQuaDichAnh"
    
    # Định dạng tiêu đề cột đầu ra gọn gàng
    ws['A1'] = "Nội dung song ngữ (Dòng gốc / Dòng dịch)"
    ws['A1'].font = openpyxl.styles.Font(bold=True, size=12)
    ws.column_dimensions['A'].width = 70
    
    row_idx = 2
    for (bbox, text, prob) in results:
        clean_text = text.strip()
        if clean_text:
            translated_text = translate_text(clean_text, src, dest)
            cell = ws.cell(row=row_idx, column=1)
            
            # Yêu cầu 2 & 3: Dòng chữ dịch nằm ngay bên dưới chữ gốc trong cùng 1 ô
            cell.value = f"{clean_text}\n{translated_text}"
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            row_idx += 1
            
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

# --- ĐIỀU KHIỂN NÚT NHẤN (YÊU CẦU 6) ---

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    st.success("Tải file lên thành công! Sẵn sàng dịch thuật.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Nút nhấn dịch
        if st.button("🚀 BẮT ĐẦU DỊCH", use_container_width=True):
            with st.spinner("Hệ thống AI đang dịch thuật..."):
                try:
                    if file_type == "File Excel (.xlsx)":
                        output_data = process_excel_pure_format(file_bytes, src_lang, dest_lang)
                    else:
                        output_data = process_image_to_excel(file_bytes, src_lang, dest_lang)
                    
                    if output_data is not None:
                        st.session_state['translated_output'] = output_data
                        st.success("🎉 Đã dịch xong toàn bộ nội dung!")
                except Exception as e:
                    st.error(f"Lỗi xử lý file: {str(e)}")

    with col2:
        # Nút download file excel sau dịch (Đáp ứng việc xuất ngược lại file Excel dù đầu vào là gì)
        if 'translated_output' in st.session_state:
            orig_name = uploaded_file.name.split('.')[0]
            st.download_button(
                label="📥 TẢI FILE EXCEL SONG NGỮ",
                data=st.session_state['translated_output'],
                file_name=f"SongNgu_{orig_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
else:
    st.info("Vui lòng tải file của bạn lên ở thanh menu bên trái để bắt đầu.")
