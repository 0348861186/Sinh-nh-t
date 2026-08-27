import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.cell.text import InlineFont, RichText
from openpyxl.utils import get_column_letter
from deep_translator import GoogleTranslator
import easyocr
from PIL import Image
import numpy as np
import io
import copy

# Cấu hình trang Streamlit
st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🇨🇳 ĐỒNG BỘ DỊCH SONG NGỮ TRUNG - VIỆT 🇻🇳")

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

# Thiết lập mã ngôn ngữ chuẩn xác
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
    st.error(f"Lỗi khởi tạo bộ lọc OCR: {str(e)}")
    reader = None

uploaded_file = None
if file_type == "File Excel (.xlsx)":
    uploaded_file = st.file_uploader("Tải lên file Excel gốc", type=["xlsx"])
else:
    uploaded_file = st.file_uploader("Tải lên file Ảnh gốc", type=["jpg", "png", "jpeg"])

# --- HÀM DỊCH THUẬT CHUYÊN DỤNG ---

def translate_text(text, src, dest):
    """Hàm dịch sử dụng API chuyên dụng, làm sạch chuỗi để tăng độ chính xác ngữ cảnh"""
    if not text or str(text).strip() == "":
        return text
    if isinstance(text, (int, float)):
        return text
    try:
        # Chuẩn hóa khoảng trắng để AI không dịch sai cấu trúc câu
        clean_text = " ".join(str(text).split())
        translated = GoogleTranslator(source=src, target=dest).translate(clean_text)
        return translated
    except Exception as e:
        return f"[Lỗi dịch: {str(e)}]"

# --- XỬ LÝ BẢO TOÀN ĐỊNH DẠNG TUYỆT ĐỐI (YÊU CẦU 1) ---

def process_excel(file_bytes, src, dest):
    """Đọc và chèn văn bản song ngữ bảo toàn Rich Text và định dạng ô tính"""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    # Bỏ qua nếu ô là công thức hoặc số thuần túy
                    if isinstance(cell.value, str) and cell.value.startswith('='):
                        continue
                    if isinstance(cell.value, (int, float)):
                        continue

                    # Trường hợp 1: Ô chứa định dạng Rich Text phức tạp (Nhiều style chữ trong 1 ô)
                    if isinstance(cell.value, RichText):
                        orig_rich = cell.value
                        new_rich = RichText()
                        
                        # Sao chép và giữ nguyên đoạn chữ gốc cùng định dạng của nó
                        for block in orig_rich:
                            new_rich.append(block.text, font=copy.copy(block.font))
                        
                        # Thêm dấu xuống dòng (Yêu cầu 2)
                        new_rich.append("\n")
                        
                        # Tiến hành dịch toàn bộ cụm văn bản Rich Text
                        full_text = "".join([block.text for block in orig_rich])
                        translated_text = translate_text(full_text, src, dest)
                        
                        # Tạo font mặc định hoặc lấy font của khối đầu tiên gán cho chữ dịch
                        dest_font = copy.copy(orig_rich[0].font) if len(orig_rich) > 0 else None
                        new_rich.append(translated_text, font=dest_font)
                        
                        cell.value = new_rich
                        
                    # Trường hợp 2: Ô chứa chuỗi văn bản thông thường (Đồng nhất 1 định dạng font)
                    elif isinstance(cell.value, str):
                        original_text = cell.value.strip()
                        if original_text:
                            translated_text = translate_text(original_text, src, dest)
                            
                            # Yêu cầu 2 & 3: Bản dịch nằm chung ô, ngay bên dưới
                            cell.value = f"{original_text}\n{translated_text}"
                    
                    # Cấu hình bắt buộc để hiển thị được dấu xuống dòng đa dòng
                    if cell.alignment:
                        cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal or "general",
                            vertical="top",
                            wrap_text=True,
                            shrink_to_fit=cell.alignment.shrink_to_fit,
                            indent=cell.alignment.indent,
                            text_rotation=cell.alignment.text_rotation
                        )
                    else:
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
                        
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

def process_image_to_excel(image_bytes, src, dest):
    """Xử lý Ảnh và định dạng bảng kết quả Excel gọn gàng"""
    if reader is None:
        return None
        
    image = Image.open(io.BytesIO(image_bytes)).convert('RGB')
    img_np = np.array(image)
    results = reader.readtext(img_np)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KetQuaDich"
    
    ws['A1'] = "Nội dung song ngữ (Gốc \n Dịch)"
    ws['A1'].font = openpyxl.styles.Font(bold=True)
    ws.column_dimensions['A'].width = 60
    
    row_idx = 2
    for (bbox, text, prob) in results:
        clean_text = text.strip()
        if clean_text:
            translated_text = translate_text(clean_text, src, dest)
            cell = ws.cell(row=row_idx, column=1)
            cell.value = f"{clean_text}\n{translated_text}"
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            row_idx += 1
            
    out_bio = io.BytesIO()
    wb.save(out_bio)
    out_bio.seek(0)
    return out_bio

# --- KHỞI CHẠY TIẾN TRÌNH (YÊU CẦU 6) ---

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    st.success("Tải file lên thành công! Sẵn sàng để dịch.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🚀 BẮT ĐẦU DỊCH", use_container_width=True):
            with st.spinner("Hệ thống AI đang dịch thuật chuyên sâu..."):
                try:
                    if file_type == "File Excel (.xlsx)":
                        output_data = process_excel(file_bytes, src_lang, dest_lang)
                    else:
                        output_data = process_image_to_excel(file_bytes, src_lang, dest_lang)
                    
                    if output_data is not None:
                        st.session_state['translated_output'] = output_data
                        st.success("🎉 Đã dịch xong toàn bộ nội dung!")
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi hệ thống: {str(e)}")

    with col2:
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
    st.info("Vui lòng tải file của bạn lên ở thanh menu bên trái để bắt đầu thực hiện.")
