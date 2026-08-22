import streamlit as st
import openpyxl
from openpyxl.styles import Alignment
import google.generativeai as genai
from PIL import Image, ImageDraw, ImageFont
import fitz  # PyMuPDF
import io
import json

st.set_page_config(page_title="Dịch Song Ngữ Trung - Việt", layout="wide")
st.title("🌐 Ứng dụng Dịch Song Ngữ Trung - Việt (Giữ Nguyên Định Dạng)")

# Nhập Gemini API Key
api_key = st.sidebar.text_input("🔑 Nhập Gemini API Key:", type="password")

if not api_key:
    st.warning("⚠️ Vui lòng nhập Gemini API Key ở thanh bên trái để sử dụng ứng dụng.")
    st.info("💡 Bạn có thể lấy Gemini API Key miễn phí tại: https://aistudio.google.com/")
    st.stop()

# Cấu hình Gemini
genai.configure(api_key=api_key)
model_text = genai.GenerativeModel('gemini-1.5-flash')

# -----------------------------------------------------------------------------
# HÀM DỊCH EXCEL
# -----------------------------------------------------------------------------
def translate_text_bilingual(text):
    """Dịch câu/đoạn văn bản sang dạng: Tiếng Trung ở trên \n Tiếng Việt ở dưới"""
    if not text or not isinstance(text, str) or text.strip().isdigit():
        return text
    if text.startswith('='):  # Bỏ qua công thức Excel
        return text

    prompt = f"""Bạn là một dịch thuật viên công xưởng chuyên nghiệp Trung - Việt.
Hãy chuyển đổi đoạn văn bản sau thành dạng song ngữ chính xác:
Dòng 1: Tiếng Trung (viết nguyên gốc hoặc dịch sang Tiếng Trung nếu đang là tiếng Việt)
Dòng 2: Tiếng Việt (dịch chuẩn ngữ cảnh công xưởng)

Quy tắc bắt buộc: Chỉ trả về 2 dòng chữ, dòng trên Tiếng Trung, dòng dưới Tiếng Việt. Không thêm bất kỳ lời giải thích nào khác.

Văn bản:
{text}"""

    try:
        response = model_text.generate_content(prompt)
        res_text = response.text.strip()
        return res_text
    except Exception as e:
        return text

def process_excel(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = translate_text_bilingual(cell.value)
                    
                    # Căn lề và bật Wrap Text để hiển thị 2 dòng
                    align = cell.alignment
                    if align:
                        cell.alignment = Alignment(
                            horizontal=align.horizontal,
                            vertical=align.vertical,
                            wrap_text=True
                        )
                    else:
                        cell.alignment = Alignment(wrap_text=True)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------------------------------------------------------
# HÀM DỊCH VÀ ĐÈ CHỮ LÊN ẢNH (Bảo toàn bố cục y chang)
# -----------------------------------------------------------------------------
def process_image(pil_image):
    # Dùng Gemini Vision nhận diện chữ, vị trí (bounding box) và dịch song ngữ
    prompt = """Hãy phân tích bức ảnh này. Nhận diện toàn bộ các vùng có chữ.
Đối với mỗi vùng chữ, hãy:
1. Xác định tọa độ hộp chứa chữ dạng tỷ lệ [ymin, xmin, ymax, xmax] từ 0 đến 1000.
2. Dịch văn bản đó sang dạng song ngữ: Dòng 1 Tiếng Trung, Dòng 2 Tiếng Việt.

Trả về kết quả duy nhất ở dạng cấu trúc JSON danh sách như sau:
[
  {"box_2d": [ymin, xmin, ymax, xmax], "text": "Tiếng Trung\\nTiếng Việt"}
]
"""
    try:
        response = model_text.generate_content([pil_image, prompt])
        res_text = response.text.strip()
        
        # Làm sạch chuỗi json trả về từ Gemini
        if "```json" in res_text:
            res_text = res_text.split("```json")[1].split("```")[0].strip()
        elif "```" in res_text:
            res_text = res_text.split("```")[1].split("```")[0].strip()
            
        items = json.loads(res_text)
    except Exception as e:
        st.error(f"Không thể phân tích ảnh bằng Gemini: {str(e)}")
        return pil_image

    # Tạo bản sao ảnh để vẽ lại chữ mới đè lên vị trí cũ
    img_draw = pil_image.copy()
    draw = ImageDraw.Draw(img_draw)
    width, height = pil_image.size

    for item in items:
        box = item.get("box_2d")
        text = item.get("text")
        if not box or not text:
            continue
            
        # Quy đổi tọa độ 0-1000 ra pixel thực tế trên ảnh
        ymin, xmin, ymax, xmax = box
        left = int(xmin * width / 1000)
        top = int(ymin * height / 1000)
        right = int(xmax * width / 1000)
        bottom = int(ymax * height / 1000)

        # 1. Xóa chữ cũ bằng cách vẽ hình chữ nhật màu trắng đè lên vùng chữ cũ
        draw.rectangle([left, top, right, bottom], fill="white", outline="white")

        # 2. Cấu hình Font chữ và vẽ chữ Song ngữ mới vào đúng vị trí
        box_height = max(10, bottom - top)
        font_size = max(10, int(box_height / 3))
        try:
            # Dùng font mặc định hỗ trợ Unicode
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()

        draw.text((left + 2, top + 2), text, fill="black", font=font)

    return img_draw

# -----------------------------------------------------------------------------
# GIAO DIỆN XỬ LÝ UPLOAD FILE
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader("📂 Tải file Excel, Ảnh hoặc PDF lên đây", type=["xlsx", "png", "jpg", "jpeg", "pdf"])

if uploaded_file is not None:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    
    if st.button("🚀 Bắt đầu dịch & Giữ nguyên định dạng"):
        with st.spinner("Hệ thống AI đang xử lý file của bạn..."):
            
            # --- Trường hợp 1: File Excel ---
            if file_ext == "xlsx":
                result_excel = process_excel(uploaded_file)
                st.success("🎉 Dịch thành công file Excel!")
                st.download_button(
                    label="📥 Tải file Excel Song Ngữ",
                    data=result_excel,
                    file_name="SongNgu_" + uploaded_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            # --- Trường hợp 2: File Ảnh hoặc PDF ---
            else:
                if file_ext == "pdf":
                    # Đọc trang đầu tiên của file PDF thành Ảnh
                    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                    page = doc.load_page(0)
                    pix = page.get_pixmap()
                    input_image = Image.open(io.BytesIO(pix.tobytes()))
                else:
                    input_image = Image.open(uploaded_file)
                
                # Xử lý dịch và đè chữ lên ảnh
                output_image = process_image(input_image)
                
                st.success("🎉 Dịch thành công file Ảnh/PDF!")
                col1, col2 = st.columns(2)
                with col1:
                    st.image(input_image, caption="File Gốc", use_column_width=True)
                with col2:
                    st.image(output_image, caption="File Dịch Song Ngữ (Y Chang Vị Trí Gốc)", use_column_width=True)
                
                # Chuẩn bị nút tải ảnh về
                img_byte_arr = io.BytesIO()
                output_image.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                
                st.download_button(
                    label="📥 Tải Ảnh Song Ngữ Sau Khi Dịch",
                    data=img_byte_arr,
                    file_name="SongNgu_" + uploaded_file.name.split(".")[0] + ".png",
                    mime="image/png"
                )
