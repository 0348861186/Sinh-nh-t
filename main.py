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
model = genai.GenerativeModel('gemini-3.6-flash')
# -----------------------------------------------------------------------------
# 1. XỬ LÝ DỊCH EXCEL (GOM BATCH TRÁNH LỖI QUOTA)
# -----------------------------------------------------------------------------
def process_excel(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file)
    cells_to_translate = []
    texts_to_translate = []

    # Quét gom toàn bộ văn bản cần dịch trong các sheet
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                val = str(cell.value).strip() if cell.value else ""
                if val and isinstance(cell.value, str) and not val.startswith('=') and not val.isdigit():
                    cells_to_translate.append(cell)
                    texts_to_translate.append(val)

    if not texts_to_translate:
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    # Gửi 1 LẦN duy nhất toàn bộ danh sách câu để tiết kiệm API
    prompt = f"""Bạn là biên dịch viên công xưởng chuyên nghiệp. 
Hãy dịch từng câu trong danh sách JSON dưới đây sang dạng song ngữ: Dòng 1 Tiếng Trung, Dòng 2 Tiếng Việt.
Giữ nguyên thứ tự danh sách.

Yêu cầu output: Trả về duy nhất 1 danh sách chuỗi JSON (JSON array of strings).

Danh sách cần dịch:
{json.dumps(texts_to_translate, ensure_ascii=False)}"""

    try:
        response = model.generate_content(prompt)
        res_text = response.text.strip()
        if "```json" in res_text:
            res_text = res_text.split("```json")[1].split("```")[0].strip()
        elif "```" in res_text:
            res_text = res_text.split("```")[1].split("```")[0].strip()
            
        translated_list = json.loads(res_text)

        # Gán lại kết quả dịch vào từng ô Excel
        for idx, cell in enumerate(cells_to_translate):
            if idx < len(translated_list):
                cell.value = translated_list[idx]
                
                # Bật Wrap Text cho ô
                align = cell.alignment
                if align:
                    cell.alignment = Alignment(
                        horizontal=align.horizontal,
                        vertical=align.vertical,
                        wrap_text=True
                    )
                else:
                    cell.alignment = Alignment(wrap_text=True)
    except Exception as e:
        st.error(f"Lỗi khi dịch Excel qua Gemini: {str(e)}")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------------------------------------------------------
# 2. XỬ LÝ DỊCH VÀ ĐÈ CHỮ LÊN ẢNH
# -----------------------------------------------------------------------------
def process_image(pil_image):
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
        response = model.generate_content([pil_image, prompt])
        res_text = response.text.strip()
        
        if "```json" in res_text:
            res_text = res_text.split("```json")[1].split("```")[0].strip()
        elif "```" in res_text:
            res_text = res_text.split("```")[1].split("```")[0].strip()
            
        items = json.loads(res_text)
    except Exception as e:
        st.error(f"Không thể phân tích ảnh bằng Gemini: {str(e)}")
        return pil_image

    img_draw = pil_image.copy()
    draw = ImageDraw.Draw(img_draw)
    width, height = pil_image.size

    for item in items:
        box = item.get("box_2d")
        text = item.get("text")
        if not box or not text:
            continue
            
        ymin, xmin, ymax, xmax = box
        left = int(xmin * width / 1000)
        top = int(ymin * height / 1000)
        right = int(xmax * width / 1000)
        bottom = int(ymax * height / 1000)

        # Xóa chữ cũ bằng khối trắng
        draw.rectangle([left, top, right, bottom], fill="white", outline="white")

        box_height = max(10, bottom - top)
        font_size = max(10, int(box_height / 3))
        
        # Sửa lỗi font trên Streamlit Cloud Linux
        try:
            font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
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
                    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                    page = doc.load_page(0)
                    pix = page.get_pixmap()
                    input_image = Image.open(io.BytesIO(pix.tobytes()))
                else:
                    input_image = Image.open(uploaded_file)
                
                output_image = process_image(input_image)
                
                st.success("🎉 Dịch thành công file Ảnh/PDF!")
                col1, col2 = st.columns(2)
                with col1:
                    st.image(input_image, caption="File Gốc", use_container_width=True)
                with col2:
                    st.image(output_image, caption="File Dịch Song Ngữ (Y Chang Vị Trí Gốc)", use_container_width=True)
                
                img_byte_arr = io.BytesIO()
                output_image.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                
                st.download_button(
                    label="📥 Tải Ảnh Song Ngữ Sau Khi Dịch",
                    data=img_byte_arr,
                    file_name="SongNgu_" + uploaded_file.name.split(".")[0] + ".png",
                    mime="image/png"
                )
