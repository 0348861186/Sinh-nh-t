import io
import os
from PIL import Image
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl.styles import Alignment

# Thử import easyocr để đọc ảnh, nếu lỗi sẽ thông báo
try:
    import easyocr
    @st.cache_resource
    def load_reader():
        return easyocr.Reader(['ch_sim', 'en']) # Hỗ trợ tiếng Trung giản thể
    reader = load_reader()
    has_easyocr = True
except Exception:
    has_easyocr = False

st.set_page_config(page_title="Phần mềm Dịch Song Ngữ Trung - Việt", layout="wide")

st.title("🈲 🇻🇳 Phần mềm Dịch Song Ngữ Trung - Việt (Streamlit Cloud)")
st.markdown("""
Ứng dụng hỗ trợ:
1. Upload file **Ảnh** hoặc **Excel** (.xlsx).
2. Dịch tiếng Trung sang tiếng Việt nằm ngay bên dưới trong **cùng một ô** (đối với Excel) giữ nguyên hoàn toàn định dạng gốc.
""")

# Khởi tạo translator
translator = GoogleTranslator(source='zh-CN', target='vi')

def translate_text(text):
    if not text or str(text).strip() == "":
        return ""
    # Nếu là số hoặc không có ký tự Trung Quốc thì giữ nguyên hoặc dịch thử
    try:
        # Kiểm tra xem có phải chuỗi cần dịch không
        res = translator.translate(str(text))
        return res
    except Exception:
        return text

# Chọn loại file upload
file_type = st.radio("Chọn định dạng file đầu vào:", ("File Excel (.xlsx)", "File Hình Ảnh (.png, .jpg, .jpeg)"))

uploaded_file = st.file_uploader("Tải file lên tại đây:", type=["xlsx", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    if "Excel" in file_type:
        st.success("Đã tải lên file Excel thành công!")
        
        # Đọc file Excel bằng openpyxl để giữ nguyên định dạng
        try:
            # Lưu file tạm để openpyxl đọc
            wb = openpyxl.load_workbook(uploaded_file)
            ws = wb.active
            
            st.write("Xem trước dữ liệu gốc trên Excel:")
            # Hiển thị dạng bảng tạm bằng pandas để người dùng xem
            df_preview = pd.DataFrame(ws.values)
            st.dataframe(df_preview.head(10), use_container_width=True)
            
            if st.button("🚀 Bắt đầu dịch file Excel"):
                with st.spinner("Đang tiến hành dịch và giữ nguyên định dạng..."):
                    for row in ws.iter_rows():
                        for cell in row:
                            val = cell.value
                            if val is not None and isinstance(val, str) and val.strip() != "":
                                # Dịch nội dung
                                translated = translate_text(val)
                                if translated and translated != val:
                                    # Gộp nội dung gốc và dịch cách nhau xuống dòng (\n) trong CÙNG MỘT Ô
                                    cell.value = f"{val}\n{translated}"
                                    # Bật tính năng xuống dòng (Wrap text) và canh lề
                                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal=cell.alignment.horizontal if cell.alignment else 'center')
                    
                    # Lưu file sau khi dịch vào bộ nhớ đệm
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("Dịch thành công!")
                    st.download_button(
                        label="📥 Tải xuống File Excel đã dịch",
                        data=output,
                        file_name="translated_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Lỗi xử lý file Excel: {e}")

    else:
        st.success("Đã tải lên hình ảnh thành công!")
        image = Image.open(uploaded_file)
        st.image(image, caption="Ảnh gốc tải lên", use_container_width=True)
        
        if not has_easyocr:
            st.error("Thư viện OCR chưa sẵn sàng trên môi trường này. Vui lòng đảm bảo easyocr được cài đặt trong requirements.txt")
        else:
            if st.button("🚀 Bắt đầu dịch ảnh sang Excel"):
                with st.spinner("Đang nhận diện chữ (OCR) và dịch..."):
                    # Chuyển ảnh sang numpy array cho easyocr
                    img_np = np.array(image)
                    results = reader.readtext(img_np)
                    
                    # Tạo một DataFrame hoặc cấu trúc Excel giả lập dựa trên kết quả OCR
                    # Vì ảnh là dạng lưới, ta xuất ra file Excel mô phỏng lại dạng bảng như yêu cầu
                    data_rows = []
                    for bbox, text, prob in results:
                        translated = translate_text(text)
                        combined_text = f"{text}\n{translated}"
                        data_rows.append({"Nội dung gốc & Dịch": combined_text})
                    
                    # Tạo file Excel mới từ kết quả nhận diện ảnh
                    wb_img = openpyxl.Workbook()
                    ws_img = wb_img.active
                    ws_img.title = "Translated Image"
                    
                    # Ghi tiêu đề bảng mô phỏng theo yêu cầu ảnh mẫu
                    ws_img.append(["STT", "Bộ phận / Nội dung gốc & Dịch"])
                    ws_img.cell(row=1, column=1).alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                    ws_img.cell(row=1, column=2).alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                    
                    for idx, item in enumerate(data_rows, start=1):
                        row_cells = ws_img.append([idx, item["Nội dung gốc & Dịch"]])
                        current_row = ws_img.max_row
                        cell_target = ws_img.cell(row=current_row, column=2)
                        cell_target.alignment = Alignment(wrap_text=True, vertical='center')
                    
                    # Lưu file Excel kết quả từ ảnh
                    output_img = io.BytesIO()
                    wb_img.save(output_img)
                    output_img.seek(0.0 if hasattr(output_img, 'seek') else None)
                    output_img.seek(0)
                    
                    st.success("Đã chuyển đổi và dịch ảnh thành công sang Excel!")
                    st.download_button(
                        label="📥 Tải xuống File Excel kết quả",
                        data=output_img,
                        file_name="translated_from_image.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
