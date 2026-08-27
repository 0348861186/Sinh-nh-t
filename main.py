import io
import os
from PIL import Image
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border

# Khởi tạo EasyOCR để đọc chữ từ ảnh (hỗ trợ Trung - Anh/Việt)
try:
    import easyocr
    @st.cache_resource
    def load_reader():
        return easyocr.Reader(['ch_sim', 'en'])
    reader = load_reader()
    has_easyocr = True
except Exception:
    has_easyocr = False

st.set_page_config(page_title="Phần mềm Dịch Song Ngữ Trung - Việt Chuẩn Format", layout="wide")

st.title("🈲 🇻🇳 Phần mềm Dịch Song Ngữ Trung - Việt (Giữ Nguyên Định Dạng)")
st.markdown("""
Ứng dụng đáp ứng các tiêu chí:
1. Chỉ dùng thư viện chuyên dụng (`openpyxl`, `deep-translator`, `easyocr`).
2. Hỗ trợ chọn file **Hình Ảnh** hoặc **Excel (.xlsx)**.
3. Dòng tiếng Việt nằm **ngay bên dưới** dòng tiếng Trung.
4. Nội dung dịch nằm **chung một ô** với nội dung gốc.
5. Giữ nguyên 100% bố cục, màu sắc, khung viền của file gốc.
""")

# Khởi tạo bộ dịch
translator = GoogleTranslator(source='zh-CN', target='vi')

def translate_text(text):
    if not text or str(text).strip() == "":
        return ""
    try:
        # Nếu là số hoặc ký tự đặc biệt, giữ nguyên
        if str(text).isdigit():
            return str(text)
        res = translator.translate(str(text))
        return res
    except Exception:
        return text

# Chọn định dạng file đầu vào
file_type = st.radio("Chọn định dạng file đầu vào:", ("File Excel (.xlsx)", "File Hình Ảnh (.png, .jpg, .jpeg)"))

uploaded_file = st.file_uploader("Tải file lên tại đây:", type=["xlsx", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    if "Excel" in file_type:
        st.success("Đã tải lên file Excel thành công!")
        
        try:
            # Đọc file bằng openpyxl để giữ trọn vẹn định dạng style, border, fill...
            wb = openpyxl.load_workbook(uploaded_file)
            ws = wb.active
            
            st.write("Xem trước cấu trúc dữ liệu Excel:")
            df_preview = pd.DataFrame(ws.values)
            st.dataframe(df_preview.head(10), use_container_width=True)
            
            if st.button("🚀 Bắt đầu dịch file Excel"):
                with st.spinner("Đang tiến hành dịch và đồng bộ vào từng ô..."):
                    for row in ws.iter_rows():
                        for cell in row:
                            val = cell.value
                            if val is not None and str(val).strip() != "":
                                # Tránh dịch lại nếu ô đã có chứa ký tự xuống dòng từ trước
                                val_str = str(val)
                                if "\n" not in val_str:
                                    translated = translate_text(val_str)
                                    if translated and translated != val_str:
                                        # Gộp chữ gốc ở trên, chữ Việt ở dưới trong CÙNG MỘT Ô
                                        cell.value = f"{val_str}\n{translated}"
                                        
                                        # Lấy định dạng cũ của cell và bật tính năng xuống dòng (Wrap text)
                                        current_alignment = cell.alignment
                                        horiz = current_alignment.horizontal if current_alignment and current_alignment.horizontal else 'center'
                                        vert = current_alignment.vertical if current_alignment and current_alignment.vertical else 'center'
                                        
                                        cell.alignment = Alignment(
                                            wrap_text=True, 
                                            vertical=vert, 
                                            horizontal=horiz
                                        )
                    
                    # Xuất file Excel sau khi dịch
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("Dịch file Excel hoàn tất!")
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
            st.error("Chưa cài đặt thư viện EasyOCR trong môi trường.")
        else:
            if st.button("🚀 Bắt đầu dịch ảnh sang Excel"):
                with st.spinner("Đang nhận diện chữ và dịch thuật chuẩn bố cục..."):
                    img_np = np.array(image)
                    # Nhận diện chữ bằng EasyOCR
                    results = reader.readtext(img_np)
                    
                    # Tạo file Excel mới mô phỏng bảng biểu đúng như ảnh mẫu yêu cầu
                    wb_img = openpyxl.Workbook()
                    ws_img = wb_img.active
                    ws_img.title = "Translated Image Table"
                    
                    # Tạo Header giống mẫu ảnh phân tích
                    headers = ["STT", "部分 / Bộ phận", "开几台机 / Số máy mở", "正式工 / Chính thức", "临时工 / Thời vụ", "备注 / Ghi chú"]
                    ws_img.append(headers)
                    
                    # Định dạng Header cam giống ảnh mẫu
                    header_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF")
                    thin_border = Border(
                        left=openpyxl.styles.borders.Side(style='thin', color='000000'),
                        right=openpyxl.styles.borders.Side(style='thin', color='000000'),
                        top=openpyxl.styles.borders.Side(style='thin', color='000000'),
                        bottom=openpyxl.styles.borders.Side(style='thin', color='000000')
                    )
                    
                    for col_num in range(1, len(headers) + 1):
                        cell = ws_img.cell(row=1, column=col_num)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                        cell.border = thin_border

                    # Trích xuất các dòng nội dung từ kết quả OCR để đưa vào bảng Excel tương ứng
                    # Sắp xếp các text theo tọa độ Y trên ảnh nếu cần, hoặc đưa lần lượt vào dòng
                    row_idx = 2
                    for bbox, text, prob in results:
                        # Bỏ qua các chữ tiêu đề nhỏ nếu muốn, hoặc dịch trực tiếp từng cụm
                        translated = translate_text(text)
                        combined_val = f"{text}\n{translated}"
                        
                        # Ghi vào dòng mô phỏng (cột 2 chứa nội dung gốc và dịch song ngữ chung 1 ô)
                        ws_img.append([row_idx - 1, combined_val, "", "", "", ""])
                        
                        target_cell = ws_img.cell(row=row_idx, column=2)
                        target_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                        target_cell.border = thin_border
                        
                        # Thêm border cho cả dòng
                        for c in range(1, 7):
                            ws_img.cell(row=row_idx, column=c).border = thin_border
                            ws_img.cell(row=row_idx, column=c).alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                        
                        row_idx += 1

                    # Lưu file kết quả từ ảnh
                    output_img = io.BytesIO()
                    wb_img.save(output_img)
                    output_img.seek(0)
                    
                    st.success("Đã chuyển đổi ảnh thành công sang Excel song ngữ!")
                    st.download_button(
                        label="📥 Tải xuống File Excel kết quả",
                        data=output_img,
                        file_name="translated_from_image.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
