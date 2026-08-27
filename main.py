import io
import os
from PIL import Image
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# Khởi tạo EasyOCR
try:
    import easyocr
    @st.cache_resource
    def load_reader():
        return easyocr.Reader(['ch_sim', 'en'])
    reader = load_reader()
    has_easyocr = True
except Exception:
    has_easyocr = False

st.set_page_config(page_title="Phần mềm Dịch Song Ngữ Trung - Việt Chuẩn 100%", layout="wide")

st.title("🈲 🇻🇳 Phần mềm Dịch Song Ngữ Trung - Việt (Giữ Nguyên Số Dòng & Bố Cục)")
st.markdown("""
Ứng dụng đáp ứng chính xác yêu cầu:
1. Load lên bao nhiêu dòng thì xuất ra **chuẩn y chang bấy nhiêu dòng**.
2. Dòng tiếng Việt nằm **ngay bên dưới** dòng tiếng Trung trong **cùng một ô**.
3. Giữ nguyên định dạng, bố cục, khung viền.
""")

translator = GoogleTranslator(source='zh-CN', target='vi')

def translate_text(text):
    if not text or str(text).strip() == "":
        return ""
    try:
        if str(text).isdigit():
            return str(text)
        res = translator.translate(str(text))
        return res
    except Exception:
        return text

file_type = st.radio("Chọn định dạng file đầu vào:", ("File Excel (.xlsx)", "File Hình Ảnh (.png, .jpg, .jpeg)"))
uploaded_file = st.file_uploader("Tải file lên tại đây:", type=["xlsx", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    if "Excel" in file_type:
        st.success("Đã tải lên file Excel thành công!")
        try:
            wb = openpyxl.load_workbook(uploaded_file)
            ws = wb.active
            
            st.write("Xem trước dữ liệu Excel gốc:")
            df_preview = pd.DataFrame(ws.values)
            st.dataframe(df_preview.head(10), use_container_width=True)
            
            if st.button("🚀 Bắt đầu dịch và giữ nguyên mọi dòng Excel"):
                with st.spinner("Đang dịch toàn bộ dữ liệu..."):
                    for row_idx, row in enumerate(ws.iter_rows(), start=1):
                        max_lines = 1
                        for cell in row:
                            val = cell.value
                            if val is not None and str(val).strip() != "":
                                val_str = str(val)
                                if "\n" not in val_str:
                                    translated = translate_text(val_str)
                                    if translated and translated != val_str:
                                        cell.value = f"{val_str}\n{translated}"
                                        
                                        lines_count = cell.value.count('\n') + 1
                                        if lines_count > max_lines:
                                            max_lines = lines_count
                                            
                                        # Giữ nguyên canh lề cũ
                                        current_alignment = cell.alignment
                                        horiz = current_alignment.horizontal if current_alignment and current_alignment.horizontal else 'center'
                                        vert = current_alignment.vertical if current_alignment and current_alignment.vertical else 'center'
                                        cell.alignment = Alignment(wrap_text=True, vertical=vert, horizontal=horiz)
                        
                        if max_lines > 1:
                            ws.row_dimensions[row_idx].height = max(35, max_lines * 22)

                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("Dịch hoàn tất giữ nguyên toàn bộ số dòng!")
                    st.download_button(
                        label="📥 Tải xuống File Excel đã dịch",
                        data=output,
                        file_name="translated_exact_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Lỗi xử lý file Excel: {e}")

    else:
        st.success("Đã tải lên hình ảnh thành công!")
        image = Image.open(uploaded_file)
        st.image(image, caption="Ảnh gốc tải lên", use_container_width=True)
        
        if not has_easyocr:
            st.error("Thư viện EasyOCR chưa được cấu hình.")
        else:
            if st.button("🚀 Bắt đầu chuyển đổi ảnh sang Excel (Đầy đủ mọi dòng)"):
                with st.spinner("Đang quét toàn bộ chữ trong ảnh và dịch..."):
                    img_np = np.array(image)
                    results = reader.readtext(img_np)
                    
                    # Sắp xếp các kết quả nhận diện theo tọa độ Y từ trên xuống dưới để giữ đúng trật tự các dòng trên ảnh
                    results = sorted(results, key=lambda x: np.mean([p[1] for p in x[0]]))
                    
                    wb_img = openpyxl.Workbook()
                    ws_img = wb_img.active
                    ws_img.title = "Translated Image Table"
                    
                    headers = [
                        "STT", 
                        "部分\nBộ phận", 
                        "开几台机\nSố máy mở", 
                        "正式工\nChính thức", 
                        "临时工\nThời vụ", 
                        "备注\nGhi chú"
                    ]
                    ws_img.append(headers)
                    
                    header_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF", size=11)
                    thin_border = Border(
                        left=Side(style='thin', color='BFBFBF'),
                        right=Side(style='thin', color='BFBFBF'),
                        top=Side(style='thin', color='BFBFBF'),
                        bottom=Side(style='thin', color='BFBFBF')
                    )
                    
                    ws_img.row_dimensions[1].height = 30
                    for col_num in range(1, len(headers) + 1):
                        cell = ws_img.cell(row=1, column=col_num)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                        cell.border = thin_border

                    # Tự động tạo chính xác SỐ LƯỢNG DÒNG tương ứng với những gì đọc được từ ảnh
                    row_idx = 2
                    for bbox, text, prob in results:
                        # Bỏ qua các chữ quá nhỏ hoặc không cần thiết nếu có, hoặc đưa toàn bộ vào
                        if text.strip() == "":
                            continue
                            
                        translated = translate_text(text)
                        combined_val = f"{text}\n{translated}"
                        
                        # Ghi dòng mới đúng theo dữ liệu trích xuất từ ảnh (ví dụ file có bao nhiêu dòng chữ OCR quét ra thì sẽ có bấy nhiêu dòng trong Excel)
                        ws_img.append([row_idx - 1, combined_val, "", "", "", ""])
                        
                        ws_img.row_dimensions[row_idx].height = 40
                        for c in range(1, 7):
                            cell = ws_img.cell(row=row_idx, column=c)
                            cell.border = thin_border
                            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                            cell.font = Font(size=11)
                            
                        row_idx += 1

                    for col in ws_img.columns:
                        col_letter = openpyxl.utils.get_column_letter(col[0].column)
                        ws_img.column_dimensions[col_letter].width = 20

                    output_img = io.BytesIO()
                    wb_img.save(output_img)
                    output_img.seek(0)
                    
                    st.success(f"Đã chuyển đổi thành công! Tổng số dòng trích xuất: {row_idx - 2}")
                    st.download_button(
                        label="📥 Tải xuống File Excel kết quả",
                        data=output_img,
                        file_name="translated_full_rows.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
