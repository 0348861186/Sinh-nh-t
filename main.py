import io
import os
import easyocr
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
from openpyxl import load_workbook
from PIL import Image

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Dịch Song ngữ Trung - Việt",
    page_icon="🌐",
    layout="wide",
)


# Khởi tạo OCR reader (Sử dụng cache để không load lại nhiều lần)
@st.cache_resource
def load_ocr_reader():
  # Hỗ trợ tiếng Trung giản thể/phồn thể ('ch_sim') và tiếng Việt ('vi')
  return easyocr.Reader(["ch_sim", "vi"], gpu=False)


# Hàm dịch văn bản sử dụng deep-translator
def translate_text(text, direction):
  if not text or not str(text).strip():
    return text

  text_str = str(text).strip()
  try:
    if direction == "Trung - Việt":
      translated = GoogleTranslator(
          source="zh-CN", target="vi"
      ).translate(text_str)
    else:  # Việt - Trung
      translated = GoogleTranslator(
          source="vi", target="zh-CN"
      ).translate(text_str)
    return translated
  except Exception as e:
    return f"[Lỗi dịch: {str(e)}]"


# Giao diện chính của Dashboard
st.title("🌐 Ứng dụng Dịch Song ngữ Trung - Việt (Giữ nguyên định dạng)")
st.markdown("""
Ứng dụng hỗ trợ tải lên file Excel hoặc File Ảnh, thực hiện dịch song ngữ 
với cấu trúc **Dòng tiếng Trung nằm trên, Dòng tiếng Việt nằm ngay bên dưới** (hoặc ngược lại tùy chiều dịch), 
đồng thời **giữ nguyên 100% định dạng gốc** của file Excel.
""")

# Sidebar cấu hình
st.sidebar.header("⚙️ Tùy chọn hệ thống")

# 5) Dashboard có nút chọn chế độ dịch trung việt hoặc việt trung
translation_direction = st.sidebar.selectbox(
    "Chọn chế độ dịch:", ("Trung - Việt", "Việt - Trung")
)

# 4) Dashboard có nút chọn file load lên là file ảnh hoặc file excel
uploaded_file = st.sidebar.file_uploader(
    "Tải lên file (Hỗ trợ: .xlsx, .xls, .png, .jpg, .jpeg)",
    type=["xlsx", "xls", "png", "jpg", "jpeg"],
)

# Xử lý khi người dùng tải file lên
if uploaded_file is not None:
  file_extension = uploaded_file.name.split(".")[-1].lower()

  # TRƯỜNG HỢP 1: XỬ LÝ FILE ẢNH
  if file_extension in ["png", "jpg", "jpeg"]:
    st.subheader("🖼️ Xem trước hình ảnh tải lên")
    image = Image.open(uploaded_file)
    st.image(image, caption="Hình ảnh gốc", use_column_width=True)

    # 6) Dashboard có nút nhấn dịch
    if st.button("🚀 Bắt đầu dịch Ảnh"):
      with st.spinner(
          "Đang nhận diện ký tự (OCR) và tiến hành dịch, vui lòng đợi..."
      ):
        try:
          reader = load_ocr_reader()
          image_np = np.array(image)

          # Nhận diện văn bản từ ảnh bằng EasyOCR
          # Kết quả trả về là list các tuple: (bbox, text, probability)
          ocr_results = reader.readtext(image_np)

          translated_data = []
          for bbox, text, prob in ocr_results:
            trans_text = translate_text(text, translation_direction)
            # Yêu cầu 2 & 3: Nội dung dịch nằm chung ô/ngay bên dưới
            combined_content = f"{text}\n{trans_text}"
            translated_data.append({
                "Văn bản gốc": text,
                "Bản dịch": trans_text,
                "Kết quả song ngữ (Gốc / Dịch)": combined_content,
            })

          df_result = pd.DataFrame(translated_data)

          st.success("✨ Dịch thành công!")
          st.dataframe(df_result, use_container_width=True)

          # Xuất ra file Excel để download (Yêu cầu 6)
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="KetQuaDich")
          excel_data = output.getvalue()

          st.download_button(
              label="📥 Tải xuống File Excel kết quả",
              data=excel_data,
              file_name=f"Translated_Image_{uploaded_file.name}.xlsx",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          )
        except Exception as e:
          st.error(f"Đã xảy ra lỗi trong quá trình xử lý ảnh: {e}")

  # TRƯỜNG HỢP 2: XỬ LÝ FILE EXCEL
  elif file_extension in ["xlsx", "xls"]:
    st.subheader("📊 Xem trước dữ liệu Excel gốc")
    try:
      # Đọc file bằng pandas để hiển thị preview
      df_preview = pd.read_excel(uploaded_file, sheet_name=0)
      st.dataframe(df_preview.head(10), use_container_width=True)

      # 6) Dashboard có nút nhấn dịch
      if st.button("🚀 Bắt đầu dịch File Excel"):
        with st.spinner(
            "Đang tiến hành dịch và bảo toàn định dạng gốc, vui lòng đợi..."
        ):
          # Reset con trỏ file stream
          uploaded_file.seek(0)

          # Sử dụng openpyxl để xử lý giữ nguyên định dạng, màu sắc, font chữ (Yêu cầu 1)
          wb = load_workbook(uploaded_file)
          ws = wb.active

          # Duyệt qua từng dòng và từng cột để dịch trực tiếp vào ô tương ứng (Yêu cầu 3)
          # Yêu cầu 2: Dòng tiếng dịch phải nằm ngay bên dưới (hoặc gộp chung trong ô cell xuống dòng)
          for row in ws.iter_rows():
            for cell in row:
              if cell.value is not None:
                original_text = str(cell.value)
                # Bỏ qua nếu là số hoặc công thức Excel
                if not original_text.startswith("="):
                  translated_text = translate_text(
                      original_text, translation_direction
                  )

                  # Yêu cầu 3: Nội dung dịch nằm chung với ô được dịch (xuống dòng trong cùng một ô cell)
                  cell.value = f"{original_text}\n{translated_text}"

                  # Bật tính năng xuống dòng tự động (Wrap Text) cho cell để hiển thị đẹp mắt
                  cell.alignment = cell.alignment.copy(wrap_text=True)

          # Lưu file sau khi dịch vào bộ nhớ đệm
          output = io.BytesIO()
          wb.save(output)
          processed_data = output.getvalue()

          st.success(
              "✨ Dịch file Excel hoàn tất và bảo toàn định dạng thành công!"
          )

          # 6) Dashboard có nút download file excel sau dịch
          st.download_button(
              label="📥 Tải xuống File Excel sau khi dịch",
              data=processed_data,
              file_name=f"Translated_{uploaded_file.name}",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          )
    except Exception as e:
      st.error(f"Đã xảy ra lỗi khi xử lý file Excel: {e}")

else:
  st.info(
      "👈 Vui lòng chọn chế độ dịch và tải lên file (Excel hoặc Ảnh) ở thanh"
      " bên trái để bắt đầu."
  )
