import io
import openpyxl
import pandas as pd
from PIL import Image
import pytesseract
import streamlit as st
from deep_translator import GoogleTranslator

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Dịch Song Ngữ Trung - Việt (Excel & Ảnh)", layout="centered"
)

st.title("🈲🇻🇳 Ứng Dụng Dịch Song Ngữ Trung - Việt")
st.markdown(
    """
Hỗ trợ tải lên file **Excel (.xlsx, .xls)** hoặc **Hình ảnh (.png, .jpg, .jpeg)**.
Bản dịch tiếng Việt sẽ nằm ngay bên dưới câu tiếng Trung và giữ nguyên bố cục gốc!
"""
)

# Khởi tạo bộ dịch
translator = GoogleTranslator(source="zh-CN", target="vi")


def translate_text(text):
  if not text or not str(text).strip():
    return ""
  # Bỏ qua nếu toàn là số hoặc ký tự đặc biệt
  if str(text).strip().isdigit():
    return str(text)
  try:
    # Chia nhỏ nếu đoạn quá dài để tránh quá giới hạn
    translated = translator.translate(str(text))
    return translated if translated else str(text)
  except Exception:
    return str(text)


# Chọn loại file đầu vào
option = st.radio(
    "Chọn định dạng file đầu vào:", ("File Excel (.xlsx)", "Hình ảnh (.png, .jpg)")
)

uploaded_file = st.file_uploader(
    "Tải file lên tại đây",
    type=["xlsx", "xls"] if "Excel" in option else ["png", "jpg", "jpeg"],
)

if uploaded_file is not None:
  if "Excel" in option:
    st.success("Đã tải lên file Excel thành công!")

    # Đọc file bằng openpyxl để giữ nguyên toàn bộ định dạng (font, màu, border...)
    try:
      wb = openpyxl.load_workbook(uploaded_file)
      sheet_names = wb.sheetnames
      selected_sheet = st.selectbox(
          "Chọn sheet cần dịch:", sheet_names, index=0
      )

      if st.button("Bắt đầu dịch"):
        with st.spinner("Đang tiến hành dịch và giữ nguyên định dạng..."):
          ws = wb[selected_sheet]

          # Duyệt qua từng ô có dữ liệu
          for row in ws.iter_rows():
            for cell in row:
              if cell.value is not None:
                original_text = str(cell.value)
                # Dịch nội dung
                translated_text = translate_text(original_text)

                # Yêu cầu 3 & 4: Dòng tiếng Việt nằm ngay bên dưới, cùng ô
                # Sử dụng ký tự xuống dòng '\n' trong Excel
                cell.value = f"{original_text}\n{translated_text}"

                # Bật tính năng xuống dòng tự động (Wrap Text) cho ô để hiển thị đẹp mắt
                cell.alignment = openpyxl.styles.Alignment(
                    wrap_text=True,
                    vertical="top",
                    horizontal=cell.alignment.horizontal,
                )

          # Lưu file vào bộ nhớ đệm
          output = io.BytesIO()
          wb.save(output)
          output.seek(0)

          st.success("Dịch thành công!")

          # Nút tải xuống file Excel
          st.download_button(
              label="📥 Tải xuống file Excel song ngữ",
              data=output,
              file_name="song_ngu_trung_viet.xlsx",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          )
    except Exception as e:
      st.error(f"Đã xảy ra lỗi khi xử lý file Excel: {e}")

  else:  # Xử lý Hình ảnh
    st.success("Đã tải lên hình ảnh thành công!")
    image = Image.open(uploaded_file)
    st.image(image, caption="Hình ảnh gốc", use_column_width=True)

    if st.button("Bắt đầu dịch từ ảnh"):
      with st.spinner(
          "Đang nhận diện chữ tiếng Trung và tạo file Excel kết quả..."
      ):
        try:
          # Nhận diện chữ tiếng Trung giản thể bằng Tesseract OCR
          extracted_data = pytesseract.image_to_data(
              image, lang="chi_sim", output_type=pytesseract.Output.DATAFRAME
          )
          extracted_data = extracted_data.dropna(subset=["text"])
          extracted_data = extracted_data[
              extracted_data["text"].str.strip() != ""
          ]

          # Gom nhóm các dòng chữ dựa trên vị trí tọa độ (block/line)
          lines = []
          for _, group in extracted_data.groupby(["block_num", "line_num"]):
            line_text = " ".join(group["text"].tolist())
            lines.append(line_text)

          # Tạo file Excel mới từ văn bản trích xuất
          wb_img = openpyxl.Workbook()
          ws_img = wb_img.active
          ws_img.title = "DichTuAnh"

          ws_img["A1"] = "Nội dung Song ngữ (Trung - Việt)"
          ws_img["A1"].font = openpyxl.styles.Font(bold=True, size=12)

          current_row = 2
          for line in lines:
            translated_line = translate_text(line)
            # Dòng tiếng Việt nằm ngay bên dưới dòng tiếng Trung trong cùng một ô
            ws_img.cell(
                row=current_row, column=1, value=f"{line}\n{translated_line}"
            )
            ws_img.cell(row=current_row, column=1).alignment = (
                openpyxl.styles.Alignment(wrap_text=True, vertical="top")
            )
            current_row += 1

          output_img = io.BytesIO()
          wb_img.save(output_img)
          output_img.seek(0)

          st.success("Nhận diện và dịch hoàn tất!")

          # Nút tải xuống file Excel
          st.download_button(
              label="📥 Tải xuống file Excel kết quả",
              data=output_img,
              file_name="ket_qua_dich_anh.xlsx",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          )
        except Exception as e:
          st.error(
              f"Lỗi khi xử lý hình ảnh (Đảm bảo đã cấu hình tesseract-ocr trên"
              f" cloud): {e}"
          )
