import io
import os
import easyocr
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Dịch Song ngữ Trung - Việt Chuẩn Xác",
    page_icon="🌐",
    layout="wide",
)


# Khởi tạo OCR reader (Sử dụng cache)
@st.cache_resource
def load_ocr_reader():
  return easyocr.Reader(["ch_sim", "en"], gpu=False)


# Hàm dịch thuật
def translate_text(text, direction):
  if not text or not str(text).strip():
    return text
  text_str = str(text).strip()

  # Bỏ qua nếu là số thuần túy hoặc công thức
  try:
    float(text_str)
    return text_str
  except ValueError:
    pass

  if text_str.startswith("="):
    return text_str

  try:
    if direction == "Trung - Việt":
      return GoogleTranslator(source="zh-CN", target="vi").translate(text_str)
    else:
      return GoogleTranslator(source="vi", target="zh-CN").translate(text_str)
  except Exception as e:
    return text_str


# Giao diện Dashboard
st.title("🌐 Hệ thống Dịch Song ngữ Trung - Việt (Bảo toàn định dạng)")
st.markdown("""
Ứng dụng hỗ trợ dịch song ngữ Trung - Việt:
- **Đối với File Excel:** Giữ nguyên 100% định dạng, màu sắc, cấu trúc gộp ô của file gốc.
- **Đối với File Ảnh:** Nhận diện bảng biểu, sắp xếp đúng hàng/cột và xuất ra file Excel sạch sẽ, trực quan.
""")

st.sidebar.header("⚙️ Tùy chọn hệ thống")
translation_direction = st.sidebar.selectbox(
    "Chọn chế độ dịch:", ("Trung - Việt", "Việt - Trung")
)

uploaded_file = st.sidebar.file_uploader(
    "Tải lên file (Hỗ trợ: .xlsx, .xls, .png, .jpg, .jpeg)",
    type=["xlsx", "xls", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
  file_extension = uploaded_file.name.split(".")[-1].lower()

  # =========================================================
  # TRƯỜNG HỢP 1: XỬ LÝ FILE ẢNH (Tái tạo bảng biểu dạng lưới)
  # =========================================================
  if file_extension in ["png", "jpg", "jpeg"]:
    st.subheader("🖼️ Xem trước hình ảnh tải lên")
    image = Image.open(uploaded_file)
    st.image(image, caption="Hình ảnh gốc", width=650)

    if st.button("🚀 Xử lý Ảnh & Dịch Song ngữ"):
      with st.spinner(
          "Đang bóc tách bố cục bảng và dịch thuật, vui lòng đợi..."
      ):
        try:
          reader = load_ocr_reader()
          image_np = np.array(image)
          results = reader.readtext(image_np)

          if not results:
            st.warning("Không tìm thấy văn bản nào trong ảnh!")
          else:
            # Gom nhóm các box chữ thành các dòng (Rows) theo tọa độ Y
            items = []
            for bbox, text, prob in results:
              y_center = (bbox[0][1] + bbox[2][1]) / 2
              x_center = (bbox[0][0] + bbox[2][0]) / 2
              items.append({"text": text, "x": x_center, "y": y_center})

            items = sorted(items, key=lambda k: k["y"])

            rows = []
            current_row = []
            tolerance = 15  # Khoảng cách sai số dòng

            for item in items:
              if not current_row:
                current_row.append(item)
              else:
                if abs(item["y"] - current_row[0]["y"]) < tolerance:
                  current_row.append(item)
                else:
                  current_row = sorted(current_row, key=lambda k: k["x"])
                  rows.append(current_row)
                  current_row = [item]
            if current_row:
              current_row = sorted(current_row, key=lambda k: k["x"])
              rows.append(current_row)

            # Tạo file Excel mới định dạng chuẩn
            wb = Workbook()
            ws = wb.active
            ws.title = "SongNgu_Table"

            thin_border = Border(
                left=Side(style="thin", color="B0B0B0"),
                right=Side(style="thin", color="B0B0B0"),
                top=Side(style="thin", color="B0B0B0"),
                bottom=Side(style="thin", color="B0B0B0"),
            )
            center_alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

            for r_idx, row_items in enumerate(rows, start=1):
              for c_idx, cell_item in enumerate(row_items, start=1):
                original_text = cell_item["text"]
                translated_text = translate_text(
                    original_text, translation_direction
                )

                # Đúng yêu cầu: Chữ tiếng Việt nằm ngay bên dưới chữ tiếng Trung trong cùng ô
                cell_value = f"{original_text}\n{translated_text}"

                cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                cell.border = thin_border
                cell.alignment = center_alignment
                cell.font = Font(name="Calibri", size=11)

            # Cân chỉnh độ rộng cột tự động
            for col in ws.columns:
              max_len = 0
              for cell in col:
                if cell.value:
                  lines = str(cell.value).split("\n")
                  max_len = max(max_len, max(len(l) for l in lines))
              col_letter = col[0].column_letter
              ws.column_dimensions[col_letter].width = max(max_len + 4, 16)

            output = io.BytesIO()
            wb.save(output)
            excel_data = output.getvalue()

            st.success("✨ Xử lý ảnh và tạo file Excel song ngữ thành công!")
            st.download_button(
                label="📥 Tải xuống File Excel kết quả",
                data=excel_data,
                file_name=f"Translated_Image_{uploaded_file.name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
          st.error(f"Đã xảy ra lỗi khi xử lý ảnh: {e}")

  # =========================================================
  # TRƯỜNG HỢP 2: XỬ LÝ FILE EXCEL (GIỮ NGUYÊN 100% ĐỊNH DẠNG GỐC)
  # =========================================================
  elif file_extension in ["xlsx", "xls"]:
    st.subheader("📊 Xem trước dữ liệu Excel gốc")
    try:
      df_preview = pd.read_excel(uploaded_file, sheet_name=0)
      st.dataframe(df_preview.head(10), use_container_width=True)

      if st.button("🚀 Bắt đầu dịch File Excel"):
        with st.spinner(
            "Đang dịch và bảo toàn toàn bộ định dạng gốc của file Excel..."
        ):
          uploaded_file.seek(0)
          # Sử dụng openpyxl để load và giữ nguyên 100% style, màu sắc, border, merge cell
          wb = load_workbook(uploaded_file)
          ws = wb.active

          # Duyệt qua mọi ô kể cả các ô được gộp (merged cells)
          for row in ws.iter_rows():
            for cell in row:
              if cell.value is not None:
                original_text = str(cell.value)
                if not original_text.startswith("="):
                  translated_text = translate_text(
                      original_text, translation_direction
                  )

                  # Đưa nội dung dịch nằm chung ngay bên dưới trong chính ô đó
                  cell.value = f"{original_text}\n{translated_text}"

                  # Bật tính năng xuống dòng tự động trong ô
                  cell.alignment = cell.alignment.copy(wrap_text=True)

          output = io.BytesIO()
          wb.save(output)
          processed_data = output.getvalue()

          st.success("✨ Dịch file Excel và bảo toàn 100% định dạng thành công!")
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
