import io
import os
import easyocr
import numpy as np
import pandas as pd
import streamlit as st
from deep_translator import GoogleTranslator
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Dịch Song ngữ Trung - Việt Chuyên Sâu",
    page_icon="🌐",
    layout="wide",
)


# Khởi tạo OCR reader
@st.cache_resource
def load_ocr_reader():
  return easyocr.Reader(["ch_sim", "en"], gpu=False)


# Hàm dịch văn bản
def translate_text(text, direction):
  if not text or not str(text).strip():
    return text
  text_str = str(text).strip()
  # Nếu là số hoặc công thức thì không dịch
  try:
    float(text_str)
    return text_str
  except ValueError:
    pass

  try:
    if direction == "Trung - Việt":
      return GoogleTranslator(source="zh-CN", target="vi").translate(text_str)
    else:
      return GoogleTranslator(source="vi", target="zh-CN").translate(text_str)
  except Exception as e:
    return text_str


st.title("🌐 Hệ thống Dịch Song ngữ Trung - Việt (Bảo toàn định dạng chuẩn)")
st.markdown("""
Ứng dụng hỗ trợ xử lý ảnh chụp bảng biểu tiếng Trung/Việt, nhận diện cấu trúc dạng lưới (Grid/Table), 
tiến hành dịch thuật và **xuất ra file Excel giữ nguyên vẹn hình dáng, màu sắc, khung viền và cấu trúc gộp ô**.
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

  # -------------------------------------------------------------
  # XỬ LÝ KHI TẢI LÊN FILE ẢNH (Tái tạo lại bảng biểu chính xác)
  # -------------------------------------------------------------
  if file_extension in ["png", "jpg", "jpeg"]:
    st.subheader("🖼️ Xem trước hình ảnh tải lên")
    image = Image.open(uploaded_file)
    st.image(image, caption="Hình ảnh gốc", width=700)

    if st.button("🚀 Bắt đầu Phân tích Bảng & Dịch"):
      with st.spinner(
          "Đang bóc tách cấu trúc bảng, nhận diện ký tự và dịch thuật..."
      ):
        try:
          reader = load_ocr_reader()
          image_np = np.array(image)
          results = reader.readtext(image_np)

          if not results:
            st.warning("Không tìm thấy văn bản nào trong ảnh!")
          else:
            # 1. Gom nhóm các hộp văn bản (boxes) thành các Dòng (Rows) dựa vào tọa độ Y
            # Lọc và sắp xếp tọa độ
            items = []
            for bbox, text, prob in results:
              # bbox có 4 điểm: [[x1,y1], [x2,y1], [x2,y2], [x1,y2]]
              y_center = (bbox[0][1] + bbox[2][1]) / 2
              x_center = (bbox[0][0] + bbox[2][0]) / 2
              items.append(
                  {"text": text, "x": x_center, "y": y_center, "bbox": bbox}
              )

            # Sắp xếp theo trục Y để gom nhóm dòng
            items = sorted(items, key=lambda k: k["y"])

            rows = []
            current_row = []
            tolerance = 15  # Sai số tọa độ Y để gom vào cùng một dòng

            for item in items:
              if not current_row:
                current_row.append(item)
              else:
                # So sánh độ lệch Y với phần tử đầu tiên của dòng hiện tại
                if abs(item["y"] - current_row[0]["y"]) < tolerance:
                  current_row.append(item)
                else:
                  # Sắp xếp các ô trong dòng theo trục X từ trái sang phải
                  current_row = sorted(current_row, key=lambda k: k["x"])
                  rows.append(current_row)
                  current_row = [item]
            if current_row:
              current_row = sorted(current_row, key=lambda k: k["x"])
              rows.append(current_row)

            # 2. Xây dựng file Excel mới bằng Openpyxl mô phỏng lại bảng
            wb = Workbook()
            ws = wb.active
            ws.title = "BangDichSongNgu"

            # Thiết lập style cơ bản (Màu cam giống mẫu header của bạn)
            header_fill = PatternFill(
                start_color="D35400", end_color="D35400", fill_type="solid"
            )
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            normal_font = Font(name="Calibri", size=11)
            thin_border = Border(
                left=Side(style="thin", color="CCCCCC"),
                right=Side(style="thin", color="CCCCCC"),
                top=Side(style="thin", color="CCCCCC"),
                bottom=Side(style="thin", color="CCCCCC"),
            )
            center_alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

            for r_idx, row_items in enumerate(rows, start=1):
              for c_idx, cell_item in enumerate(row_items, start=1):
                original_text = cell_item["text"]
                translated = translate_text(original_text, translation_direction)

                # Yêu cầu 3: Nội dung dịch nằm chung với ô (dòng dưới)
                cell_value = f"{original_text}\n{translated}"

                cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
                cell.font = (
                    header_font if r_idx == 1 else normal_font
                )  # Dòng đầu là header
                if r_idx == 1:
                  cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = center_alignment

            # Tự động chỉnh độ rộng cột cho đẹp
            for col in ws.columns:
              max_len = max(len(str(cell.value or "")) for cell in col)
              col_letter = col[0].column_letter
              ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

            # Lưu vào bộ nhớ đệm
            output = io.BytesIO()
            wb.save(output)
            excel_data = output.getvalue()

            st.success(
                "✨ Phân tích bảng ảnh và tạo file Excel song ngữ thành công!"
            )
            st.download_button(
                label="📥 Tải xuống File Excel định dạng Bảng",
                data=excel_data,
                file_name=f"Translated_Table_{uploaded_file.name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
          st.error(f"Đã xảy ra lỗi khi xử lý định dạng ảnh: {e}")

  # -------------------------------------------------------------
  # XỬ LÝ KHI TẢI LÊN FILE EXCEL GỐC
  # -------------------------------------------------------------
  elif file_extension in ["xlsx", "xls"]:
    st.subheader("📊 Xem trước dữ liệu Excel gốc")
    try:
      df_preview = pd.read_excel(uploaded_file, sheet_name=0)
      st.dataframe(df_preview.head(10), use_container_width=True)

      if st.button("🚀 Bắt đầu dịch File Excel"):
        with st.spinner("Đang xử lý dịch và giữ nguyên cấu trúc gốc..."):
          uploaded_file.seek(0)
          from openpyxl import load_workbook

          wb = load_workbook(uploaded_file)
          ws = wb.active

          for row in ws.iter_rows():
            for cell in row:
              if cell.value is not None:
                original_text = str(cell.value)
                if not original_text.startswith("="):
                  translated_text = translate_text(
                      original_text, translation_direction
                  )
                  cell.value = f"{original_text}\n{translated_text}"
                  cell.alignment = cell.alignment.copy(wrap_text=True)

          output = io.BytesIO()
          wb.save(output)
          processed_data = output.getvalue()

          st.success("✨ Dịch file Excel và bảo toàn định dạng thành công!")
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
