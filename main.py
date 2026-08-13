import io
import json
import os
import tempfile
import cv2
import ezdxf
from google import genai
from google.genai import types
import numpy as np
from PIL import Image
from pydantic import BaseModel, Field
import streamlit as st


# ==============================================================================
# 1. GEMINI OCR SCHEMA & PARSER (ĐỌC KÍCH THƯỚC)
# ==============================================================================
class DrawingDimensions(BaseModel):
    width_mm: float = Field(
        description="Chiều rộng tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )
    height_mm: float = Field(
        description="Chiều cao tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )
    frame_thickness_mm: float = Field(
        default=70.0, description="Độ dày khung bao (mm)"
    )


def extract_dimensions_with_gemini(image_bytes, api_key):
    """Đọc thông số kích thước từ ảnh qua Gemini AI."""
    try:
        client = genai.Client(api_key=api_key)
        img = Image.open(io.BytesIO(image_bytes))

        prompt = """
        Hãy đọc các con số ghi chú kích thước trên bức ảnh này:
        1. Chiều rộng tổng thể (Ví dụ: 810mm hoặc 2280mm).
        2. Chiều cao tổng thể (Ví dụ: 2280mm hoặc 810mm).
        3. Độ dày khung bao (Ví dụ: 70mm).
        Trả về đúng định dạng JSON Schema.
        """

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[img, prompt],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=DrawingDimensions,
            ),
        )

        data = json.loads(response.text.strip())
        w = data.get("width_mm", 810.0)
        h = data.get("height_mm", 2280.0)
        f = data.get("frame_thickness_mm", 70.0)

        # Tự động điều chỉnh chiều ngang/dọc theo thực tế
        return min(w, h), max(w, h), f
    except Exception as e:
        st.warning(
            f"⚠️ Không đọc được OCR ({e}). Dùng mặc định: 810mm x 2280mm, khung 70mm."
        )
        return 810.0, 2280.0, 70.0


# ==============================================================================
# 2. VECTOR ENGINE (OPENCV CONTOURS)
# ==============================================================================
def process_and_vectorize(image_bytes, target_w_mm, target_h_mm):
    """Trích xuất hoa văn bên trong bằng OpenCV."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Binarization
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # Khử nhiễu
    kernel = np.ones((3, 3), np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)

    h_px, w_px = thresh.shape

    # Tỉ lệ quy đổi Pixel -> mm
    scale_x = target_w_mm / w_px
    scale_y = target_h_mm / h_px

    # Tìm Contours hoa văn
    contours, _ = cv2.findContours(
        thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )
    polylines = []

    for cnt in contours:
        epsilon = 0.0015 * cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, epsilon, True)

        pts = []
        for p in approx:
            x_px, y_px = p[0][0], p[0][1]
            # Đảo trục Y cho chuẩn tọa độ CAD
            pts.append((x_px * scale_x, (h_px - y_px) * scale_y))

        if len(pts) > 2:
            pts.append(pts[0])
            polylines.append(pts)

    return polylines, (w_px, h_px)


# ==============================================================================
# 3. DXF EXPORTER (BỔ SUNG KHUNG CHUẨN 70MM)
# ==============================================================================
def add_exact_frame_to_dxf(
    msp, total_w_mm=810.0, total_h_mm=2280.0, frame_thick_mm=70.0
):
    """Vẽ 2 đường khung chữ nhật chuẩn mm vào Layer riêng."""
    # Khung ngoài cùng
    outer_rect = [
        (0, 0),
        (total_w_mm, 0),
        (total_w_mm, total_h_mm),
        (0, total_h_mm),
        (0, 0),
    ]

    # Khung trong (lùi vào đúng 70mm mỗi cạnh)
    inner_rect = [
        (frame_thick_mm, frame_thick_mm),
        (total_w_mm - frame_thick_mm, frame_thick_mm),
        (total_w_mm - frame_thick_mm, total_h_mm - frame_thick_mm),
        (frame_thick_mm, total_h_mm - frame_thick_mm),
        (frame_thick_mm, frame_thick_mm),
    ]

    msp.add_lwpolyline(outer_rect, dxfattribs={"layer": "CNC_FRAME_70MM"})
    msp.add_lwpolyline(inner_rect, dxfattribs={"layer": "CNC_FRAME_70MM"})


def create_dxf_file(
    polylines, total_w_mm, total_h_mm, frame_thick_mm=70.0
):
    """Xuất file DXF chứa cả hoa văn và khung 70mm."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    # Tạo Layer
    doc.layers.add(name="CNC_PATTERN", color=1)  # Đỏ (Hoa văn)
    doc.layers.add(name="CNC_FRAME_70MM", color=5)  # Xanh dương (Khung 70mm)

    # 1. Vẽ khung chuẩn 70mm
    add_exact_frame_to_dxf(msp, total_w_mm, total_h_mm, frame_thick_mm)

    # 2. Vẽ hoa văn
    for poly in polylines:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "CNC_PATTERN"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "output_cnc_pattern.dxf")
    doc.saveas(filepath)

    with open(filepath, "rb") as f:
        dxf_bytes = f.read()

    return dxf_bytes


# ==============================================================================
# 4. STREAMLIT APP UI
# ==============================================================================
st.set_page_config(
    page_title="Auto CAD Vectorizer & Gemini OCR", layout="wide"
)
st.title("🌺 Hybrid AI Engine: Xuất File DXF Cắt CNC (Khung Chuẩn 70mm)")

st.sidebar.header("⚙️ Cấu hình Hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

# Cho phép chỉnh độ dày khung thủ công nếu muốn
frame_thickness = st.sidebar.number_input(
    "Độ dày khung viền (mm):", value=70.0, step=5.0
)

uploaded_file = st.file_uploader(
    "Nạp ảnh hoa văn CNC (JPG, PNG):", type=["png", "jpg", "jpeg"]
)

if uploaded_file:
    img_bytes = uploaded_file.getvalue()

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("🖼️ Ảnh Gốc Phác Thảo")
        st.image(img_bytes, use_container_width=True)

    if st.button("🚀 Bắt đầu Phân tích & Xuất DXF"):
        if not api_key:
            st.error("Vui lòng nhập Gemini API Key ở Sidebar!")
        else:
            with st.spinner("1/2: Đang đọc kích thước bằng Gemini OCR..."):
                width_mm, height_mm, _ = extract_dimensions_with_gemini(
                    img_bytes, api_key
                )

            st.success(
                f"✅ Kích thước nhận diện: **{width_mm}mm x {height_mm}mm** | Khung viền: **{frame_thickness}mm**"
            )

            with st.spinner("2/2: Đang tạo Vector hoa văn & Khung 70mm..."):
                polylines, (w_px, h_px) = process_and_vectorize(
                    img_bytes, width_mm, height_mm
                )

                # Xuất DXF truyền vào thêm thông số khung
                dxf_bytes = create_dxf_file(
                    polylines, width_mm, height_mm, frame_thickness
                )

            with col2:
                st.subheader("📐 Kết quả Xuất DXF")
                st.info(
                    "📌 File DXF xuất ra sẽ có 2 Layer:\n"
                    "- `CNC_FRAME_70MM`: Khung chữ nhật ngoài & trong cách nhau đúng 70mm.\n"
                    "- `CNC_PATTERN`: Các đường vector hoa văn uốn bên trong."
                )
                st.metric("Tổng số đường nét hoa văn", len(polylines))

                st.download_button(
                    label="💾 TẢI FILE DXF CẮT CNC (CÓ KHUNG 70MM)",
                    data=dxf_bytes,
                    file_name="Hoa_Van_Khung_70mm.dxf",
                    mime="application/dxf",
                )
