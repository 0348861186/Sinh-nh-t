import io
import json
import math
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

# Thử import potrace, nếu hệ thống chưa cài C-extension sẽ fallback sang OpenCV Contour
try:
    import potrace

    HAS_POTRACE = True
except ImportError:
    HAS_POTRACE = False


# ==============================================================================
# 1. GEMINI OCR SCHEMA & PARSER (CHỈ ĐỌC KÍCH THƯỚC THỰC TẾ)
# ==============================================================================
class DrawingDimensions(BaseModel):
    width_mm: float = Field(
        description="Chiều rộng tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )
    height_mm: float = Field(
        description="Chiều cao tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )
    frame_thickness_mm: float = Field(
        default=70.0, description="Độ dày khung bao nếu có (mm)"
    )


def extract_dimensions_with_gemini(image_bytes, api_key):
    """Sử dụng Gemini AI để đọc thông số kích thước cơ khí từ ghi chú trên ảnh."""
    try:
        client = genai.Client(api_key=api_key)
        img = Image.open(io.BytesIO(image_bytes))

        prompt = """
        Bạn là chuyên gia đọc bản vẽ cơ khí OCR.
        Hãy đọc các con số ghi chú kích thước tổng thể trên bức ảnh này:
        1. Tìm chiều rộng tổng thể (Ví dụ: 810mm).
        2. Tìm chiều cao tổng thể (Ví dụ: 2280mm).
        3. Tìm độ dày khung biên nếu có (Ví dụ: 70mm).
        Chỉ trả về dữ liệu đúng chuẩn JSON Schema.
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
        return data.get("width_mm", 810.0), data.get("height_mm", 2280.0)
    except Exception as e:
        st.warning(f"⚠️ Không đọc được kích thước qua Gemini ({e}). Sử dụng kích thước mặc định 810x2280mm.")
        return 810.0, 2280.0


# ==============================================================================
# 2. VECTOR ENGINE (OPENCV + POTRACE): TẠO ĐƯỜNG CONG TỰ ĐỘNG
# ==============================================================================
def process_and_vectorize(image_bytes, target_w_mm, target_h_mm):
    """Tách hoa văn bằng OpenCV và trích xuất Vector mượt bằng Potrace/Contours."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarization (Nhi phân hóa ảnh đen trắng)
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # Khử nhiễu ảnh
    kernel = np.ones((3, 3), np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)

    h_px, w_px = thresh.shape

    # 2. Tính tỉ lệ quy đổi Pixel -> mm
    scale_x = target_w_mm / w_px
    scale_y = target_h_mm / h_px

    polylines = []

    # 3. Trích xuất đường nét Vector
    if HAS_POTRACE:
        # Phương án tối ưu: Dùng Potrace lấy đường cong mịn
        bmp = potrace.Bitmap(thresh)
        path = bmp.trace(2, potrace.POTRACE_TURNPOLICY_MINORITY, 1.0, 0.5, 1)

        for curve in path:
            pts = []
            start = curve.start_point
            pts.append((start[0] * scale_x, (h_px - start[1]) * scale_y))

            for segment in curve.segments:
                if segment.is_corner:
                    c = segment.c
                    end = segment.end_point
                    pts.append((c[0] * scale_x, (h_px - c[1]) * scale_y))
                    pts.append((end[0] * scale_x, (h_px - end[1]) * scale_y))
                else:
                    # Chuyển Bezier thành các đoạn điểm nội suy mượt
                    c1 = segment.c1
                    c2 = segment.c2
                    end = segment.end_point
                    for t in np.linspace(0, 1, 10):
                        bx = (
                            (1 - t) ** 3 * start[0]
                            + 3 * (1 - t) ** 2 * t * c1[0]
                            + 3 * (1 - t) * t**2 * c2[0]
                            + t**3 * end[0]
                        )
                        by = (
                            (1 - t) ** 3 * start[1]
                            + 3 * (1 - t) ** 2 * t * c1[1]
                            + 3 * (1 - t) * t**2 * c2[1]
                            + t**3 * end[1]
                        )
                        pts.append((bx * scale_x, (h_px - by) * scale_y))
                    start = end
            if len(pts) > 2:
                polylines.append(pts)
    else:
        # Phương án Fallback: Dùng OpenCV Contours + Ramer-Douglas-Peucker
        contours, _ = cv2.findContours(
            thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
        )
        for cnt in contours:
            # Làm mượt vết gấp rắc rắc
            epsilon = 0.002 * cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, epsilon, True)

            pts = []
            for p in approx:
                x_px, y_px = p[0][0], p[0][1]
                # Đảo trục Y vì hệ tọa độ ảnh ngược với hệ tọa độ CAD/CAM
                pts.append((x_px * scale_x, (h_px - y_px) * scale_y))

            if len(pts) > 2:
                pts.append(pts[0])  # Khép kín đường vẽ
                polylines.append(pts)

    return polylines, (w_px, h_px)


# ==============================================================================
# 3. DXF EXPORTER (XUẤT FILE DXF CHUẨN CAD/CAM)
# ==============================================================================
def create_dxf_file(polylines):
    """Xuất danh sách polylines thành file DXF chuẩn."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    # Tạo Layer riêng cho cắt CNC
    doc.layers.add(name="CNC_CUT", color=1)  # Đỏ

    for poly in polylines:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "CNC_CUT"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "output_pattern.dxf")
    doc.saveas(filepath)

    with open(filepath, "rb") as f:
        dxf_bytes = f.read()

    return dxf_bytes


# ==============================================================================
# 4. GIAO DIỆN PHẦN MỀM STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Auto CAD Vectorizer & Gemini OCR", layout="wide"
)
st.title("🌺 Hybrid AI Engine: Chuyển Ảnh Hoa Văn CNC Sang DXF Kỹ Thuật")

st.sidebar.header("⚙️ Cấu hình Hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

uploaded_file = st.file_uploader(
    "Nạp ảnh hoa văn CNC (JPG, PNG):", type=["png", "jpg", "jpeg"]
)

if uploaded_file:
    img_bytes = uploaded_file.getvalue()

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("🖼️ Ảnh Gốc Phác Thảo")
        st.image(img_bytes, use_container_width=True)

    if st.button("🚀 Bắt đầu Phân tích & Vectorize DXF"):
        if not api_key:
            st.error("Vui lòng nhập Gemini API Key ở thanh bên góc trái!")
        else:
            with st.spinner(
                "1/2: Gemini AI đang đọc thước đo kích thước OCR..."
            ):
                width_mm, height_mm = extract_dimensions_with_gemini(
                    img_bytes, api_key
                )

            st.success(
                f"✅ Gemini đã nhận diện kích thước: **{width_mm}mm x {height_mm}mm**"
            )

            with st.spinner(
                "2/2: OpenCV & Potrace đang trích xuất đường cong Vector..."
            ):
                polylines, (w_px, h_px) = process_and_vectorize(
                    img_bytes, width_mm, height_mm
                )
                dxf_bytes = create_dxf_file(polylines)

            with col2:
                st.subheader("📐 Kết quả Vectorized & Thông số")
                st.metric("Tổng số đường nét (Contours)", len(polylines))
                st.metric(
                    "Độ phân giải ảnh gốc", f"{w_px} x {h_px} Pixels"
                )
                st.metric(
                    "Tỉ lệ quy đổi X", f"{width_mm / w_px:.4f} mm/pixel"
                )
                st.metric(
                    "Tỉ lệ quy đổi Y", f"{height_mm / h_px:.4f} mm/pixel"
                )

                st.download_button(
                    label="💾 TẢI FILE DXF CẮT CNC (KÍCH THƯỚC CHUẨN)",
                    data=dxf_bytes,
                    file_name="Hoa_Van_CNC_Precision.dxf",
                    mime="application/dxf",
                )
