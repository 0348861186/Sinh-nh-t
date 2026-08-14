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
from shapely.geometry import LineString
from skimage.morphology import skeletonize
import streamlit as st


# ==============================================================================
# 1. GEMINI OCR SCHEMA & PARSER (ĐỌC KÍCH THƯỚC BẢN VẼ)
# ==============================================================================
class DrawingDimensions(BaseModel):
    width_mm: float = Field(
        description="Chiều rộng tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )
    height_mm: float = Field(
        description="Chiều cao tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )


def extract_dimensions_with_gemini(image_bytes, api_key):
    """Sử dụng Gemini AI để đọc thông số kích thước cơ khí từ ghi chú trên ảnh."""
    try:
        client = genai.Client(api_key=api_key)
        img = Image.open(io.BytesIO(image_bytes))

        prompt = """
        Bạn là chuyên gia đọc bản vẽ cơ khí OCR.
        Hãy đọc các con số ghi chú kích thước tổng thể trên bức ảnh này:
        1. Chiều rộng tổng thể (Ví dụ: 810mm hoặc 2280mm).
        2. Chiều cao tổng thể (Ví dụ: 2280mm hoặc 810mm).
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
        w = data.get("width_mm", 810.0)
        h = data.get("height_mm", 2280.0)

        # Tự động quy đổi cạnh ngắn là Rộng, cạnh dài là Cao
        return min(w, h), max(w, h)
    except Exception as e:
        st.warning(
            f"⚠️ Không đọc được OCR ({e}). Tự động dùng kích thước mặc định: 810mm x 2280mm."
        )
        return 810.0, 2280.0


# ==============================================================================
# 2. XỬ LÝ HOA VĂN: RÚT XƯƠNG & ÉP ĐỘ RỘNG HOÀN TOÀN VỀ 70MM
# ==============================================================================
def process_pattern_with_fixed_thickness(
    image_bytes, target_w_mm, target_h_mm, stroke_width_mm=70.0
):
    """Rút xương (Skeletonize) hoa văn và nới rộng đều 2 bên 35mm để toàn bộ nét đạt đúng 70mm."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize (Nhi phân hóa ảnh)
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # Khử nhiễu ảnh
    kernel = np.ones((3, 3), np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel, iterations=1)

    h_px, w_px = thresh.shape
    scale_x = target_w_mm / w_px
    scale_y = target_h_mm / h_px

    # 2. Thuật toán Rút xương (Biến các nét dày thành đường nét đơn 1-pixel ở chính giữa)
    binary_bool = thresh > 0
    skeleton = skeletonize(binary_bool)
    skeleton_uint8 = (skeleton * 255).astype(np.uint8)

    # 3. Trích xuất đường Centerlines từ xương
    contours, _ = cv2.findContours(
        skeleton_uint8, cv2.RETR_LIST, cv2.CHAIN_APPROX_TC89_KCOS
    )

    offset_distance = stroke_width_mm / 2.0  # Lấy 35mm về mỗi bên
    pattern_polygons = []

    for cnt in contours:
        if len(cnt) < 3:
            continue

        # Làm mượt vết răng cưa
        epsilon = 0.002 * cv2.arcLength(cnt, False)
        approx = cv2.approxPolyDP(cnt, epsilon, False)

        pts = []
        for p in approx:
            x_px, y_px = p[0][0], p[0][1]
            # Đảo trục Y cho hệ tọa độ CAD
            pts.append((x_px * scale_x, (h_px - y_px) * scale_y))

        if len(pts) >= 2:
            line = LineString(pts)
            # 4. Shapely Buffer: Nới rộng đường xương đúng 35mm mỗi bên -> Tổng nét uốn = 70mm
            buffered_poly = line.buffer(
                offset_distance, cap_style=1, join_style=1
            )

            if buffered_poly.is_valid and not buffered_poly.is_empty:
                if buffered_poly.geom_type == "Polygon":
                    coords = list(buffered_poly.exterior.coords)
                    pattern_polygons.append(coords)
                elif buffered_poly.geom_type == "MultiPolygon":
                    for poly in buffered_poly.geoms:
                        coords = list(poly.exterior.coords)
                        pattern_polygons.append(coords)

    return pattern_polygons, (w_px, h_px)


# ==============================================================================
# 3. XUẤT FILE DXF CHUẨN CẮT CNC (CÓ KHUNG & HOA VĂN 70MM)
# ==============================================================================
def add_exact_frame_to_dxf(
    msp, total_w_mm=810.0, total_h_mm=2280.0, frame_thick_mm=70.0
):
    """Vẽ 2 đường khung chữ nhật bao ngoài dày đúng 70mm chuẩn kỹ thuật."""
    outer_rect = [
        (0, 0),
        (total_w_mm, 0),
        (total_w_mm, total_h_mm),
        (0, total_h_mm),
        (0, 0),
    ]

    inner_rect = [
        (frame_thick_mm, frame_thick_mm),
        (total_w_mm - frame_thick_mm, frame_thick_mm),
        (total_w_mm - frame_thick_mm, total_h_mm - frame_thick_mm),
        (frame_thick_mm, total_h_mm - frame_thick_mm),
        (frame_thick_mm, frame_thick_mm),
    ]

    msp.add_lwpolyline(outer_rect, dxfattribs={"layer": "KHUNG_NGOAI_70MM"})
    msp.add_lwpolyline(inner_rect, dxfattribs={"layer": "KHUNG_NGOAI_70MM"})


def create_dxf_file(
    pattern_polygons, total_w_mm, total_h_mm, stroke_width_mm=70.0
):
    """Xuất file DXF chứa cả khung và toàn bộ hoa văn nét 70mm."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    # Tạo các Layer riêng biệt để máy CNC dễ phân biệt đường cắt
    doc.layers.add(name="HOA_VAN_70MM", color=1)  # Đỏ (Hoa văn)
    doc.layers.add(name="KHUNG_NGOAI_70MM", color=5)  # Xanh dương (Khung viền)

    # 1. Vẽ khung chữ nhật chuẩn 70mm
    add_exact_frame_to_dxf(
        msp, total_w_mm, total_h_mm, frame_thick_mm=stroke_width_mm
    )

    # 2. Vẽ hoa văn đã ép độ rộng 70mm
    for poly in pattern_polygons:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "HOA_VAN_70MM"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "Hoa_Van_Full_70mm.dxf")
    doc.saveas(filepath)

    with open(filepath, "rb") as f:
        dxf_bytes = f.read()

    return dxf_bytes


# ==============================================================================
# 4. GIAO DIỆN WEB STREAMLIT
# ==============================================================================
st.set_page_config(
    page_title="Auto CAD Vectorizer & Gemini OCR", layout="wide"
)
st.title("🌺 Hybrid AI Engine: Ép Toàn Bộ Hoa Văn & Khung Về Độ Rộng 70mm")

st.sidebar.header("⚙️ Cấu hình Hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

# Cho phép tùy chỉnh độ rộng nét nếu muốn
fixed_stroke_width = st.sidebar.number_input(
    "Độ rộng Toàn bộ Nét & Khung (mm):", value=70.0, step=5.0
)

uploaded_file = st.file_uploader(
    "Nạp ảnh hoa văn CNC (JPG, PNG):", type=["png", "jpg", "jpeg"]
)

if uploaded_file:
    img_bytes = uploaded_file.getvalue()

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("🖼️ Ảnh Gốc Upload")
        st.image(img_bytes, use_container_width=True)

    if st.button("🚀 Bắt đầu Phân tích & Xuất DXF 70mm"):
        if not api_key:
            st.error("Vui lòng nhập Gemini API Key ở thanh bên góc trái!")
        else:
            with st.spinner("1/2: Gemini AI đang đọc kích thước bản vẽ..."):
                width_mm, height_mm = extract_dimensions_with_gemini(
                    img_bytes, api_key
                )

            st.success(
                f"✅ Kích thước phôi: **{width_mm}mm x {height_mm}mm** | Độ rộng nét ép: **{fixed_stroke_width}mm**"
            )

            with st.spinner(
                f"2/2: Đang rút xương hoa văn & ép độ rộng nét về đúng {fixed_stroke_width}mm..."
            ):
                pattern_polygons, (w_px, h_px) = (
                    process_pattern_with_fixed_thickness(
                        img_bytes,
                        width_mm,
                        height_mm,
                        stroke_width_mm=fixed_stroke_width,
                    )
                )

                dxf_bytes = create_dxf_file(
                    pattern_polygons,
                    width_mm,
                    height_mm,
                    stroke_width_mm=fixed_stroke_width,
                )

            with col2:
                st.subheader("📐 Thông Số DXF Chuẩn Cắt CNC")
                st.info(
                    f"📌 File DXF đã được chuẩn hóa:\n"
                    f"- **Layer `KHUNG_NGOAI_70MM`**: Khung viền dày đúng {fixed_stroke_width}mm.\n"
                    f"- **Layer `HOA_VAN_70MM`**: Toàn bộ đường uốn bên trong đã được nới rộng chính xác {fixed_stroke_width}mm."
                )
                st.metric("Tổng số cụm hoa văn đã xử lý", len(pattern_polygons))

                st.download_button(
                    label="💾 TẢI FILE DXF CẮT CNC (FULL NÉT 70MM)",
                    data=dxf_bytes,
                    file_name=f"Hoa_Van_Full_{int(fixed_stroke_width)}mm.dxf",
                    mime="application/dxf",
                )
