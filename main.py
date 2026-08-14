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
from shapely.geometry import LineString, MultiPolygon, Polygon
from shapely.ops import unary_union
import streamlit as st


# ==============================================================================
# 1. GEMINI OCR SCHEMA & PARSER
# ==============================================================================
class DrawingDimensions(BaseModel):
    width_mm: float = Field(
        description="Chiều rộng tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )
    height_mm: float = Field(
        description="Chiều cao tổng thể của sản phẩm đọc trên ghi chú (mm)"
    )


def extract_dimensions_with_gemini(image_bytes, api_key):
    """Sử dụng Gemini AI để đọc thông số kích thước từ ghi chú trên ảnh."""
    try:
        client = genai.Client(api_key=api_key)
        img = Image.open(io.BytesIO(image_bytes))

        prompt = """
        Hãy đọc các con số ghi chú kích thước tổng thể trên bức ảnh này:
        1. Chiều rộng tổng thể (Ví dụ: 810mm).
        2. Chiều cao tổng thể (Ví dụ: 2280mm).
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
        return min(w, h), max(w, h)
    except Exception as e:
        st.warning(
            f"⚠️ Không đọc được OCR ({e}). Dùng kích thước mặc định: 810mm x 2280mm."
        )
        return 810.0, 2280.0


# ==============================================================================
# 2. XỬ LÝ HÌNH HỌC CAD: ÉP NẾT 70MM SẠCH SẼ (KHÔNG LỖI VÒNG TRÒN ĐÈ NHAU)
# ==============================================================================
def process_pattern_exact_70mm(
    image_bytes, target_w_mm=810.0, target_h_mm=2280.0, stroke_width_mm=70.0
):
    """Trích xuất đường viền và sử dụng Hình học Shapely để tạo nét 70mm chuẩn CAD."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize ảnh
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # 2. Crop bỏ sạch chữ kích thước xung quanh bằng Bounding Box
    contours_all, _ = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    if not contours_all:
        return [], (thresh.shape[1], thresh.shape[0])

    main_contour = max(contours_all, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(main_contour)
    pattern_crop = thresh[y : y + h, x : x + w]

    scale_x = target_w_mm / w
    scale_y = target_h_mm / h

    # 3. Trích xuất các đường Contour ban đầu
    contours, _ = cv2.findContours(
        pattern_crop, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )

    lines_shapely = []
    half_stroke = stroke_width_mm / 2.0  # Lấy 35mm mỗi bên

    for cnt in contours:
        if cv2.contourArea(cnt) < 20:
            continue

        epsilon = 0.0015 * cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, epsilon, True)

        pts = []
        for p in approx:
            px, py = p[0][0], p[0][1]
            # Đảo trục Y cho CAD
            pts.append((px * scale_x, (h - py) * scale_y))

        if len(pts) > 2:
            pts.append(pts[0])
            line = LineString(pts)
            # Dùng join_style=2 (mitre) để KHÔNG TẠO BỌC TRÒN
            buffered = line.buffer(
                half_stroke, cap_style=2, join_style=2, mitre_limit=2.0
            )
            if buffered.is_valid and not buffered.is_empty:
                lines_shapely.append(buffered)

    # 4. HỢP NHẤT HÌNH HỌC (UNARY UNION) - Triệt tiêu hoàn toàn các nét đè chồng
    unified_geometry = unary_union(lines_shapely)

    pattern_polygons = []

    def extract_coords(geom):
        if isinstance(geom, Polygon):
            pattern_polygons.append(list(geom.exterior.coords))
            for interior in geom.interiors:
                pattern_polygons.append(list(interior.coords))
        elif isinstance(geom, MultiPolygon):
            for poly in geom.geoms:
                extract_coords(poly)

    extract_coords(unified_geometry)

    return pattern_polygons, (w, h)


# ==============================================================================
# 3. XUẤT FILE DXF CHUẨN CẮT CNC
# ==============================================================================
def create_dxf_file(pattern_polygons):
    """Xuất file DXF chứa đường cắt hoa văn đã hợp nhất chuẩn nét."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    doc.layers.add(name="HOA_VAN_70MM", color=1)  # Layer màu đỏ

    for poly in pattern_polygons:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "HOA_VAN_70MM"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "Hoa_Van_Fix_70mm.dxf")
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
st.title("🌺 Hybrid AI Engine: Ép Chuẩn Nét 70mm (Thuật Toán Shapely Union)")

st.sidebar.header("⚙️ Cấu hình Hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

fixed_stroke = st.sidebar.number_input(
    "Độ rộng nét hoa văn (mm):", value=70.0, step=5.0
)

if "width_mm" not in st.session_state:
    st.session_state.width_mm = 810.0
if "height_mm" not in st.session_state:
    st.session_state.height_mm = 2280.0

uploaded_file = st.file_uploader(
    "Nạp ảnh hoa văn CNC (JPG, PNG):", type=["png", "jpg", "jpeg"]
)

if uploaded_file:
    img_bytes = uploaded_file.getvalue()

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("🖼️ Ảnh Gốc Upload")
        st.image(img_bytes, use_container_width=True)

        if st.button("🤖 NÚT AI: Đọc Kích Thước Bản Vẽ (Gemini OCR)"):
            if not api_key:
                st.error("Vui lòng nhập Gemini API Key ở Sidebar!")
            else:
                with st.spinner("🤖 Gemini AI đang đọc kích thước..."):
                    w, h = extract_dimensions_with_gemini(img_bytes, api_key)
                    st.session_state.width_mm = w
                    st.session_state.height_mm = h
                    st.success(
                        f"🎉 AI đã đọc thành công: Rộng {w}mm x Cao {h}mm"
                    )

    with col2:
        st.subheader("📐 Bảng Điều Chỉnh & Xuất DXF")

        st.session_state.width_mm = st.number_input(
            "Chiều rộng phôi (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều cao phôi (mm):", value=st.session_state.height_mm
        )

        st.markdown("---")

        if st.button("⚙️ NÚT CAD: Hợp Nhất Shape & Xuất DXF 70mm"):
            with st.spinner(
                f"Đang tính toán hợp nhất hình học nét {fixed_stroke}mm..."
            ):
                pattern_polygons, (w_px, h_px) = process_pattern_exact_70mm(
                    img_bytes,
                    st.session_state.width_mm,
                    st.session_state.height_mm,
                    stroke_width_mm=fixed_stroke,
                )

                dxf_bytes = create_dxf_file(pattern_polygons)

            st.success("✅ Đã xử lý hợp nhất hình học thành công!")

            st.download_button(
                label="💾 TẢI FILE DXF CẮT CNC (SẠCH NẾT 70MM)",
                data=dxf_bytes,
                file_name=f"Hoa_Van_Clean_{int(fixed_stroke)}mm.dxf",
                mime="application/dxf",
            )
