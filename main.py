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
    """Đọc kích thước từ ghi chú trên ảnh bằng Gemini AI."""
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
# 2. XỬ LÝ TRÍCH XUẤT VECTOR NGUYÊN BẢN & ĐIỀU CHỈNH HOA VĂN + KHUNG SEPARATE
# ==============================================================================
def process_pattern_and_frame(
    image_bytes,
    target_w_mm=810.0,
    target_h_mm=2280.0,
    pattern_stroke_mm=70.0,
    frame_stroke_mm=70.0,
):
    """Trích xuất chính xác 100% vector dáng gốc, cho phép tùy chỉnh riêng độ rộng hoa văn và khung bao."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize (Tách phần phôi gỗ màu nâu ra khỏi nền)
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # 2. Loại bỏ hoàn toàn chữ OCR & mũi tên xung quanh bằng Bounding Box
    contours_external, _ = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    if not contours_external:
        return [], (thresh.shape[1], thresh.shape[0])

    main_contour = max(contours_external, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(main_contour)

    pattern_crop = thresh[y : y + h, x : x + w]

    scale_x = target_w_mm / w
    scale_y = target_h_mm / h
    avg_scale = (scale_x + scale_y) / 2.0

    # 3. Xử lý tùy chỉnh độ rộng nét hoa văn nếu khác mặc định 70mm
    base_stroke_mm = 70.0
    diff_pattern_mm = pattern_stroke_mm - base_stroke_mm

    if abs(diff_pattern_mm) > 1.0:
        kernel_size = int(abs(diff_pattern_mm) / avg_scale)
        if kernel_size > 0:
            if kernel_size % 2 == 0:
                kernel_size += 1
            kernel = cv2.getStructuringElement(
                cv2.MORPH_ELLIPSE, (kernel_size, kernel_size)
            )

            if diff_pattern_mm > 0:
                pattern_crop = cv2.dilate(pattern_crop, kernel, iterations=1)
            else:
                pattern_crop = cv2.erode(pattern_crop, kernel, iterations=1)

    # 4. Trích xuất toàn bộ cấu trúc đường viền (Viền khung ngoài + Các lỗ trống hoa văn)
    contours, hierarchy = cv2.findContours(
        pattern_crop, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )

    pattern_polygons = []

    # A. Dựng khung ngoài theo độ rộng frame_stroke_mm tùy chỉnh
    # Đường viền biên ngoài cùng phôi
    outer_frame = [
        (0.0, 0.0),
        (target_w_mm, 0.0),
        (target_w_mm, target_h_mm),
        (0.0, target_h_mm),
        (0.0, 0.0),
    ]

    # Đường viền lọt lòng trong của khung (bằng chiều rộng tổng trừ đi độ rộng khung 2 bên)
    inner_x1 = frame_stroke_mm
    inner_y1 = frame_stroke_mm
    inner_x2 = target_w_mm - frame_stroke_mm
    inner_y2 = target_h_mm - frame_stroke_mm

    if inner_x2 > inner_x1 and inner_y2 > inner_y1:
        inner_frame = [
            (inner_x1, inner_y1),
            (inner_x2, inner_y1),
            (inner_x2, inner_y2),
            (inner_x1, inner_y2),
            (inner_x1, inner_y1),
        ]
        pattern_polygons.append(outer_frame)
        pattern_polygons.append(inner_frame)

    # B. Dựng toàn bộ chi tiết hoa văn chuẩn dáng gốc
    if hierarchy is not None:
        for i, cnt in enumerate(contours):
            if cv2.contourArea(cnt) < 40:
                continue

            # Bỏ qua contour ngoài cùng của crop để dùng khung dựng chuẩn ở trên
            if hierarchy[0][i][3] == -1 and i == 0:
                continue

            epsilon = 0.001 * cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, epsilon, True)

            pts = []
            for p in approx:
                px, py = p[0][0], p[0][1]
                # Đảo trục Y cho AutoCAD
                pts.append((px * scale_x, (h - py) * scale_y))

            if len(pts) > 2:
                pts.append(pts[0])
                pattern_polygons.append(pts)

    return pattern_polygons, (w, h)


# ==============================================================================
# 3. XUẤT FILE DXF CẮT CNC
# ==============================================================================
def create_dxf_file(pattern_polygons):
    """Xuất file DXF chứa đầy đủ đường cắt chuẩn."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    doc.layers.add(name="CAT_CNC_HOA_VAN", color=1)

    for poly in pattern_polygons:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "CAT_CNC_HOA_VAN"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "Hoa_Van_Chuan_CNC.dxf")
    doc.saveas(filepath)

    with open(filepath, "rb") as f:
        dxf_bytes = f.read()

    return dxf_bytes


# ==============================================================================
# 4. GIAO DIỆN STREAMLIT WEB
# ==============================================================================
st.set_page_config(
    page_title="Auto CAD Vectorizer & Gemini OCR", layout="wide"
)
st.title("🌺 Hybrid AI Engine: Trích Xuất Vector Chuẩn Dáng Gốc 100%")

st.sidebar.header("⚙️ Cấu hình Hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

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

        # NÚT 1: Gemini AI đọc kích thước
        if st.button("🤖 NÚT AI: Đọc Kích Thước Bản Vẽ (Gemini OCR)"):
            if not api_key:
                st.error("Vui lòng nhập Gemini API Key ở Sidebar!")
            else:
                with st.spinner("🤖 Gemini AI đang quét kích thước..."):
                    w, h = extract_dimensions_with_gemini(img_bytes, api_key)
                    st.session_state.width_mm = w
                    st.session_state.height_mm = h
                    st.success(
                        f"🎉 AI đã đọc thành công: Rộng {w}mm x Cao {h}mm"
                    )

    with col2:
        st.subheader("📐 Thông Số Kích Thước & Tùy Chỉnh Nét")

        st.session_state.width_mm = st.number_input(
            "Chiều rộng phôi (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều cao phôi (mm):", value=st.session_state.height_mm
        )

        st.markdown("---")
        st.subheader("🎛️ Tùy Chỉnh Tách Biệt Nét")

        # 2 Ô NHẬP TÙY CHỈNH THEO YÊU CẦU
        pattern_stroke = st.number_input(
            "1. Độ rộng nét HOA VĂN (mm):",
            value=70.0,
            step=5.0,
            help="Chỉnh độ dày mỏng riêng cho các họa tiết hoa văn bên trong",
        )

        frame_stroke = st.number_input(
            "2. Độ rộng KHUNG BAO ngoài (mm):",
            value=70.0,
            step=5.0,
            help="Chỉnh độ rộng bản khung gỗ bao quanh ngoài",
        )

        st.markdown("---")

        # NÚT 2: Xuất DXF Vector
        if st.button("⚙️ NÚT CAD: Trích Xuất Vector & Tải File DXF"):
            with st.spinner(
                f"Đang xử lý vector (Hoa văn: {pattern_stroke}mm | Khung: {frame_stroke}mm)..."
            ):
                pattern_polygons, (w_px, h_px) = process_pattern_and_frame(
                    img_bytes,
                    st.session_state.width_mm,
                    st.session_state.height_mm,
                    pattern_stroke_mm=pattern_stroke,
                    frame_stroke_mm=frame_stroke,
                )

                dxf_bytes = create_dxf_file(pattern_polygons)

            st.success(
                f"✅ Đã trích xuất thành công **{len(pattern_polygons)}** đường cắt vector trùng khớp ảnh gốc!"
            )

            st.download_button(
                label=f"💾 TẢI FILE DXF (Hoa văn {int(pattern_stroke)}mm - Khung {int(frame_stroke)}mm)",
                data=dxf_bytes,
                file_name=f"Hoa_Van_CNC_{int(pattern_stroke)}mm_Khung_{int(frame_stroke)}mm.dxf",
                mime="application/dxf",
            )
