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
# 2. XỬ LÝ TRÍCH XUẤT VECTOR HOA VĂN CHUẨN DÁNG NGUYÊN BẢN 100%
# ==============================================================================
def process_pattern_exact_original(
    image_bytes, target_w_mm=810.0, target_h_mm=2280.0
):
    """Trích xuất chính xác 100% đường viền hoa văn từ khối phôi màu nâu, không làm méo hay rách nét."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize (Tách phần phôi gỗ màu nâu ra khỏi nền trắng)
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # 2. Loại bỏ hoàn toàn chữ kích thước & mũi tên ở viền ngoài bằng Bounding Box
    contours_external, _ = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    if not contours_external:
        return [], (thresh.shape[1], thresh.shape[0])

    # Tìm khung gỗ lớn nhất
    main_contour = max(contours_external, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(main_contour)

    # Crop riêng phần phôi hoa văn gỗ
    pattern_crop = thresh[y : y + h, x : x + w]

    # Tính tỉ lệ quy đổi Pixel sang mm
    scale_x = target_w_mm / w
    scale_y = target_h_mm / h

    # 3. Trích xuất toàn bộ cấu trúc đường viền (Viền khung ngoài + Các lỗ trống hoa văn)
    # RETR_TREE giúp giữ nguyên phân cấp lòng trong / lòng ngoài chuẩn xác
    contours, hierarchy = cv2.findContours(
        pattern_crop, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )

    pattern_polygons = []
    if hierarchy is not None:
        for i, cnt in enumerate(contours):
            # Bỏ qua các vết chấm nhiễu rác nhỏ
            if cv2.contourArea(cnt) < 40:
                continue

            # Làm mượt đường cong hoa văn mượt mà
            epsilon = 0.001 * cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, epsilon, True)

            pts = []
            for p in approx:
                px, py = p[0][0], p[0][1]
                # Chuyển sang mm thực tế & Đảo trục Y cho AutoCAD
                pts.append((px * scale_x, (h - py) * scale_y))

            if len(pts) > 2:
                pts.append(pts[0])  # Khép kín đường polyline
                pattern_polygons.append(pts)

    return pattern_polygons, (w, h)


# ==============================================================================
# 3. XUẤT FILE DXF CẮT CNC
# ==============================================================================
def create_dxf_file(pattern_polygons):
    """Xuất file DXF chứa toàn bộ đường nét cắt chuẩn."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    # Tạo Layer chuẩn đỏ
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
st.title("🌺 Hybrid AI Engine: Trích Xuất Vector Hoa Văn CNC Nguyên Bản")

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
                with st.spinner("🤖 Gemini AI đang đọc kích thước..."):
                    w, h = extract_dimensions_with_gemini(img_bytes, api_key)
                    st.session_state.width_mm = w
                    st.session_state.height_mm = h
                    st.success(
                        f"🎉 AI đã đọc thành công: Rộng {w}mm x Cao {h}mm"
                    )

    with col2:
        st.subheader("📐 Thông Số Phôi & Xuất DXF")

        st.session_state.width_mm = st.number_input(
            "Chiều rộng phôi (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều cao phôi (mm):", value=st.session_state.height_mm
        )

        st.markdown("---")

        # NÚT 2: Xuất DXF Vector
        if st.button("⚙️ NÚT CAD: Trích Xuất Vector & Tải File DXF"):
            with st.spinner("Đang trích xuất vector dáng gốc chuẩn xác..."):
                pattern_polygons, (w_px, h_px) = process_pattern_exact_original(
                    img_bytes,
                    st.session_state.width_mm,
                    st.session_state.height_mm,
                )

                dxf_bytes = create_dxf_file(pattern_polygons)

            st.success(
                f"✅ Đã tạo thành công **{len(pattern_polygons)}** đường cắt vector trùng khớp ảnh gốc!"
            )

            st.download_button(
                label="💾 TẢI FILE DXF CẮT CNC (CHUẨN DÁNG GỐC)",
                data=dxf_bytes,
                file_name="Hoa_Van_CNC_ChuanNhat.dxf",
                mime="application/dxf",
            )
