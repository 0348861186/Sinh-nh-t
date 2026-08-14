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
    """Đọc kích thước bản vẽ tự động qua Gemini AI."""
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
# 2. XỬ LÝ TRÍCH XUẤT VECTOR & CHỈNH ĐỘ RỘNG HOA VĂN MƯỢT MÀ
# ==============================================================================
def process_pattern_with_custom_thickness(
    image_bytes,
    target_w_mm=810.0,
    target_h_mm=2280.0,
    stroke_width_mm=70.0,
):
    """Trích xuất hoa văn và tự động căn chỉnh độ rộng nét theo đúng thông số nhập vào (mm)."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize (Tách phần phôi gỗ màu nâu ra khỏi nền)
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # 2. Bỏ sạch chữ OCR & mũi tên xung quanh bằng Bounding Box
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

    # 3. Tự động vi điều chỉnh độ rộng nét theo thông số người dùng nhập (stroke_width_mm)
    # Ảnh gốc chuẩn đã có nét khoảng ~70mm. Nếu người dùng nhập thay đổi, OpenCV sẽ tự căn chỉnh kernel
    base_stroke_mm = 70.0
    diff_mm = stroke_width_mm - base_stroke_mm

    if abs(diff_mm) > 2.0:
        kernel_size_px = int(abs(diff_mm) / avg_scale)
        if kernel_size_px > 0:
            if kernel_size_px % 2 == 0:
                kernel_size_px += 1
            kernel = cv2.getStructuringElement(
                cv2.MORPH_ELLIPSE, (kernel_size_px, kernel_size_px)
            )

            if diff_mm > 0:
                # Phóng rộng nét nếu muốn nét dày hơn 70mm
                pattern_crop = cv2.dilate(pattern_crop, kernel, iterations=1)
            else:
                # Thu nhỏ nét nếu muốn nét mỏng hơn 70mm
                pattern_crop = cv2.erode(pattern_crop, kernel, iterations=1)

    # 4. Trích xuất đường viền chuẩn phân cấp lòng trong / lòng ngoài
    contours, hierarchy = cv2.findContours(
        pattern_crop, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )

    pattern_polygons = []
    if hierarchy is not None:
        for cnt in contours:
            if cv2.contourArea(cnt) < 40:
                continue

            epsilon = 0.001 * cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, epsilon, True)

            pts = []
            for p in approx:
                px, py = p[0][0], p[0][1]
                # Đảo trục Y cho hệ tọa độ CAD
                pts.append((px * scale_x, (h - py) * scale_y))

            if len(pts) > 2:
                pts.append(pts[0])
                pattern_polygons.append(pts)

    return pattern_polygons, (w, h)


# ==============================================================================
# 3. XUẤT FILE DXF CẮT CNC
# ==============================================================================
def create_dxf_file(pattern_polygons):
    """Xuất file DXF chứa đầy đủ đường cắt hoa văn."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    doc.layers.add(name="CAT_CNC_HOA_VAN", color=1)

    for poly in pattern_polygons:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "CAT_CNC_HOA_VAN"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "Hoa_Van_Chuan_DXF.dxf")
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
st.title("🌺 Hybrid AI Engine: Trích Xuất Vector & Chỉnh Độ Rộng Nét Hoa Văn")

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
        st.subheader("📐 Bảng Tùy Chỉnh Kích Thước & Độ Rộng Nét")

        st.session_state.width_mm = st.number_input(
            "Chiều rộng phôi (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều cao phôi (mm):", value=st.session_state.height_mm
        )

        # 🎯 Ô NHẬP ĐỘ RỘNG NẾT HOA VĂN BẠN YÊU CẦU
        custom_stroke_mm = st.number_input(
            "Độ rộng nét hoa văn & khung (mm):",
            value=70.0,
            step=5.0,
            help="Mặc định là 70mm theo bản vẽ. Bạn có thể tăng/giảm tùy ý.",
        )

        st.markdown("---")

        # NÚT 2: Trích xuất Vector & Xuất DXF
        if st.button("⚙️ NÚT CAD: Trích Xuất Vector & Tải File DXF"):
            with st.spinner(
                f"Đang xử lý vector với độ rộng nét {custom_stroke_mm}mm..."
            ):
                pattern_polygons, (w_px, h_px) = (
                    process_pattern_with_custom_thickness(
                        img_bytes,
                        st.session_state.width_mm,
                        st.session_state.height_mm,
                        stroke_width_mm=custom_stroke_mm,
                    )
                )

                dxf_bytes = create_dxf_file(pattern_polygons)

            st.success(
                f"✅ Đã tạo file DXF nét **{custom_stroke_mm}mm** thành công với **{len(pattern_polygons)}** đường nét chuẩn!"
            )

            st.download_button(
                label=f"💾 TẢI FILE DXF CẮT CNC (NẾT {int(custom_stroke_mm)}MM)",
                data=dxf_bytes,
                file_name=f"Hoa_Van_Net_{int(custom_stroke_mm)}mm.dxf",
                mime="application/dxf",
            )
