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
# 2. XỬ LÝ TRÍCH XUẤT VECTOR & ÉP CHÍNH XÁC ĐỘ RỘNG NẾT (MM)
# ==============================================================================
def measure_current_stroke_width(binary_img, scale_mm_per_px):
    """Đo độ rộng nét trung bình hiện tại của phôi ảnh gốc (tính bằng mm)."""
    dist_transform = cv2.distanceTransform(binary_img, cv2.DIST_L2, 5)
    # Bán kính nét trung bình tại các vùng phôi chính
    max_radii = dist_transform[dist_transform > 0]
    if len(max_radii) == 0:
        return 46.0  # Mặc định nếu không đo được

    # Lấy giá trị bán kính trung bình của các đường nét
    avg_radius_px = np.median(max_radii)
    current_stroke_mm = (avg_radius_px * 2.0) * scale_mm_per_px
    return current_stroke_mm


def process_pattern_exact_thickness(
    image_bytes,
    target_w_mm=810.0,
    target_h_mm=2280.0,
    desired_stroke_mm=70.0,
):
    """Trích xuất hoa văn và bù trừ chính xác để đạt độ rộng nét tuyệt đối theo mm."""
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
        return [], (thresh.shape[1], thresh.shape[0]), 0.0

    main_contour = max(contours_external, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(main_contour)
    pattern_crop = thresh[y : y + h, x : x + w]

    scale_x = target_w_mm / w
    scale_y = target_h_mm / h
    avg_scale = (scale_x + scale_y) / 2.0

    # 3. ĐO ĐỘ RỘNG NẾT THỰC TẾ TRÊN ẢNH GỐC
    measured_stroke_mm = measure_current_stroke_width(pattern_crop, avg_scale)

    # 4. TÍNH TOÁN BÙ TRỪ KÍCH THƯỚC ĐỂ ĐẠT ĐÚNG DESIRED_STROKE_MM
    diff_mm = desired_stroke_mm - measured_stroke_mm

    if abs(diff_mm) > 1.0:
        # Số pixel cần bù trừ về mỗi bên
        pad_px = int(round((abs(diff_mm) / 2.0) / avg_scale))
        if pad_px > 0:
            kernel_size = pad_px * 2 + 1
            kernel = cv2.getStructuringElement(
                cv2.MORPH_ELLIPSE, (kernel_size, kernel_size)
            )

            if diff_mm > 0:
                # Phóng nét nếu kích thước đo được nhỏ hơn mong muốn
                pattern_crop = cv2.dilate(pattern_crop, kernel, iterations=1)
            else:
                # Thu nhỏ nét nếu kích thước đo được lớn hơn mong muốn
                pattern_crop = cv2.erode(pattern_crop, kernel, iterations=1)

    # 5. Trích xuất đường viền chuẩn phân cấp
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

    return pattern_polygons, (w, h), measured_stroke_mm


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
st.title("🌺 Hybrid AI Engine: ĐO VÀ ÉP NẾT HOA VĂN THỰC TẾ (MM)")

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
        st.subheader("📐 Bảng Tùy Chỉnh Kích Thước Phôi & Nét")

        st.session_state.width_mm = st.number_input(
            "Chiều rộng phôi (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều cao phôi (mm):", value=st.session_state.height_mm
        )

        desired_stroke = st.number_input(
            "Độ rộng nét mong muốn trong DXF (mm):",
            value=70.0,
            step=1.0,
            help="Hệ thống sẽ đo nét thực tế trên ảnh và tự động bù trừ đủ số mm này.",
        )

        st.markdown("---")

        if st.button("⚙️ NÚT CAD: Trích Xuất Vector & Tải File DXF"):
            with st.spinner("Đang đo nét gốc và bù trừ kích thước mm..."):
                pattern_polygons, (w_px, h_px), measured_mm = (
                    process_pattern_exact_thickness(
                        img_bytes,
                        st.session_state.width_mm,
                        st.session_state.height_mm,
                        desired_stroke_mm=desired_stroke,
                    )
                )

                dxf_bytes = create_dxf_file(pattern_polygons)

            st.info(
                f"📏 Nét hoa văn đo được trên ảnh gốc: **{measured_mm:.1f}mm**"
            )
            st.success(
                f"✅ Đã tự động bù trừ **{desired_stroke - measured_mm:+.1f}mm** để đưa nét về đúng **{desired_stroke}mm**!"
            )

            st.download_button(
                label=f"💾 TẢI FILE DXF CHUẨN {int(desired_stroke)}MM",
                data=dxf_bytes,
                file_name=f"Hoa_Van_Chuan_{int(desired_stroke)}mm.dxf",
                mime="application/dxf",
            )
