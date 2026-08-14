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
    """Đọc kích thước bản vẽ tự động bằng Gemini AI."""
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
# 2. THUẬT TOÁN ĐO NẾT THỰC TẾ & BÙ TRỪ KÍCH THƯỚC ĐÚNG ABSOLUTE MM
# ==============================================================================
def measure_pattern_thickness_px(binary_img):
    """Đo bán kính/độ rộng nét trung bình của hoa văn trên phôi binary."""
    dist = cv2.distanceTransform(binary_img, cv2.DIST_L2, 5)
    vals = dist[dist > 2]
    if len(vals) == 0:
        return 20.0
    # Bán kính nét trung bình (pixel)
    median_radius = np.median(vals)
    return median_radius * 2.0  # Đường kính nét (pixel)


def process_vector_exact_mm(
    image_bytes,
    target_w_mm=810.0,
    target_h_mm=2280.0,
    pattern_stroke_mm=70.0,
    frame_stroke_mm=70.0,
):
    """Tách biệt xử lý Khung và Hoa văn, bù trừ chính xác đúng số mm nhập vào."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # 2. Cắt bỏ lề chứa chữ & ghi chú bên ngoài
    contours_ext, _ = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    if not contours_ext:
        return []

    main_cnt = max(contours_ext, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(main_cnt)

    crop_thresh = thresh[y : y + h, x : x + w]

    # Tỉ lệ đổi pixel sang mm
    scale_x = target_w_mm / w
    scale_y = target_h_mm / h
    avg_scale = (scale_x + scale_y) / 2.0

    # 3. XỬ LÝ HOA VĂN BÊN TRONG (BÙ TRỪ ĐÚNG mm PATTERN)
    # Đo nét hoa văn hiện tại trên ảnh gốc (tính ra mm)
    current_stroke_px = measure_pattern_thickness_px(crop_thresh)
    current_stroke_mm = current_stroke_px * avg_scale

    # Tính lượng chênh lệch mm cần bù
    diff_pattern_mm = pattern_stroke_mm - current_stroke_mm

    pattern_img = crop_thresh.copy()
    if abs(diff_pattern_mm) > 1.0:
        # Số pixel cần dilate/erode
        pad_px = int(round((abs(diff_pattern_mm) / 2.0) / avg_scale))
        if pad_px > 0:
            k_size = pad_px * 2 + 1
            kernel = cv2.getStructuringElement(
                cv2.MORPH_ELLIPSE, (k_size, k_size)
            )
            if diff_pattern_mm > 0:
                pattern_img = cv2.dilate(pattern_img, kernel, iterations=1)
            else:
                pattern_img = cv2.erode(pattern_img, kernel, iterations=1)

    # Trích xuất đường viền hoa văn
    contours, hierarchy = cv2.findContours(
        pattern_img, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )

    all_polygons = []

    # 4. TẠO KHUNG BAO NGOÀI CHUẨN XÁC KÍCH THƯỚC FRAME_STROKE_MM
    # A. Đường viền ngoài cùng của tấm phôi (0,0) -> (W, H)
    outer_box = [
        (0.0, 0.0),
        (target_w_mm, 0.0),
        (target_w_mm, target_h_mm),
        (0.0, target_h_mm),
        (0.0, 0.0),
    ]
    all_polygons.append(outer_box)

    # B. Đường viền lòng trong của khung bao (Lùi vào đúng frame_stroke_mm mỗi bên)
    fx1 = frame_stroke_mm
    fy1 = frame_stroke_mm
    fx2 = target_w_mm - frame_stroke_mm
    fy2 = target_h_mm - frame_stroke_mm

    if fx2 > fx1 and fy2 > fy1:
        inner_box = [
            (fx1, fy1),
            (fx2, fy1),
            (fx2, fy2),
            (fx1, fy2),
            (fx1, fy1),
        ]
        all_polygons.append(inner_box)

    # 5. LỌC VÀ LẤY CHỈ CÁC HOẠT TIẾT HOA VĂN NẰM TRONG KHUNG
    if hierarchy is not None:
        for i, cnt in enumerate(contours):
            if cv2.contourArea(cnt) < 50:
                continue

            # Bỏ qua contour ngoài cùng của ảnh crop (vì đã tạo khung chuẩn ở trên)
            if hierarchy[0][i][3] == -1:
                continue

            epsilon = 0.001 * cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, epsilon, True)

            pts = []
            for p in approx:
                px, py = p[0][0], p[0][1]
                # Chuyển đổi tọa độ mm và đảo trục Y cho CAD
                mx = px * scale_x
                my = (h - py) * scale_y
                pts.append((mx, my))

            if len(pts) > 2:
                pts.append(pts[0])
                all_polygons.append(pts)

    return all_polygons, current_stroke_mm


# ==============================================================================
# 3. XUẤT FILE DXF CẮT CNC
# ==============================================================================
def create_dxf_file(polygons):
    """Xuất file DXF chứa đầy đủ đường cắt chuẩn."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    doc.layers.add(name="CAT_CNC_HOA_VAN", color=1)

    for poly in polygons:
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
st.title("🌺 Hybrid AI Engine: Ép Chuẩn Kích Thước DXF (Hoa Văn & Khung)")

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
        st.subheader("📐 Kích Thước Phôi Tổng Thể (mm)")

        st.session_state.width_mm = st.number_input(
            "Chiều rộng phôi tổng thể (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều cao phôi tổng thể (mm):", value=st.session_state.height_mm
        )

        st.markdown("---")
        st.subheader("🎛️ Tùy Chỉnh Độ Rộng Nét Thực Tế (DXF)")

        pattern_stroke = st.number_input(
            "1. Độ rộng nét HOA VĂN trong DXF (mm):",
            value=70.0,
            step=1.0,
            help="Đúng số mm khi dùng thước đo trong AutoCAD cho hoa văn",
        )

        frame_stroke = st.number_input(
            "2. Độ rộng KHUNG BAO NGOÀI trong DXF (mm):",
            value=70.0,
            step=1.0,
            help="Đúng số mm khi dùng thước đo trong AutoCAD cho khung viền",
        )

        st.markdown("---")

        if st.button("⚙️ NÚT CAD: Trích Xuất Vector & Tải File DXF"):
            with st.spinner("Đang tính toán bù trừ kích thước chuẩn xác..."):
                polygons, measured_stroke_mm = process_vector_exact_mm(
                    img_bytes,
                    st.session_state.width_mm,
                    st.session_state.height_mm,
                    pattern_stroke_mm=pattern_stroke,
                    frame_stroke_mm=frame_stroke,
                )

                dxf_bytes = create_dxf_file(polygons)

            st.info(
                f"📏 Độ rộng nét hoa văn gốc trên ảnh đo được: **{measured_stroke_mm:.1f}mm**"
            )
            st.success(
                f"✅ Đã ép chuẩn file DXF: **Khung = {frame_stroke}mm** | **Hoa văn = {pattern_stroke}mm**!"
            )

            st.download_button(
                label=f"💾 TẢI FILE DXF (Hoa văn {int(pattern_stroke)}mm - Khung {int(frame_stroke)}mm)",
                data=dxf_bytes,
                file_name=f"Hoa_Van_Chuan_{int(pattern_stroke)}mm_Khung_{int(frame_stroke)}mm.dxf",
                mime="application/dxf",
            )
