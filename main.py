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
        description="Chiều rộng ghi chú trên ảnh (mm) - Ví dụ 2280"
    )
    height_mm: float = Field(
        description="Chiều cao ghi chú trên ảnh (mm) - Ví dụ 810"
    )


def extract_dimensions_with_gemini(image_bytes, api_key):
    """Đọc kích thước tự động bằng Gemini 1.5 Flash."""
    try:
        client = genai.Client(api_key=api_key)
        img = Image.open(io.BytesIO(image_bytes))

        prompt = """
        Đọc 2 kích thước ghi chú trên bản vẽ:
        - Chiều rộng nằm ngang (Ví dụ: 2280mm)
        - Chiều cao thẳng đứng (Ví dụ: 810mm)
        Trả về đúng định dạng JSON Schema.
        """

        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=[img, prompt],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=DrawingDimensions,
            ),
        )

        data = json.loads(response.text.strip())
        w = data.get("width_mm", 2280.0)
        h = data.get("height_mm", 810.0)
        return w, h
    except Exception as e:
        st.warning(f"⚠️ OCR AI: {e}. Dùng kích thước mặc định: 2280mm x 810mm.")
        return 2280.0, 810.0


# ==============================================================================
# 2. XỬ LÝ TRÍCH XUẤT VECTOR CHUẨN DÁNG + KHUNG THEO YÊU CẦU
# ==============================================================================
def process_exact_pattern_with_custom_frame(
    image_bytes, target_w_mm=2280.0, target_h_mm=810.0, frame_stroke_mm=70.0
):
    """Trích xuất trọn vẹn hoa văn gốc 100% kết hợp DỰNG KHUNG BAO CHUẨN KÍCH THƯỚC (mm)."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize tách phôi màu nâu
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # 2. Cắt bỏ lề ghi chú chữ xung quanh
    contours_ext, _ = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    if not contours_ext:
        return []

    main_cnt = max(contours_ext, key=cv2.contourArea)
    x, y, w_px, h_px = cv2.boundingRect(main_cnt)

    pattern_crop = thresh[y : y + h_px, x : x + w_px]

    scale_x = target_w_mm / float(w_px)
    scale_y = target_h_mm / float(h_px)

    all_polygons = []

    # 3. DỰNG KHUNG BAO NGOÀI BẰNG HÌNH HỌC TỰ ĐỘNG (ĐẢM BẢO VIỀN KHUNG ĐÚNG MM)
    # A. Đường viền biên ngoài cùng phôi (0,0) -> (target_w_mm, target_h_mm)
    outer_box = [
        (0.0, 0.0),
        (target_w_mm, 0.0),
        (target_w_mm, target_h_mm),
        (0.0, target_h_mm),
        (0.0, 0.0),
    ]
    all_polygons.append(outer_box)

    # B. Đường viền lọt lòng trong của khung (Lùi vào đúng frame_stroke_mm mỗi bên)
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

    # 4. TRÍCH XUẤT CÁC LỖ THỦNG HOA VĂN BÊN TRONG (GIỮ 100% ĐƯỜNG CONG NGUYÊN BẢN)
    contours, hierarchy = cv2.findContours(
        pattern_crop, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )

    if hierarchy is not None:
        for i, cnt in enumerate(contours):
            # Lọc bớt vết chấm nhiễu rác
            if cv2.contourArea(cnt) < 30:
                continue

            # Bỏ qua viền ngoài cùng của ảnh crop (vì đã dựng khung chuẩn ở bước 3)
            if hierarchy[0][i][3] == -1:
                continue

            epsilon = 0.0008 * cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, epsilon, True)

            pts = []
            for p in approx:
                px, py = p[0][0], p[0][1]
                # Đổi sang mm thực tế & Đảo trục Y cho AutoCAD
                mx = px * scale_x
                my = (h_px - py) * scale_y
                pts.append((mx, my))

            if len(pts) > 2:
                pts.append(pts[0])
                all_polygons.append(pts)

    return all_polygons


# ==============================================================================
# 3. XUẤT FILE DXF CẮT CNC
# ==============================================================================
def create_dxf_file(polygons):
    """Xuất file DXF chuẩn layer CNC."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    doc.layers.add(name="CAT_CNC_HOA_VAN", color=1)

    for poly in polygons:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "CAT_CNC_HOA_VAN"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "Hoa_Van_CNC_Chuan.dxf")
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
st.title("🌺 Hybrid AI Engine: Trích Xuất Vector Chuẩn CNC")

st.sidebar.header("⚙️ Cấu hình Hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

if "width_mm" not in st.session_state:
    st.session_state.width_mm = 2280.0
if "height_mm" not in st.session_state:
    st.session_state.height_mm = 810.0

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
        st.subheader("📐 Thông Số Kích Thước Phôi & Khung Bao")

        st.session_state.width_mm = st.number_input(
            "Chiều RỘNG phôi (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều CAO phôi (mm):", value=st.session_state.height_mm
        )

        st.markdown("---")

        frame_stroke = st.number_input(
            "Độ rộng KHUNG BAO NGOÀI (mm):",
            value=70.0,
            step=5.0,
            help="Hệ thống sẽ dựng khung chữ nhật toán học chính xác số mm này",
        )

        st.markdown("---")

        if st.button("⚙️ NÚT CAD: Trích Xuất Vector & Tải File DXF"):
            with st.spinner("Đang dựng khung và trích xuất hoa văn..."):
                polygons = process_exact_pattern_with_custom_frame(
                    img_bytes,
                    st.session_state.width_mm,
                    st.session_state.height_mm,
                    frame_stroke_mm=frame_stroke,
                )

                dxf_bytes = create_dxf_file(polygons)

            st.success(
                f"✅ Đã tạo thành công **{len(polygons)}** đường nét cắt DXF chuẩn xác!"
            )

            st.download_button(
                label=f"💾 TẢI FILE DXF (Khung {int(frame_stroke)}mm - Hoa văn nguyên bản)",
                data=dxf_bytes,
                file_name=f"Hoa_Van_CNC_Khung_{int(frame_stroke)}mm.dxf",
                mime="application/dxf",
            )
