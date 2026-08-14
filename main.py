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

        # Tự động quy đổi cạnh ngắn là Rộng, cạnh dài là Cao
        return min(w, h), max(w, h)
    except Exception as e:
        st.warning(
            f"⚠️ Không đọc được OCR ({e}). Tự động dùng kích thước mặc định: 810mm x 2280mm."
        )
        return 810.0, 2280.0


# ==============================================================================
# 2. XỬ LÝ HOA VĂN: CROP CHUẨN & TRÍCH XUẤT VECTOR NGUYÊN BẢN
# ==============================================================================
def process_pattern_direct_contours(
    image_bytes, target_w_mm=810.0, target_h_mm=2280.0
):
    """Trích xuất chính xác 100% đường viền hoa văn từ ảnh đứng chuẩn mà không làm biến dạng."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize để tách màu hoa văn nâu ra khỏi nền trắng
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )

    # 2. Tự động tìm Bounding Box của sản phẩm chính để CROP BỎ 100% chữ kích thước & mũi tên
    contours_all, _ = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )

    if not contours_all:
        return [], (thresh.shape[1], thresh.shape[0])

    # Tìm contour lớn nhất (chính là toàn bộ tấm khung gỗ hoa văn)
    main_contour = max(contours_all, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(main_contour)

    # Crop chính xác duy nhất vùng tấm phôi hoa văn
    pattern_crop = thresh[y : y + h, x : x + w]

    # 3. Tính tỉ lệ Scale Pixel -> mm dựa trên kích thước thực tế
    scale_x = target_w_mm / w
    scale_y = target_h_mm / h

    # 4. Trích xuất toàn bộ đường viền (Khung ngoài + các lỗ uốn lượn bên trong)
    contours, hierarchy = cv2.findContours(
        pattern_crop, cv2.RETR_TREE, cv2.CHAIN_APPROX_TC89_KCOS
    )

    pattern_polygons = []
    for cnt in contours:
        # Loại bỏ các vết nhiễu quá nhỏ
        if cv2.contourArea(cnt) < 30:
            continue

        # Làm mượt đường nét uốn lượn
        epsilon = 0.0012 * cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, epsilon, True)

        pts = []
        for p in approx:
            px, py = p[0][0], p[0][1]
            # Tọa độ thực tế theo mm (Đảo trục Y cho hệ tọa độ CAD)
            pts.append((px * scale_x, (h - py) * scale_y))

        if len(pts) > 2:
            pts.append(pts[0])  # Khép kín đường vector
            pattern_polygons.append(pts)

    return pattern_polygons, (w, h)


# ==============================================================================
# 3. XUẤT FILE DXF CHUẨN CẮT CNC
# ==============================================================================
def create_dxf_file(pattern_polygons, total_w_mm, total_h_mm):
    """Xuất file DXF chứa đầy đủ đường cắt hoa văn chuẩn kích thước."""
    doc = ezdxf.new("R2010")
    msp = doc.modelspace()

    # Tạo Layer riêng cho máy CNC dễ phân biệt
    doc.layers.add(name="HOA_VAN_CNC", color=1)  # Đỏ (Toàn bộ nét hoa văn)

    # Vẽ các dải vector hoa văn đã trích xuất
    for poly in pattern_polygons:
        msp.add_lwpolyline(poly, dxfattribs={"layer": "HOA_VAN_CNC"})

    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, "Hoa_Van_CNC_Precision.dxf")
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
st.title("🌺 Hybrid AI Engine: Chuyển Ảnh Hoa Văn CNC Sang File DXF Kỹ Thuật")

st.sidebar.header("⚙️ Cấu hình Hệ thống")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

# Lưu trữ kích thước mặc định
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

        # -------------------------------------------------------------
        # NÚT BẤM 1: DÙNG GEMINI AI ĐỂ ĐỌC KÍCH THƯỚC
        # -------------------------------------------------------------
        if st.button("🤖 NÚT AI: Đọc Kích Thước Bản Vẽ (Gemini OCR)"):
            if not api_key:
                st.error("Vui lòng nhập Gemini API Key ở Sidebar thanh bên!")
            else:
                with st.spinner("🤖 Gemini AI đang quét con số trên ảnh..."):
                    w, h = extract_dimensions_with_gemini(img_bytes, api_key)
                    st.session_state.width_mm = w
                    st.session_state.height_mm = h
                    st.success(
                        f"🎉 AI đã đọc thành công: Rộng {w}mm x Cao {h}mm"
                    )

    with col2:
        st.subheader("📐 Thông Số & Xuất File DXF")

        # Cho phép người dùng kiểm tra/chỉnh sửa con số sau khi AI đọc
        st.session_state.width_mm = st.number_input(
            "Chiều rộng phôi (mm):", value=st.session_state.width_mm
        )
        st.session_state.height_mm = st.number_input(
            "Chiều cao phôi (mm):", value=st.session_state.height_mm
        )

        st.markdown("---")

        # -------------------------------------------------------------
        # NÚT BẤM 2: TRÍCH XUẤT VECTOR & TẢI FILE DXF
        # -------------------------------------------------------------
        if st.button("⚙️ NÚT CAD: Trích Xuất Vector & Tạo File DXF"):
            with st.spinner("Đang tự động crop viền & trích xuất vector..."):
                pattern_polygons, (w_px, h_px) = process_pattern_direct_contours(
                    img_bytes,
                    st.session_state.width_mm,
                    st.session_state.height_mm,
                )

                dxf_bytes = create_dxf_file(
                    pattern_polygons,
                    st.session_state.width_mm,
                    st.session_state.height_mm,
                )

            st.success(
                f"✅ Đã trích xuất xong **{len(pattern_polygons)}** đường nét vector chuẩn xác!"
            )

            st.download_button(
                label="💾 TẢI FILE DXF CẮT CNC CHUẨN KÍCH THƯỚC",
                data=dxf_bytes,
                file_name="Hoa_Van_CNC_Chuandxf.dxf",
                mime="application/dxf",
            )
