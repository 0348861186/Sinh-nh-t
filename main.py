import streamlit as st
import cv2
import numpy as np
import ezdxf
from PIL import Image
import io
import os
import re
import math
from google import genai
import matplotlib.pyplot as plt


# ============================================================
# CẤU HÌNH TRANG
# ============================================================

st.set_page_config(
    page_title="Bảng Điều Khiển AI CAD/CAM",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Bảng Điều Khiển AI CAD/CAM")
st.caption("Hình ảnh / DXF → Hình học → DXF / Mã G-Code")


# ============================================================
# TRẠNG THÁI PHIÊN (SESSION STATE)
# ============================================================

DEFAULT_STATE = {
    "source_type": None,
    "source_image": None,
    "geometry": [],
    "processed_image": None,
    "dxf_content": None,
    "gcode_content": None,
    "ai_thoughts": "",
    "scale_mm_per_pixel": 1.0,
    "source_width": 0.0,
    "source_height": 0.0,
    "final_width": 0.0,
    "final_height": 0.0,
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# CẤU TRÚC DỮ LIỆU
# ============================================================

def make_geometry(points, closed=True, source="image", is_hole=False):
    """
    Chuẩn hóa geometry thành dictionary.

    points:
        list[(x, y)]

    closed:
        contour có đóng hay không

    source:
        image / dxf

    is_hole:
        True nếu là contour bên trong
    """

    return {
        "points": [(float(x), float(y)) for x, y in points],
        "closed": bool(closed),
        "source": source,
        "is_hole": bool(is_hole),
    }


# ============================================================
# HÀM HỖ TRỢ HÌNH HỌC (GEOMETRY HELPERS)
# ============================================================

def geometry_bounds(geometry):
    """Lấy bounding box của toàn bộ geometry."""

    all_points = []

    for item in geometry:
        all_points.extend(item["points"])

    if not all_points:
        return 0, 0, 0, 0

    xs = [p[0] for p in all_points]
    ys = [p[1] for p in all_points]

    min_x = min(xs)
    max_x = max(xs)
    min_y = min(ys)
    max_y = max(ys)

    return min_x, min_y, max_x, max_y


def geometry_size(geometry):
    """Trả về Chiều rộng / Chiều cao."""

    min_x, min_y, max_x, max_y = geometry_bounds(geometry)

    return (
        max_x - min_x,
        max_y - min_y
    )


def polygon_area(points):
    """
    Công thức Shoelace.
    Dùng để xác định kích thước contour.
    """

    if len(points) < 3:
        return 0.0

    area = 0.0

    for i in range(len(points)):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % len(points)]

        area += x1 * y2 - x2 * y1

    return abs(area) / 2.0


def point_distance(p1, p2):
    return math.sqrt(
        (p1[0] - p2[0]) ** 2 +
        (p1[1] - p2[1]) ** 2
    )


# ============================================================
# XỬ LÝ HÌNH ẢNH
# ============================================================

def process_image_contours(
    img_bytes,
    min_area_px=20,
    threshold_mode="OTSU"
):
    """
    Hình ảnh → Các contour OpenCV.
    """

    image = Image.open(
        io.BytesIO(img_bytes)
    ).convert("RGB")

    img_np = np.array(image)

    gray = cv2.cvtColor(
        img_np,
        cv2.COLOR_RGB2GRAY
    )

    # -----------------------------------------
    # Làm mờ (Blur)
    # -----------------------------------------

    blurred = cv2.GaussianBlur(
        gray,
        (5, 5),
        0
    )

    # -----------------------------------------
    # Phân ngưỡng (Threshold)
    # -----------------------------------------

    if threshold_mode == "OTSU":

        _, thresh = cv2.threshold(
            blurred,
            0,
            255,
            cv2.THRESH_BINARY_INV +
            cv2.THRESH_OTSU
        )

    elif threshold_mode == "ADAPTIVE":

        thresh = cv2.adaptiveThreshold(
            blurred,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            11,
            2
        )

    else:

        _, thresh = cv2.threshold(
            blurred,
            127,
            255,
            cv2.THRESH_BINARY_INV
        )

    # -----------------------------------------
    # Làm sạch hình thái học (Morphological cleanup)
    # -----------------------------------------

    kernel = np.ones((3, 3), np.uint8)

    thresh = cv2.morphologyEx(
        thresh,
        cv2.MORPH_OPEN,
        kernel
    )

    thresh = cv2.morphologyEx(
        thresh,
        cv2.MORPH_CLOSE,
        kernel
    )

    # -----------------------------------------
    # Tìm đường viền (Contours)
    # -----------------------------------------

    contours, hierarchy = cv2.findContours(
        thresh,
        cv2.RETR_TREE,
        cv2.CHAIN_APPROX_SIMPLE
    )

    geometry = []

    if hierarchy is not None:
        hierarchy = hierarchy[0]

    for index, contour in enumerate(contours):

        area = cv2.contourArea(contour)

        if area < min_area_px:
            continue

        if len(contour) < 3:
            continue

        pts = contour.reshape(-1, 2)

        points = [
            (float(p[0]), float(p[1]))
            for p in pts
        ]

        # -----------------------------------------
        # Xác định lỗ thủng (hole) bằng cây phân cấp (hierarchy)
        # -----------------------------------------

        is_hole = False

        if hierarchy is not None:

            parent = hierarchy[index][3]

            if parent != -1:
                is_hole = True

        geometry.append(
            make_geometry(
                points,
                closed=True,
                source="image",
                is_hole=is_hole
            )
        )

    return img_np, geometry


# ============================================================
# ĐƠN GIẢN HÓA HÌNH HỌC (SIMPLIFICATION)
# ============================================================

def simplify_geometry(
    geometry,
    tolerance_px=1.0
):
    """
    Giảm số lượng đỉnh (vertex).

    tolerance_px:
        Độ dung sai tính bằng pixel đối với Hình ảnh.
    """

    result = []

    for item in geometry:

        points = item["points"]

        if len(points) < 3:
            continue

        np_points = np.array(
            points,
            dtype=np.float32
        )

        contour = np_points.reshape(
            (-1, 1, 2)
        )

        approx = cv2.approxPolyDP(
            contour,
            tolerance_px,
            True
        )

        simplified = [
            (
                float(p[0][0]),
                float(p[0][1])
            )
            for p in approx
        ]

        if len(simplified) >= 3:

            result.append(
                make_geometry(
                    simplified,
                    closed=item["closed"],
                    source=item["source"],
                    is_hole=item["is_hole"]
                )
            )

    return result


# ============================================================
# TRÌNH ĐỌC FILE DXF
# ============================================================

def read_dxf_file(dxf_bytes):
    """
    Đọc các thực thể cơ bản từ DXF:

    LINE
    LWPOLYLINE
    POLYLINE
    CIRCLE
    ARC
    """

    geometry = []

    stream = io.StringIO(
        dxf_bytes.decode(
            "utf-8",
            errors="ignore"
        )
    )

    try:
        doc = ezdxf.read(stream)

    except Exception as e:
        raise RuntimeError(
            f"Không thể đọc DXF: {e}"
        )

    msp = doc.modelspace()

    for entity in msp:

        dxftype = entity.dxftype()

        # =====================================
        # LINE (ĐƯỜNG THẲNG)
        # =====================================

        if dxftype == "LINE":

            start = entity.dxf.start
            end = entity.dxf.end

            points = [
                (start.x, start.y),
                (end.x, end.y)
            ]

            geometry.append(
                make_geometry(
                    points,
                    closed=False,
                    source="dxf"
                )
            )

        # =====================================
        # LWPOLYLINE (ĐA TUYẾN TRỌNG LƯỢNG NHẸ)
        # =====================================

        elif dxftype == "LWPOLYLINE":

            points = []

            for point in entity.get_points():

                x = point[0]
                y = point[1]

                points.append(
                    (x, y)
                )

            closed = bool(
                entity.closed
            )

            if len(points) >= 2:

                geometry.append(
                    make_geometry(
                        points,
                        closed=closed,
                        source="dxf"
                    )
                )

        # =====================================
        # POLYLINE (ĐA TUYẾN)
        # =====================================

        elif dxftype == "POLYLINE":

            points = []

            for vertex in entity.vertices:

                location = vertex.dxf.location

                points.append(
                    (
                        location.x,
                        location.y
                    )
                )

            closed = bool(
                entity.is_closed
            )

            if len(points) >= 2:

                geometry.append(
                    make_geometry(
                        points,
                        closed=closed,
                        source="dxf"
                    )
                )

        # =====================================
        # CIRCLE (ĐƯỜNG TRÒN)
        # =====================================

        elif dxftype == "CIRCLE":

            center = entity.dxf.center
            radius = entity.dxf.radius

            points = []

            segments = 96

            for i in range(segments):

                angle = (
                    2 *
                    math.pi *
                    i /
                    segments
                )

                x = (
                    center.x +
                    radius *
                    math.cos(angle)
                )

                y = (
                    center.y +
                    radius *
                    math.sin(angle)
                )

                points.append(
                    (x, y)
                )

            geometry.append(
                make_geometry(
                    points,
                    closed=True,
                    source="dxf"
                )
            )

        # =====================================
        # ARC (CUNG TRÒN)
        # =====================================

        elif dxftype == "ARC":

            center = entity.dxf.center
            radius = entity.dxf.radius

            start_angle = math.radians(
                entity.dxf.start_angle
            )

            end_angle = math.radians(
                entity.dxf.end_angle
            )

            if end_angle <= start_angle:
                end_angle += 2 * math.pi

            segments = max(
                12,
                int(
                    abs(end_angle - start_angle)
                    * 180 /
                    math.pi
                    / 3
                )
            )

            points = []

            for i in range(segments + 1):

                angle = (
                    start_angle +
                    (
                        end_angle -
                        start_angle
                    ) *
                    i /
                    segments
                )

                x = (
                    center.x +
                    radius *
                    math.cos(angle)
                )

                y = (
                    center.y +
                    radius *
                    math.sin(angle)
                )

                points.append(
                    (x, y)
                )

            geometry.append(
                make_geometry(
                    points,
                    closed=False,
                    source="dxf"
                )
            )

    return geometry


# ============================================================
# CHUYỂN ĐỔI TỌA ĐỘ HÌNH ẢNH
# ============================================================

def convert_image_geometry_to_mm(
    geometry,
    scale_mm_per_pixel
):
    """
    Hình ảnh:
        X → X
        Y ↓

    Máy CNC / CAD:
        X → X
        Y ↑

    Vì vậy Y được đảo chiều.
    """

    result = []

    for item in geometry:

        converted = []

        for x, y in item["points"]:

            x_mm = (
                x *
                scale_mm_per_pixel
            )

            y_mm = (
                -y *
                scale_mm_per_pixel
            )

            converted.append(
                (x_mm, y_mm)
            )

        result.append(
            make_geometry(
                converted,
                closed=item["closed"],
                source="image",
                is_hole=item["is_hole"]
            )
        )

    return result


# ============================================================
# THAY ĐỔI TỶ LỆ (SCALE) DXF
# ============================================================

def scale_geometry(
    geometry,
    scale_x=1.0,
    scale_y=1.0
):

    result = []

    for item in geometry:

        points = [
            (
                x * scale_x,
                y * scale_y
            )
            for x, y in item["points"]
        ]

        result.append(
            make_geometry(
                points,
                closed=item["closed"],
                source=item["source"],
                is_hole=item["is_hole"]
            )
        )

    return result


# ============================================================
# THIẾT LẬP GỐC TỌA ĐỘ G54 (WORK ZERO)
# ============================================================

def apply_work_zero(
    geometry,
    work_zero
):
    """
    Căn chỉnh geometry theo gốc tọa độ máy đã chọn.
    """

    min_x, min_y, max_x, max_y = (
        geometry_bounds(geometry)
    )

    width = max_x - min_x
    height = max_y - min_y

    if work_zero == "Bottom Left":

        offset_x = -min_x
        offset_y = -min_y

    elif work_zero == "Bottom Center":

        offset_x = -(
            min_x + width / 2
        )

        offset_y = -min_y

    elif work_zero == "Center":

        offset_x = -(
            min_x + width / 2
        )

        offset_y = -(
            min_y + height / 2
        )

    elif work_zero == "Top Left":

        offset_x = -min_x
        offset_y = -max_y

    elif work_zero == "Top Right":

        offset_x = -max_x
        offset_y = -max_y

    else:

        offset_x = 0
        offset_y = 0

    result = []

    for item in geometry:

        points = [
            (
                x + offset_x,
                y + offset_y
            )
            for x, y in item["points"]
        ]

        result.append(
            make_geometry(
                points,
                closed=item["closed"],
                source=item["source"],
                is_hole=item["is_hole"]
            )
        )

    return result


# ============================================================
# SẮP XẾP THỨ TỰ CẮT
# ============================================================

def sort_geometry_for_cutting(
    geometry
):
    """
    Đường viền trong (hole) cắt trước.
    Đường viền ngoài (outer) cắt sau.

    Nếu không xác định được hole,
    dùng diện tích nhỏ → lớn.
    """

    def sort_key(item):

        area = polygon_area(
            item["points"]
        )

        # Lỗ thủng (Hole) ưu tiên trước
        hole_priority = (
            0
            if item["is_hole"]
            else 1
        )

        return (
            hole_priority,
            area
        )

    return sorted(
        geometry,
        key=sort_key
    )


# ============================================================
# TRÌNH TẠO FILE DXF
# ============================================================

def generate_dxf(
    geometry
):
    """
    Geometry → DXF R2010.
    """

    doc = ezdxf.new(
        "R2010"
    )

    msp = doc.modelspace()

    # -----------------------------------------
    # Các lớp (Layers)
    # -----------------------------------------

    if "CUT" not in doc.layers:

        doc.layers.new(
            name="CUT"
        )

    for item in geometry:

        points = item["points"]

        if len(points) < 2:
            continue

        if item["closed"]:

            msp.add_lwpolyline(
                points,
                close=True,
                dxfattribs={
                    "layer": "CUT"
                }
            )

        else:

            msp.add_lwpolyline(
                points,
                close=False,
                dxfattribs={
                    "layer": "CUT"
                }
            )

    out_stream = io.StringIO()

    doc.write(
        out_stream
    )

    return out_stream.getvalue()


# ============================================================
# TRÌNH TẠO MÃ G-CODE
# ============================================================

def generate_gcode(
    geometry,
    feed_rate=800,
    plunge_rate=300,
    safe_z=5.0,
    cut_z=-1.0,
    spindle_rpm=12000,
    spindle_direction="M3"
):
    """
    Geometry → G-Code.

    Đường viền bên trong được cắt trước đường viền bên ngoài.
    """

    geometry = sort_geometry_for_cutting(
        geometry
    )

    lines = []

    lines.append(
        "; ========================================="
    )
    lines.append(
        "; Ma G-Code duoc tao boi AI CAD/CAM"
    )
    lines.append(
        "; ========================================="
    )

    lines.append(
        "G21 ; Don vi: mm"
    )

    lines.append(
        "G90 ; Toa do tuyet doi"
    )

    lines.append(
        "G17 ; Mat phang XY"
    )

    lines.append(
        f"G0 Z{safe_z:.3f}"
    )

    lines.append(
        f"{spindle_direction} S{spindle_rpm}"
    )

    lines.append(
        ""
    )

    for index, item in enumerate(
        geometry,
        start=1
    ):

        points = item["points"]

        if len(points) < 2:
            continue

        start_x, start_y = points[0]

        lines.append(
            f"; --- Duong vien {index} ---"
        )

        if item["is_hole"]:

            lines.append(
                "; Duong vien trong / Lo thung"
            )

        else:

            lines.append(
                "; Duong vien ngoai"
            )

        # -------------------------------------
        # Di chuyển đến điểm bắt đầu
        # -------------------------------------

        lines.append(
            f"G0 X{start_x:.3f} "
            f"Y{start_y:.3f}"
        )

        # -------------------------------------
        # Đâm dao (Plunge)
        # -------------------------------------

        lines.append(
            f"G1 Z{cut_z:.3f} "
            f"F{plunge_rate}"
        )

        # -------------------------------------
        # Gi công cắt (Cutting)
        # -------------------------------------

        for x, y in points[1:]:

            lines.append(
                f"G1 X{x:.3f} "
                f"Y{y:.3f} "
                f"F{feed_rate}"
            )

        # -------------------------------------
        # Đóng đường viền
        # -------------------------------------

        if item["closed"]:

            lines.append(
                f"G1 X{start_x:.3f} "
                f"Y{start_y:.3f} "
                f"F{feed_rate}"
            )

        # -------------------------------------
        # Rút dao (Retract)
        # -------------------------------------

        lines.append(
            f"G0 Z{safe_z:.3f}"
        )

        lines.append("")

    # -----------------------------------------
    # Kết thúc
    # -----------------------------------------

    lines.append(
        "M5 ; Tat truc chinh"
    )

    lines.append(
        "G0 Z%.3f" % safe_z
    )

    lines.append(
        "G0 X0 Y0 ; Ve goc toa do"
    )

    lines.append(
        "M30 ; Ket thuc chuong trinh"
    )

    return "\n".join(lines)


# ============================================================
# TRÌNH XEM TRƯỚC (PREVIEW)
# ============================================================

def create_geometry_preview(
    geometry,
    title="Xem Trước Hình Học"
):
    """
    Vẽ geometry bằng Matplotlib.
    """

    fig, ax = plt.subplots(
        figsize=(10, 7)
    )

    for index, item in enumerate(
        geometry
    ):

        points = item["points"]

        if len(points) < 2:
            continue

        xs = [
            p[0]
            for p in points
        ]

        ys = [
            p[1]
            for p in points
        ]

        if item["closed"]:

            xs.append(xs[0])
            ys.append(ys[0])

        ax.plot(
            xs,
            ys,
            linewidth=1
        )

    ax.set_title(title)

    ax.set_aspect(
        "equal",
        adjustable="box"
    )

    ax.grid(True)

    return fig


# ============================================================
# HÌNH ẢNH CÓ CHÚ THÍCH KÍCH THƯỚC
# ============================================================

def draw_dimensions_on_img(
    img_np,
    geometry,
    target_width_mm=None
):

    annotated = img_np.copy()

    if not geometry:

        return annotated

    all_pts = []

    for item in geometry:

        all_pts.extend(
            item["points"]
        )

    if not all_pts:

        return annotated

    pts_np = np.array(
        all_pts,
        dtype=np.int32
    )

    x, y, w, h = cv2.boundingRect(
        pts_np
    )

    # -----------------------------------------
    # Bounding box (Khung bao)
    # -----------------------------------------

    cv2.rectangle(
        annotated,
        (x, y),
        (x + w, y + h),
        (0, 255, 0),
        2
    )

    # -----------------------------------------
    # Các đường viền (Contours)
    # -----------------------------------------

    for item in geometry:

        pts = np.array(
            item["points"],
            dtype=np.int32
        )

        if len(pts) < 2:
            continue

        cv2.polylines(
            annotated,
            [pts],
            item["closed"],
            (255, 0, 0),
            1
        )

    # -----------------------------------------
    # Nhãn hiển thị kích thước
    # -----------------------------------------

    if target_width_mm:

        scale = (
            target_width_mm /
            w
            if w > 0
            else 1
        )

        height_mm = (
            h * scale
        )

        label = (
            f"Rong: {target_width_mm:.2f} mm | "
            f"Cao: {height_mm:.2f} mm"
        )

    else:

        label = (
            f"Rong: {w}px | Cao: {h}px"
        )

    cv2.putText(
        annotated,
        label,
        (
            x,
            max(y - 10, 25)
        ),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.6,
        (0, 255, 0),
        2
    )

    return annotated


# ============================================================
# GEMINI AI
# ============================================================

def ask_gemini(
    api_key,
    width,
    height,
    user_prompt
):
    """
    Gemini chỉ phân tích / tư vấn.
    Không để AI tự ý thay đổi geometry.
    """

    client = genai.Client(
        api_key=api_key
    )

    prompt = f"""
Bạn là kỹ sư CAD/CAM.

Thông tin geometry:
Width = {width:.3f} mm
Height = {height:.3f} mm

Yêu cầu người dùng:
{user_prompt if user_prompt else "Không có yêu cầu đặc biệt."}

Hãy phân tích ngắn gọn:

1. Geometry hiện tại có kích thước bao nhiêu.
2. Nếu người dùng yêu cầu kích thước mới thì cần scale như thế nào.
3. Gợi ý Feed Rate.
4. Gợi ý Plunge Rate.
5. Gợi ý Safe Z.
6. Gợi ý Cut Z.
7. Các vấn đề CAM cần kiểm tra trước khi đưa vào máy.

Không được tự ý giả định kích thước vật liệu nếu người dùng chưa cung cấp.

Trả lời chuyên nghiệp, ngắn gọn.
"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt
    )

    return response.text


# ============================================================
# THANH BÊN (SIDEBAR)
# ============================================================

st.sidebar.header(
    "🔑 Cấu hình AI"
)

api_key = st.sidebar.text_input(
    "Gemini API Key",
    type="password"
)

st.sidebar.divider()

st.sidebar.header(
    "⚙️ Xử lý hình ảnh"
)

min_area_px = st.sidebar.number_input(
    "Diện tích contour tối thiểu (px²)",
    min_value=1,
    value=20,
    step=1
)

threshold_mode = st.sidebar.selectbox(
    "Chế độ phân ngưỡng (Threshold)",
    [
        "OTSU",
        "ADAPTIVE",
        "FIXED"
    ]
)

simplify_tolerance = st.sidebar.number_input(
    "Dung sai đường viền (pixel)",
    min_value=0.0,
    value=1.0,
    step=0.1
)


# ============================================================
# ĐẦU VÀO (INPUT)
# ============================================================

st.subheader(
    "1️⃣ Dữ liệu đầu vào"
)

col1, col2 = st.columns(2)

with col1:

    uploaded_image = st.file_uploader(
        "🖼️ Tải lên hình ảnh",
        type=[
            "png",
            "jpg",
            "jpeg",
            "bmp",
            "tif",
            "tiff"
        ]
    )

with col2:

    uploaded_dxf = st.file_uploader(
        "📐 Tải lên tệp DXF",
        type=["dxf"]
    )


# ============================================================
# TẢI HÌNH ẢNH (LOAD IMAGE)
# ============================================================

if uploaded_image:

    try:

        img_bytes = uploaded_image.getvalue()

        (
            raw_img_np,
            geometry
        ) = process_image_contours(
            img_bytes,
            min_area_px=min_area_px,
            threshold_mode=threshold_mode
        )

        geometry = simplify_geometry(
            geometry,
            tolerance_px=simplify_tolerance
        )

        st.session_state.source_type = "image"

        st.session_state.source_image = raw_img_np

        st.session_state.geometry = geometry

        width_px, height_px = (
            geometry_size(
                geometry
            )
        )

        st.session_state.source_width = width_px
        st.session_state.source_height = height_px

        st.success(
            f"Đã nhận Hình ảnh: "
            f"{len(geometry)} đường viền (contour)"
        )

    except Exception as e:

        st.error(
            f"Lỗi xử lý Hình ảnh: {e}"
        )


# ============================================================
# TẢI FILE DXF (LOAD DXF)
# ============================================================

elif uploaded_dxf:

    try:

        dxf_bytes = (
            uploaded_dxf.getvalue()
        )

        geometry = read_dxf_file(
            dxf_bytes
        )

        st.session_state.source_type = "dxf"

        st.session_state.geometry = geometry

        width, height = (
            geometry_size(
                geometry
            )
        )

        st.session_state.source_width = width
        st.session_state.source_height = height

        st.success(
            f"Đã đọc file DXF: "
            f"{len(geometry)} thực thể (entities)"
        )

    except Exception as e:

        st.error(
            f"Lỗi đọc file DXF: {e}"
        )


# ============================================================
# DỪNG CHƯƠNG TRÌNH NẾU KHÔNG CÓ ĐẦU VÀO
# ============================================================

geometry = st.session_state.geometry

if not geometry:

    st.info(
        "Vui lòng tải lên Hình ảnh hoặc tệp DXF."
    )

    st.stop()


# ============================================================
# THÔNG TIN NGUỒN (SOURCE INFORMATION)
# ============================================================

st.divider()

st.subheader(
    "2️⃣ Thông tin hình học gốc"
)

width, height = geometry_size(
    geometry
)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "Số thực thể",
        len(geometry)
    )

with c2:
    st.metric(
        "Rộng (Nguồn)",
        f"{width:.3f}"
    )

with c3:
    st.metric(
        "Cao (Nguồn)",
        f"{height:.3f}"
    )

with c4:

    st.metric(
        "Nguồn dữ liệu",
        st.session_state.source_type
    )


# ============================================================
# KÍCH THƯỚC / TỶ LỆ (DIMENSION / SCALE)
# ============================================================

st.subheader(
    "3️⃣ Kích thước & Tỷ lệ quy đổi"
)

if st.session_state.source_type == "image":

    st.info(
        "Hình ảnh được quy đổi từ pixel sang mm. "
        "Bạn cần cung cấp kích thước thực tế."
    )

    target_width_mm = st.number_input(
        "Chiều rộng mục tiêu (mm)",
        min_value=0.001,
        value=100.0,
        step=1.0
    )

    pixel_width = width

    if pixel_width > 0:

        scale_mm_per_pixel = (
            target_width_mm /
            pixel_width
        )

    else:

        scale_mm_per_pixel = 1.0

    st.session_state.scale_mm_per_pixel = (
        scale_mm_per_pixel
    )

    st.write(
        f"**Tỷ lệ Scale:** "
        f"{scale_mm_per_pixel:.6f} mm/pixel"
    )

    geometry_mm = (
        convert_image_geometry_to_mm(
            geometry,
            scale_mm_per_pixel
        )
    )

else:

    st.info(
        "DXF được xem là đã sử dụng đơn vị CAD hiện tại."
    )

    scale_dxf = st.number_input(
        "Tỷ lệ Scale DXF",
        min_value=0.000001,
        value=1.0,
        step=0.1
    )

    geometry_mm = scale_geometry(
        geometry,
        scale_x=scale_dxf,
        scale_y=scale_dxf
    )


# ============================================================
# ĐIỀU CHỈNH KÍCH THƯỚC CUỐI CÙNG (TARGET DIMENSION)
# ============================================================

st.subheader(
    "4️⃣ Kích thước thành phẩm"
)

current_width, current_height = (
    geometry_size(
        geometry_mm
    )
)

fc1, fc2 = st.columns(2)

with fc1:

    final_width = st.number_input(
        "Chiều rộng thành phẩm (mm)",
        min_value=0.001,
        value=float(
            round(
                current_width,
                3
            )
        ),
        step=1.0
    )

with fc2:

    final_height = (
        current_height
        *
        final_width
        /
        current_width
        if current_width > 0
        else current_height
    )

    st.metric(
        "Chiều cao tính toán",
        f"{final_height:.3f} mm"
    )


# Scale geometry theo chiều rộng mục tiêu

if current_width > 0:

    final_scale = (
        final_width /
        current_width
    )

else:

    final_scale = 1.0


final_geometry = scale_geometry(
    geometry_mm,
    scale_x=final_scale,
    scale_y=final_scale
)


# ============================================================
# GỐC TỌA ĐỘ PHÔI / MÁY (WORK ZERO)
# ============================================================

st.subheader(
    "5️⃣ Gốc tọa độ máy / Phôi (Work Zero)"
)

work_zero = st.selectbox(
    "Gốc tọa độ gia công",
    [
        "Bottom Left",
        "Bottom Center",
        "Center",
        "Top Left",
        "Top Right"
    ]
)

final_geometry = apply_work_zero(
    final_geometry,
    work_zero
)

final_width, final_height = (
    geometry_size(
        final_geometry
    )
)

z1, z2 = st.columns(2)

with z1:

    st.metric(
        "Kích thước X cuối",
        f"{final_width:.3f} mm"
    )

with z2:

    st.metric(
        "Kích thước Y cuối",
        f"{final_height:.3f} mm"
    )


# ============================================================
# THÔNG SỐ CAM (CAM PARAMETERS)
# ============================================================

st.subheader(
    "6️⃣ Thông số gia công CAM"
)

cam1, cam2, cam3 = st.columns(3)

with cam1:

    material_thickness = st.number_input(
        "Độ dày vật liệu (mm)",
        min_value=0.01,
        value=3.0,
        step=0.1
    )

    cut_z = st.number_input(
        "Độ sâu cắt Cut Z (mm)",
        value=-3.2,
        step=0.1
    )

with cam2:

    safe_z = st.number_input(
        "Độ cao an toàn Safe Z (mm)",
        min_value=0.1,
        value=5.0,
        step=0.5
    )

    feed_rate = st.number_input(
        "Tốc độ cắt Feed Rate (mm/phút)",
        min_value=1,
        value=800,
        step=50
    )

with cam3:

    plunge_rate = st.number_input(
        "Tốc độ đâm dao Plunge Rate (mm/phút)",
        min_value=1,
        value=300,
        step=10
    )
