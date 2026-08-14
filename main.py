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
# CONFIG
# ============================================================

st.set_page_config(
    page_title="AI CAD/CAM Dashboard",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 AI CAD/CAM Dashboard")
st.caption("Image / DXF → Geometry → DXF / G-Code")


# ============================================================
# SESSION STATE
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
# DATA STRUCTURE
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
# GEOMETRY HELPERS
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
    """Trả về Width / Height."""

    min_x, min_y, max_x, max_y = geometry_bounds(geometry)

    return (
        max_x - min_x,
        max_y - min_y
    )


def polygon_area(points):
    """
    Shoelace formula.
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
# IMAGE PROCESSING
# ============================================================

def process_image_contours(
    img_bytes,
    min_area_px=20,
    threshold_mode="OTSU"
):
    """
    Image → OpenCV contours.
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
    # Blur
    # -----------------------------------------

    blurred = cv2.GaussianBlur(
        gray,
        (5, 5),
        0
    )

    # -----------------------------------------
    # Threshold
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
    # Morphological cleanup
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
    # Find contours
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
        # Determine hole using hierarchy
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
# GEOMETRY SIMPLIFICATION
# ============================================================

def simplify_geometry(
    geometry,
    tolerance_px=1.0
):
    """
    Reduce jumlah vertex.

    tolerance_px:
        tolerance dalam pixel đối với Image.
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
# DXF READER
# ============================================================

def read_dxf_file(dxf_bytes):
    """
    Đọc các entity cơ bản từ DXF:

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
        # LINE
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
        # LWPOLYLINE
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
        # POLYLINE
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
        # CIRCLE
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
        # ARC
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
# IMAGE COORDINATE CONVERSION
# ============================================================

def convert_image_geometry_to_mm(
    geometry,
    scale_mm_per_pixel
):
    """
    Image:
        X → X
        Y ↓

    Machine/CAD:
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
# SCALE DXF
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
# WORK ZERO
# ============================================================

def apply_work_zero(
    geometry,
    work_zero
):
    """
    Align geometry to selected machine origin.
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
# SORT CUT ORDER
# ============================================================

def sort_geometry_for_cutting(
    geometry
):
    """
    Inner contours first.
    Outer contours last.

    Nếu không xác định được hole,
    dùng area nhỏ → lớn.
    """

    def sort_key(item):

        area = polygon_area(
            item["points"]
        )

        # Hole ưu tiên trước
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
# DXF GENERATOR
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
    # Layers
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
# G-CODE GENERATOR
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

    Inner contours được cắt trước outer contours.
    """

    geometry = sort_geometry_for_cutting(
        geometry
    )

    lines = []

    lines.append(
        "; ========================================="
    )
    lines.append(
        "; AI CAD/CAM Generated G-Code"
    )
    lines.append(
        "; ========================================="
    )

    lines.append(
        "G21 ; Units: mm"
    )

    lines.append(
        "G90 ; Absolute positioning"
    )

    lines.append(
        "G17 ; XY plane"
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
            f"; --- Contour {index} ---"
        )

        if item["is_hole"]:

            lines.append(
                "; Inner contour / hole"
            )

        else:

            lines.append(
                "; Outer contour"
            )

        # -------------------------------------
        # Move to start
        # -------------------------------------

        lines.append(
            f"G0 X{start_x:.3f} "
            f"Y{start_y:.3f}"
        )

        # -------------------------------------
        # Plunge
        # -------------------------------------

        lines.append(
            f"G1 Z{cut_z:.3f} "
            f"F{plunge_rate}"
        )

        # -------------------------------------
        # Cutting
        # -------------------------------------

        for x, y in points[1:]:

            lines.append(
                f"G1 X{x:.3f} "
                f"Y{y:.3f} "
                f"F{feed_rate}"
            )

        # -------------------------------------
        # Close contour
        # -------------------------------------

        if item["closed"]:

            lines.append(
                f"G1 X{start_x:.3f} "
                f"Y{start_y:.3f} "
                f"F{feed_rate}"
            )

        # -------------------------------------
        # Retract
        # -------------------------------------

        lines.append(
            f"G0 Z{safe_z:.3f}"
        )

        lines.append("")

    # -----------------------------------------
    # End
    # -----------------------------------------

    lines.append(
        "M5 ; Spindle OFF"
    )

    lines.append(
        "G0 Z%.3f" % safe_z
    )

    lines.append(
        "G0 X0 Y0 ; Return Home"
    )

    lines.append(
        "M30 ; Program End"
    )

    return "\n".join(lines)


# ============================================================
# PREVIEW
# ============================================================

def create_geometry_preview(
    geometry,
    title="Geometry Preview"
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
# ANNOTATED IMAGE
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
    # Bounding box
    # -----------------------------------------

    cv2.rectangle(
        annotated,
        (x, y),
        (x + w, y + h),
        (0, 255, 0),
        2
    )

    # -----------------------------------------
    # Contours
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
    # Label
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
            f"W: {target_width_mm:.2f} mm | "
            f"H: {height_mm:.2f} mm"
        )

    else:

        label = (
            f"W: {w}px | H: {h}px"
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
        gemini-1.5-flash,
        contents=prompt
    )

    return response.text


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header(
    "🔑 AI Configuration"
)

api_key = st.sidebar.text_input(
    "Gemini API Key",
    type="password"
)

st.sidebar.divider()

st.sidebar.header(
    "⚙️ Image Processing"
)

min_area_px = st.sidebar.number_input(
    "Minimum contour area (px²)",
    min_value=1,
    value=20,
    step=1
)

threshold_mode = st.sidebar.selectbox(
    "Threshold Mode",
    [
        "OTSU",
        "ADAPTIVE",
        "FIXED"
    ]
)

simplify_tolerance = st.sidebar.number_input(
    "Contour tolerance (pixel)",
    min_value=0.0,
    value=1.0,
    step=0.1
)


# ============================================================
# INPUT
# ============================================================

st.subheader(
    "1️⃣ Input"
)

col1, col2 = st.columns(2)

with col1:

    uploaded_image = st.file_uploader(
        "🖼️ Upload Image",
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
        "📐 Upload DXF",
        type=["dxf"]
    )


# ============================================================
# LOAD IMAGE
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
            f"Đã nhận Image: "
            f"{len(geometry)} contour"
        )

    except Exception as e:

        st.error(
            f"Lỗi xử lý Image: {e}"
        )


# ============================================================
# LOAD DXF
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
            f"Đã đọc DXF: "
            f"{len(geometry)} entities"
        )

    except Exception as e:

        st.error(
            f"Lỗi đọc DXF: {e}"
        )


# ============================================================
# STOP IF NO INPUT
# ============================================================

geometry = st.session_state.geometry

if not geometry:

    st.info(
        "Vui lòng upload Image hoặc DXF."
    )

    st.stop()


# ============================================================
# SOURCE INFORMATION
# ============================================================

st.divider()

st.subheader(
    "2️⃣ Geometry Information"
)

width, height = geometry_size(
    geometry
)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "Entities",
        len(geometry)
    )

with c2:
    st.metric(
        "Source Width",
        f"{width:.3f}"
    )

with c3:
    st.metric(
        "Source Height",
        f"{height:.3f}"
    )

with c4:

    st.metric(
        "Source",
        st.session_state.source_type
    )


# ============================================================
# DIMENSION / SCALE
# ============================================================

st.subheader(
    "3️⃣ Dimension & Scale"
)

if st.session_state.source_type == "image":

    st.info(
        "Image được quy đổi từ pixel sang mm. "
        "Bạn cần cung cấp kích thước thực tế."
    )

    target_width_mm = st.number_input(
        "Target Width (mm)",
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
        f"**Scale:** "
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
        "DXF Scale",
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
# TARGET DIMENSION ADJUSTMENT
# ============================================================

st.subheader(
    "4️⃣ Final Dimension"
)

current_width, current_height = (
    geometry_size(
        geometry_mm
    )
)

fc1, fc2 = st.columns(2)

with fc1:

    final_width = st.number_input(
        "Final Width (mm)",
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
        "Calculated Height",
        f"{final_height:.3f} mm"
    )


# Scale geometry to target width

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
# WORK ZERO
# ============================================================

st.subheader(
    "5️⃣ Work Zero / Origin"
)

work_zero = st.selectbox(
    "Machine Work Zero",
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
        "Final X Size",
        f"{final_width:.3f} mm"
    )

with z2:

    st.metric(
        "Final Y Size",
        f"{final_height:.3f} mm"
    )


# ============================================================
# CAM PARAMETERS
# ============================================================

st.subheader(
    "6️⃣ CAM Parameters"
)

cam1, cam2, cam3 = st.columns(3)

with cam1:

    material_thickness = st.number_input(
        "Material Thickness (mm)",
        min_value=0.01,
        value=3.0,
        step=0.1
    )

    cut_z = st.number_input(
        "Cut Z (mm)",
        value=-3.2,
        step=0.1
    )

with cam2:

    safe_z = st.number_input(
        "Safe Z (mm)",
        min_value=0.1,
        value=5.0,
        step=0.5
    )

    feed_rate = st.number_input(
        "Feed Rate (mm/min)",
        min_value=1,
        value=800,
        step=50
    )

with cam3:

    plunge_rate = st.number_input(
        "Plunge Rate (mm/min)",
        min_value=1,
        value=300,
        step=25
    )

    spindle_rpm = st.number_input(
        "Spindle RPM",
        min_value=0,
        value=12000,
        step=500
    )


# ============================================================
# AI
# ============================================================

st.subheader(
    "7️⃣ AI CAD/CAM Assistant"
)

user_prompt = st.text_area(
    "Yêu cầu AI",
    placeholder=(
        "Ví dụ: Chi tiết này cần rộng 70mm, "
        "hãy kiểm tra scale và thông số cắt."
    )
)

if st.button(
    "🤖 AI Analyze",
    use_container_width=True
):

    if not api_key:

        st.warning(
            "Vui lòng nhập Gemini API Key."
        )

    else:

        with st.spinner(
            "AI đang phân tích..."
        ):

            try:

                ai_result = ask_gemini(
                    api_key,
                    final_width,
                    final_height,
                    user_prompt
                )

                st.session_state.ai_thoughts = (
                    ai_result
                )

            except Exception as e:

                st.error(
                    f"Gemini Error: {e}"
                )


if st.session_state.ai_thoughts:

    with st.expander(
        "🧠 AI Analysis",
        expanded=True
    ):

        st.markdown(
            st.session_state.ai_thoughts
        )


# ============================================================
# PREVIEW
# ============================================================

st.divider()

st.subheader(
    "8️⃣ Geometry Preview"
)

preview_col1, preview_col2 = (
    st.columns(2)
)

with preview_col1:

    if (
        st.session_state.source_type
        == "image"
        and st.session_state.source_image
        is not None
    ):

        annotated = draw_dimensions_on_img(
            st.session_state.source_image,
            geometry,
            target_width_mm=final_width
        )

        st.image(
            annotated,
            caption="Image + Detected Geometry",
            use_container_width=True
        )

    else:

        fig1 = create_geometry_preview(
            geometry_mm,
            title="Original DXF Geometry"
        )

        st.pyplot(
            fig1,
            use_container_width=True
        )

with preview_col2:

    fig2 = create_geometry_preview(
        final_geometry,
        title=(
            f"Final CAM Geometry "
            f"({work_zero})"
        )
    )

    st.pyplot(
        fig2,
        use_container_width=True
    )


# ============================================================
# GENERATE OUTPUT
# ============================================================

st.divider()

st.subheader(
    "9️⃣ Generate Output"
)

out1, out2 = st.columns(2)

# ============================================================
# DXF
# ============================================================

with out1:

    if st.button(
        "📐 Generate DXF",
        use_container_width=True
    ):

        try:

            dxf_content = generate_dxf(
                final_geometry
            )

            st.session_state.dxf_content = (
                dxf_content
            )

            st.success(
                "DXF đã được tạo."
            )

        except Exception as e:

            st.error(
                f"Lỗi tạo DXF: {e}"
            )


# ============================================================
# G-CODE
# ============================================================

with out2:

    if st.button(
        "⚙️ Generate G-Code",
        use_container_width=True
    ):

        try:

            # -----------------------------------------
            # Safety check
            # -----------------------------------------

            if cut_z >= 0:

                st.warning(
                    "Cut Z đang >= 0. "
                    "Kiểm tra lại trước khi chạy máy."
                )

            if cut_z <= -material_thickness * 2:

                st.warning(
                    "Cut Z sâu bất thường so với "
                    "độ dày vật liệu."
                )

            gcode_content = generate_gcode(
                final_geometry,
                feed_rate=feed_rate,
                plunge_rate=plunge_rate,
                safe_z=safe_z,
                cut_z=cut_z,
                spindle_rpm=spindle_rpm
            )

            st.session_state.gcode_content = (
                gcode_content
            )

            st.success(
                "G-Code đã được tạo."
            )

        except Exception as e:

            st.error(
                f"Lỗi tạo G-Code: {e}"
            )


# ============================================================
# DOWNLOAD
# ============================================================

st.subheader(
    "🔽 Download"
)

dl1, dl2 = st.columns(2)

with dl1:

    if st.session_state.dxf_content:

        st.download_button(
            label="💾 Download DXF",
            data=st.session_state.dxf_content,
            file_name="output_ai_cad.dxf",
            mime="application/dxf",
            use_container_width=True
        )

with dl2:

    if st.session_state.gcode_content:

        st.download_button(
            label="💾 Download G-Code",
            data=st.session_state.gcode_content,
            file_name="output_ai_cam.nc",
            mime="text/plain",
            use_container_width=True
        )


# ============================================================
# G-CODE PREVIEW
# ============================================================

if st.session_state.gcode_content:

    st.divider()

    st.subheader(
        "⚙️ G-Code Preview"
    )

    st.code(
        st.session_state.gcode_content,
        language="gcode"
    )


# ============================================================
# FINAL SAFETY NOTICE
# ============================================================

st.divider()

st.warning(
    """
⚠️ SAFETY CHECK

G-Code được tạo tự động chỉ nên được xem là dữ liệu CAM.
Trước khi chạy máy thực tế cần kiểm tra:

• Work Zero
• Chiều X/Y
• Chiều sâu Z
• Safe Z
• Feed Rate
• Plunge Rate
• Spindle RPM
• Hướng chạy dao
• Kích thước vật liệu
• Tool diameter
• Collision / travel limits

Luôn chạy simulation hoặc dry-run trước khi cắt vật liệu thật.
"""
)
