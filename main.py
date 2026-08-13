import cv2
import ezdxf
import numpy as np
from skimage.morphology import skeletonize
from shapely.geometry import LineString, MultiLineString


def process_pattern_with_fixed_thickness(
    image_bytes, target_w_mm, target_h_mm, stroke_width_mm=70.0
):
    """Trích xuất đường xương (Skeleton) của hoa văn và nới rộng đều 2 bên để đạt độ rộng đúng 70mm."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Binarize (Nhi phân hóa ảnh)
    _, thresh = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )
    h_px, w_px = thresh.shape

    scale_x = target_w_mm / w_px
    scale_y = target_h_mm / h_px

    # 2. Rút xương hoa văn (Biến các nét dày thành đường nét đơn 1-pixel ở giữa)
    binary_bool = thresh > 0
    skeleton = skeletonize(binary_bool)
    skeleton_uint8 = (skeleton * 255).astype(np.uint8)

    # 3. Tìm các đường nét xương (Centerlines)
    contours, _ = cv2.findContours(
        skeleton_uint8, cv2.RETR_LIST, cv2.CHAIN_APPROX_TC89_KCOS
    )

    offset_distance = stroke_width_mm / 2.0  # Lấy 35mm mỗi bên
    pattern_polygons = []

    for cnt in contours:
        if len(cnt) < 3:
            continue

        # Làm mượt đường xương
        epsilon = 0.002 * cv2.arcLength(cnt, False)
        approx = cv2.approxPolyDP(cnt, epsilon, False)

        pts = []
        for p in approx:
            x_px, y_px = p[0][0], p[0][1]
            pts.append((x_px * scale_x, (h_px - y_px) * scale_y))

        if len(pts) >= 2:
            line = LineString(pts)
            # 4. Dùng Shapely tạo dải bọc (Buffer) đúng 35mm về 2 phía -> Tổng độ rộng = 70mm
            buffered_poly = line.buffer(
                offset_distance, cap_style=1, join_style=1
            )

            if buffered_poly.is_valid and not buffered_poly.is_empty:
                # Lấy tọa độ đường viền sau khi đã ép độ rộng 70mm
                if buffered_poly.geom_type == "Polygon":
                    coords = list(buffered_poly.exterior.coords)
                    pattern_polygons.append(coords)
                elif buffered_poly.geom_type == "MultiPolygon":
                    for poly in buffered_poly.geoms:
                        coords = list(poly.exterior.coords)
                        pattern_polygons.append(coords)

    return pattern_polygons
