import streamlit as st
import cv2
import numpy as np
import ezdxf
from PIL import Image
import io
import os
from google import genai
from google.genai import types

# ----------------------------------------------------
# CONFIG & PAGE SETUP
# ----------------------------------------------------
st.set_page_config(page_title="AI CAD/CAM Dashboard", layout="wide")
st.title("🤖 AI CAD/CAM Dashboard: Image/DXF to DXF & G-Code")

# Khởi tạo Session State để lưu trữ dữ liệu giữa các lần làm mới giao diện
if 'processed_image' not in st.session_state:
    st.session_state.processed_image = None
if 'dxf_content' not in st.session_state:
    st.session_state.dxf_content = None
if 'gcode_content' not in st.session_state:
    st.session_state.gcode_content = None
if 'ai_thoughts' not in st.session_state:
    st.session_state.ai_thoughts = ""
if 'detected_contours' not in st.session_state:
    st.session_state.detected_contours = []
if 'scale_factor' not in st.session_state:
    st.session_state.scale_factor = 1.0

# Sidebar cho API Key
st.sidebar.header("🔑 Cấu hình API")
api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")

# ----------------------------------------------------
# HELPER FUNCTIONS (CAD / CAM / VISION)
# ----------------------------------------------------

def process_image_contours(img_bytes):
    """Sử dụng OpenCV để quét tất cả loại hình dạng (Contours) từ ảnh"""
    image = Image.open(io.BytesIO(img_bytes)).convert('RGB')
    img_np = np.array(image)
    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    
    # Khử nhiễu & Thresholding tự động
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    
    # Tìm tất cả contours (bao gồm mọi hình dạng phức tạp/hoa văn)
    contours, _ = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    return img_np, contours

def draw_dimensions_on_img(img_np, contours, target_width_mm=None):
    """Vẽ đường viền và thể hiện kích thước lên ảnh"""
    annotated = img_np.copy()
    if not contours:
        return annotated, 100.0, 100.0 # Default dimensions
    
    # Tìm Bounding Box cho toàn bộ vật thể
    all_pts = np.vstack([cnt for cnt in contours])
    x, y, w, h = cv2.boundingRect(all_pts)
    
    # Tính toán Scale nếu người dùng muốn tinh chỉnh chiều rộng cụ thể (mm)
    pixel_width = w
    pixel_height = h
    
    # Trực quan hóa Bounding Box và Kích thước
    cv2.rectangle(annotated, (x, y), (x + w, y + h), (0, 255, 0), 2)
    cv2.drawContours(annotated, contours, -1, (255, 0, 0), 1)
    
    label = f"W: {pixel_width}px, H: {pixel_height}px"
    if target_width_mm:
        scale = target_width_mm / pixel_width
        mm_h = pixel_height * scale
        label = f"W: {target_width_mm:.1f}mm, H: {mm_h:.1f}mm"
        
    cv2.putText(annotated, label, (x, max(y - 10, 20)), 
                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)
    return annotated, pixel_width, pixel_height

def generate_dxf(contours, scale=1.0):
    """Dựng lại file DXF từ Contours"""
    doc = ezdxf.new('R2010')
    msp = doc.modelspace()
    
    for cnt in contours:
        # Làm mượt đường cong bằng approxPolyDP nếu cần
        pts = cnt.reshape(-1, 2)
        if len(pts) > 1:
            # Chuyển đổi hệ tọa độ (OpenCV y hướng xuống, DXF y hướng lên)
            dxf_pts = [(float(p[0]) * scale, -float(p[1]) * scale) for p in pts]
            dxf_pts.append(dxf_pts[0]) # Đóng contour
            msp.add_lwpolyline(dxf_pts)
            
    out_stream = io.StringIO()
    doc.write(out_stream)
    return out_stream.getvalue()

def generate_gcode(contours, scale=1.0, feed_rate=1000, safe_z=5, cut_z=-1):
    """Xuất mã G-Code từ Contours"""
    gcode = [
        "G21 ; Set units to mm",
        "G90 ; Absolute positioning",
        f"G0 Z{safe_z} ; Rapid move to safe Z",
        "M3 S10000 ; Spindle ON"
    ]
    
    for cnt in contours:
        pts = cnt.reshape(-1, 2)
        if len(pts) <= 1:
            continue
            
        # Di chuyển tới điểm đầu
        start_x, start_y = pts[0][0] * scale, -pts[0][1] * scale
        gcode.append(f"G0 X{start_x:.3f} Y{start_y:.3f}")
        gcode.append(f"G1 Z{cut_z} F300 ; Cut down")
        
        # Cắt theo contour
        for p in pts[1:]:
            x, y = p[0] * scale, -p[1] * scale
            gcode.append(f"G1 X{x:.3f} Y{y:.3f} F{feed_rate}")
            
        # Cắt về điểm đầu và nhấc dao
        gcode.append(f"G1 X{start_x:.3f} Y{start_y:.3f} F{feed_rate}")
        gcode.append(f"G0 Z{safe_z}")
        
    gcode.append("M5 ; Spindle OFF")
    gcode.append("G0 X0 Y0 ; Return Home")
    gcode.append("M30 ; End of program")
    return "\n".join(gcode)

# ----------------------------------------------------
# 1) GIAO DIỆN DASHBOARD - INPUT FILE
# ----------------------------------------------------
st.subheader("1️⃣ Tải lên File Ảnh hoặc DXF")
col_input1, col_input2 = st.columns(2)

with col_input1:
    uploaded_image = st.file_uploader("Tải lên File Ảnh (PNG, JPG, JPEG)", type=["png", "jpg", "jpeg"])
with col_input2:
    uploaded_dxf = st.file_uploader("Tải lên File DXF", type=["dxf"])

# Tự động đọc file input nếu có
raw_img_np = None
if uploaded_image:
    img_bytes = uploaded_image.read()
    raw_img_np, contours = process_image_contours(img_bytes)
    st.session_state.detected_contours = contours
elif uploaded_dxf:
    st.info("Đã tải lên DXF. Hệ thống sẵn sàng trích xuất đường nét và tạo G-code.")

st.divider()

# ----------------------------------------------------
# 4) KHU VỰC ĐẦU RA VÀ TƯƠNG TÁC AI
# ----------------------------------------------------
st.subheader("2️⃣ Điều khiển & Tinh chỉnh AI")

# Tinh chỉnh thông số (Prompt / Yêu cầu kích thước)
user_prompt = st.text_input(
    "💬 Tương tác với AI để tinh chỉnh (VD: 'Tải bức ảnh lên muốn độ rộng của họa tiết hoa văn là 70mm'):",
    placeholder="Nhập kích thước hoặc yêu cầu đặc biệt..."
)

col_b1, col_b2, col_b3, col_b4 = st.columns(4)

with col_b1:
    btn_ai_process = st.button("🚀 1. AI Xử lý & Dựng lại", use_container_width=True)
with col_b2:
    btn_export_dxf = st.button("📐 2. Xuất File DXF", use_container_width=True)
with col_b3:
    btn_export_gcode = st.button("⚙️ 3. Xuất G-Code", use_container_width=True)
with col_b4:
    btn_ai_chat = st.button("💡 4. Tinh chỉnh theo Yêu cầu", use_container_width=True)

# ----------------------------------------------------
# 2) AI SUY NGHĨ & BỘ XỬ LÝ TRUNG TÂM
# ----------------------------------------------------
if btn_ai_process or btn_ai_chat:
    if not api_key:
        st.error("Vui lòng nhập Gemini API Key ở thanh bên (Sidebar) để kích hoạt AI!")
    elif uploaded_image is None and uploaded_dxf is None:
        st.warning("Vui lòng tải lên file ảnh hoặc DXF trước!")
    else:
        with st.spinner("AI đang suy nghĩ, phân tích kích thước và tối ưu hóa CAD/CAM..."):
            client = genai.Client(api_key=api_key)
            
            # Đọc kích thước từ Contour
            contours = st.session_state.detected_contours
            annotated_img, px_w, px_h = draw_dimensions_on_img(raw_img_np, contours)
            
            # Gửi thông tin cho Gemini AI "Suy nghĩ"
            prompt_text = f"""
            Bạn là một kỹ sư CAD/CAM AI cao cấp. 
            Ảnh/Vật thể đầu vào có kích thước pixel: Width={px_w}px, Height={px_h}px.
            Yêu cầu từ người dùng: "{user_prompt if user_prompt else 'Phân tích hình dạng và tối ưu hóa để xuất CAD/Gcode'}".
            
            Hãy suy nghĩ và thực hiện các bước sau:
            1. Phân tích hình dạng vật thể (Đường cong, họa tiết hoa văn, hay hình học đơn giản).
            2. Xác định tỉ lệ scale thích hợp dựa trên yêu cầu người dùng (Ví dụ nếu người dùng yêu cầu rộng 70mm, tính scale = 70 / {px_w}).
            3. Đưa ra hướng dẫn tạo file DXF và thông số cắt G-Code (Feedrate, Safe Z, Cut Z).
            Trả về câu trả lời phân tích ngắn gọn, chuyên nghiệp và trích xuất rõ ràng giá trị scale_mm (Ví dụ: SCALE_MM: 0.123).
            """
            
            try:
                # Gọi Model Gemini 2.5
                response = client.models.generate_content(
                    model='gemini-1.5-flash',
                    contents=[prompt_text]
                )
                
                st.session_state.ai_thoughts = response.text
                
                # Trích xuất chiều rộng mong muốn từ user_prompt nếu có (VD: 70mm)
                import re
                match = re.search(r'(\d+(\.\d+)?)\s*mm', user_prompt.lower())
                if match:
                    target_w = float(match.group(1))
                    scale = target_w / px_w
                else:
                    scale = 1.0 # Mặc định 1px = 1mm nếu không chỉ định
                    
                st.session_state.scale_factor = scale
                
                # Tiến hành tái tạo DXF & G-Code dựa trên suy nghĩ của AI
                st.session_state.dxf_content = generate_dxf(contours, scale=scale)
                st.session_state.gcode_content = generate_gcode(contours, scale=scale)
                
                # Cập nhật ảnh có thể hiện kích thước mm
                final_annotated, _, _ = draw_dimensions_on_img(raw_img_np, contours, target_width_mm=px_w*scale)
                st.session_state.processed_image = final_annotated
                
                st.success("AI đã hoàn tất suy nghĩ và dựng lại mô hình!")
                
            except Exception as e:
                st.error(f"Lỗi khi kết nối với AI: {e}")

# ----------------------------------------------------
# 3) HIỂN THỊ KẾT QUẢ SO SÁNH & ĐẦU RA
# ----------------------------------------------------
st.divider()
st.subheader("3️⃣ Kết quả Tối ưu & So sánh Kích thước")

col_img1, col_img2 = st.columns(2)

with col_img1:
    st.markdown("**🖼️ Ảnh Gốc Tải Lên**")
    if uploaded_image:
        st.image(uploaded_image, use_container_width=True)
    else:
        st.info("Chưa có ảnh gốc.")

with col_img2:
    st.markdown("**🎯 Ảnh Sau Xử Lý (Thể hiện Kích thước & Contour)**")
    if st.session_state.processed_image is not None:
        st.image(st.session_state.processed_image, use_container_width=True)
    else:
        st.info("Nhấn 'AI Xử lý' để xem kết quả.")

# Hiển thị suy nghĩ của AI
if st.session_state.ai_thoughts:
    with st.expander("🧠 Xem AI Suy Nghĩ & Phân Tích Chi Tiết", expanded=True):
        st.write(st.session_state.ai_thoughts)

# Nút Tải Xuất File DXF và G-Code
st.divider()
col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if st.session_state.dxf_content:
        st.download_button(
            label="💾 Tải về File DXF",
            data=st.session_state.dxf_content,
            file_name="output_ai_cad.dxf",
            mime="application/dxf",
            use_container_width=True
        )

with col_dl2:
    if st.session_state.gcode_content:
        st.download_button(
            label="💾 Tải về File G-Code (.nc)",
            data=st.session_state.gcode_content,
            file_name="output_ai_cam.nc",
            mime="text/plain",
            use_container_width=True
        )
