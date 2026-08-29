import io
import json
import calendar
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from google import genai

# ============================================================
# CẤU HÌNH & KHỞI TẠO THƯ MỤC
# ============================================================

# SỬ DỤNG ST.SECRETS ĐỂ BẢO MẬT TOKEN
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()
TELEGRAM_BOT_TOKEN = st.secrets.get("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = str(st.secrets.get("TELEGRAM_CHAT_ID", "")).strip()

# Thư mục chứa báo cáo xuất ra
REPORT_DIR = Path("reports")
REPORT_DIR.mkdir(exist_ok=True)
LAST_SENT_FILE = REPORT_DIR / "last_sent.txt"

# Thư mục chứa dữ liệu cấu hình/database Excel
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
CURRENT_DB_FILE = DATA_DIR / "current_hr_data.xlsx"

# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="Quản lý sinh nhật & hợp đồng",
    page_icon="📊",
    layout="wide",
)

# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 10px;
    }
    .sub-title {
        color: #666;
        margin-bottom: 25px;
    }
    .metric-box {
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #ddd;
        background: #fafafa;
        text-align: center;
    }
    .metric-number {
        font-size: 30px;
        font-weight: 700;
    }
    .metric-label {
        font-size: 15px;
        color: #666;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HÀM GEMINI 2.5 PRO
# ============================================================

GEMINI_MODEL = "gemini-2.5-pro"


def get_gemini_client():
    """
    Khởi tạo Gemini Client bằng API Key trong Streamlit Secrets.
    """
    if not GEMINI_API_KEY:
        return None

    try:
        return genai.Client(api_key=GEMINI_API_KEY)
    except Exception as e:
        st.error(f"Không thể khởi tạo Gemini Client: {e}")
        return None


def detect_columns_with_gemini(columns):
    """
    Sử dụng Gemini 2.5 Pro để nhận diện cấu trúc cột Excel nhân sự.

    Gemini sẽ phân tích tên cột và trả về mapping:
        ma_nv
        ho_ten
        ngay_sinh
        bo_phan
        ngay_het_han_hop_dong

    Kết quả được yêu cầu ở dạng JSON chuẩn.
    """

    client = get_gemini_client()

    if client is None:
        return None

    column_list = "\n".join(
        [f"- {str(col)}" for col in columns]
    )

    prompt = f"""
Bạn là chuyên gia phân tích dữ liệu nhân sự và Excel.

Nhiệm vụ:
Phân tích danh sách tên cột Excel bên dưới và xác định cột nào
tương ứng với các trường dữ liệu nhân sự của hệ thống.

DANH SÁCH CỘT EXCEL:
{column_list}

CÁC TRƯỜNG CẦN NHẬN DIỆN:

1. ma_nv
   - Mã nhân viên
   - Employee ID
   - Employee Code
   - Staff ID
   - Mã NV

2. ho_ten
   - Họ và tên
   - Họ tên
   - Tên nhân viên
   - Full Name
   - Employee Name

3. ngay_sinh
   - Ngày sinh
   - Ngày tháng năm sinh
   - DOB
   - Date of Birth
   - Birthday

4. bo_phan
   - Bộ phận
   - Phòng ban
   - Đơn vị
   - Department
   - Division

5. ngay_het_han_hop_dong
   - Ngày hết hạn hợp đồng
   - Ngày kết thúc hợp đồng
   - Ngày hết hạn HĐ
   - Contract End Date
   - Contract Expiry Date

QUY TẮC:

- Chọn đúng tên cột thực tế trong danh sách.
- Không được tự tạo tên cột mới.
- Nếu không tìm thấy cột tương ứng thì trả về null.
- Nếu có nhiều cột có vẻ phù hợp, hãy chọn cột có ý nghĩa chính xác nhất.
- Không thay đổi tên cột.
- Không giải thích.
- Chỉ trả về JSON hợp lệ.
"""

    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=genai.types.GenerateContentConfig(
                temperature=0.1,
                max_output_tokens=1000,
                response_mime_type="application/json",
            ),
        )

        text = response.text.strip()

        if not text:
            st.warning("Gemini không trả về dữ liệu.")
            return None

        result = json.loads(text)

        # ----------------------------------------------------
        # Kiểm tra kết quả Gemini
        # ----------------------------------------------------

        required_keys = [
            "ma_nv",
            "ho_ten",
            "ngay_sinh",
            "bo_phan",
            "ngay_het_han_hop_dong",
        ]

        for key in required_keys:
            if key not in result:
                result[key] = None

        # ----------------------------------------------------
        # BẢO VỆ: Gemini chỉ được phép trả về tên cột
        # thực sự tồn tại trong Excel
        # ----------------------------------------------------

        valid_columns = set(columns)

        for key in required_keys:
            value = result.get(key)

            if value is not None:
                value = str(value).strip()

                if value not in valid_columns:
                    result[key] = None
                else:
                    result[key] = value

        return result

    except json.JSONDecodeError as e:
        st.warning(
            f"Gemini trả về JSON không hợp lệ: {e}"
        )
        return None

    except Exception as e:
        st.warning(
            f"Gemini 2.5 Pro không thể nhận diện cột: {e}"
        )
        return None

# ============================================================
# NHẬN DIỆN CỘT BẰNG QUY TẮC DỰ PHÒNG
# ============================================================

def normalize_text(text):
    return str(text).strip().lower().replace("_", " ").replace("-", " ")

def find_column_by_keywords(columns, keywords):
    normalized = {col: normalize_text(col) for col in columns}
    # Ưu tiên khớp chính xác
    for col, norm in normalized.items():
        if norm in keywords: return col
    # Sau đó kiểm tra chứa từ khóa
    for col, norm in normalized.items():
        for keyword in keywords:
            if keyword in norm: return col
    return None

def detect_columns_fallback(columns):
    mapping = {}
    mapping["ma_nv"] = find_column_by_keywords(columns, ["mã nv", "ma nv", "mã nhân viên", "ma nhan vien", "employee id", "employee code"])
    mapping["ho_ten"] = find_column_by_keywords(columns, ["họ và tên", "ho va ten", "họ tên", "ho ten", "tên nhân viên", "ten nhan vien", "full name"])
    mapping["ngay_sinh"] = find_column_by_keywords(columns, ["ngày sinh", "ngay sinh", "ns", "date of birth", "dob", "birthday"])
    mapping["bo_phan"] = find_column_by_keywords(columns, ["bộ phận", "bo phan", "phòng ban", "phong ban", "đơn vị", "don vi", "department"])
    mapping["ngay_het_han_hop_dong"] = find_column_by_keywords(columns, ["ngày hết hạn hợp đồng", "ngay het han hop dong", "ngày hết hạn hđ", "ngày kết thúc hợp đồng", "contract end date"])
    return mapping

# ============================================================
# ĐỌC EXCEL
# ============================================================

def read_excel(file_path_or_buffer):
    try:
        df = pd.read_excel(file_path_or_buffer, engine="openpyxl")
        df = df.dropna(how="all").copy()
        df.columns = [str(col).strip() for col in df.columns]
        return df
    except Exception as e:
        st.error(f"Không thể đọc file Excel: {e}")
        return None

# ============================================================
# CHUẨN HÓA DATAFRAME
# ============================================================

def prepare_dataframe(df, mapping):
    required = ["ma_nv", "ho_ten", "ngay_sinh", "bo_phan", "ngay_het_han_hop_dong"]
    missing = [key for key in required if not mapping.get(key)]
    if missing:
        return None, missing

    result = pd.DataFrame()
    result["Mã NV"] = df[mapping["ma_nv"]]
    result["Họ và tên"] = df[mapping["ho_ten"]]
    result["Ngày tháng năm sinh"] = pd.to_datetime(df[mapping["ngay_sinh"]], errors="coerce", dayfirst=True)
    result["Bộ Phận"] = df[mapping["bo_phan"]]
    result["Ngày hết hạn hợp đồng"] = pd.to_datetime(df[mapping["ngay_het_han_hop_dong"]], errors="coerce", dayfirst=True)
    
    return result, []

# ============================================================
# LỌC DỮ LIỆU
# ============================================================

def get_birthday_report(df, month):
    result = df[df["Ngày tháng năm sinh"].dt.month == month].copy()
    result = result[["Mã NV", "Họ và tên", "Ngày tháng năm sinh", "Bộ Phận"]]
    return result.sort_values(by="Ngày tháng năm sinh")

def get_contract_report(df, month, year):
    result = df[(df["Ngày hết hạn hợp đồng"].dt.month == month) & (df["Ngày hết hạn hợp đồng"].dt.year == year)].copy()
    result = result[["Mã NV", "Họ và tên", "Ngày tháng năm sinh", "Bộ Phận", "Ngày hết hạn hợp đồng"]]
    return result.sort_values(by="Ngày hết hạn hợp đồng")

# ============================================================
# TẠO EXCEL ĐỂ XUẤT
# ============================================================

def dataframe_to_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        for row in worksheet.iter_rows():
            for cell in row:
                if hasattr(cell.value, "strftime"):
                    cell.number_format = "dd/mm/yyyy"
        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                try: 
                    max_length = max(max_length, len(str(cell.value)))
                except Exception: 
                    pass
            worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 40)
    output.seek(0)
    return output

def save_excel_file(df, filename, sheet_name):
    path = REPORT_DIR / filename
    output = dataframe_to_excel(df, sheet_name)
    with open(path, "wb") as f:
        f.write(output.getvalue())
    return path

# ============================================================
# TELEGRAM
# ============================================================

def telegram_send_message(message):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID: 
        return False, "Thiếu Token hoặc Chat ID"
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        resp = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "text": message}, timeout=30)
        return (True, "OK") if resp.ok else (False, resp.text)
    except Exception as e: 
        return False, str(e)

def telegram_send_file(file_path, caption):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID: 
        return False, "Thiếu Token hoặc Chat ID"
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    try:
        with open(file_path, "rb") as f:
            resp = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption}, files={"document": f}, timeout=60)
        return (True, "OK") if resp.ok else (False, resp.text)
    except Exception as e: 
        return False, str(e)

def telegram_test():
    if not TELEGRAM_BOT_TOKEN: 
        return False, "Chưa nhập TELEGRAM_BOT_TOKEN"
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getMe"
    try:
        resp = requests.get(url, timeout=15)
        if resp.ok and resp.json().get("ok"): 
            return True, resp.json()["result"].get("username", "")
        return False, resp.text
    except Exception as e: 
        return False, str(e)

def send_monthly_report(birthday_df, contract_df, month, year):
    birthday_file = save_excel_file(birthday_df, "danh_sach_sinh_nhat.xlsx", "Sinh nhật")
    contract_file = save_excel_file(contract_df, "danh_sach_den_han_ky_hop_dong.xlsx", "Hợp đồng")

    message = f"""
📊 BÁO CÁO NHÂN SỰ THÁNG {month:02d}/{year}

🎂 Sinh nhật trong tháng: {len(birthday_df)} nhân viên
📄 Hợp đồng hết hạn trong tháng: {len(contract_df)} nhân viên

⏰ Thời gian báo cáo: {datetime.now().strftime("%d/%m/%Y %H:%M")}
"""
    ok, error = telegram_send_message(message.strip())
    if not ok: return False, error

    ok, error = telegram_send_file(birthday_file, f"🎂 Danh sách sinh nhật tháng {month:02d}/{year}")
    if not ok: return False, error

    ok, error = telegram_send_file(contract_file, f"📄 Danh sách hợp đồng hết hạn tháng {month:02d}/{year}")
    if not ok: return False, error

    return True, "Đã gửi thành công"

# ============================================================
# CHỐNG GỬI TRÙNG
# ============================================================

def get_last_sent():
    if not LAST_SENT_FILE.exists(): 
        return ""
    try: 
        return LAST_SENT_FILE.read_text(encoding="utf-8").strip()
    except: 
        return ""

def set_last_sent(month, year):
    LAST_SENT_FILE.write_text(f"{year:04d}-{month:02d}", encoding="utf-8")

def already_sent(month, year):
    return get_last_sent() == f"{year:04d}-{month:02d}"

# ============================================================
# KIỂM TRA TỰ ĐỘNG GỬI
# ============================================================

def automatic_monthly_check(birthday_df, contract_df):
    now = datetime.now()
    if now.day != 1 or now.hour < 10: 
        return
    month, year = now.month, now.year
    if already_sent(month, year): 
        return

    success, message = send_monthly_report(birthday_df, contract_df, month, year)
    if success:
        set_last_sent(month, year)
        st.success("🤖 Đã tự động gửi báo cáo Telegram.")
    else:
        st.error("Không thể tự động gửi Telegram: " + message)

# ============================================================
# GIAO DIỆN
# ============================================================

st.markdown('<div class="main-title">📊 QUẢN LÝ SINH NHẬT & HỢP ĐỒNG</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Theo dõi sinh nhật nhân viên và hợp đồng đến hạn ký</div>', unsafe_allow_html=True)

# SIDEBAR
with st.sidebar:
    st.header("⚙️ Cấu hình")
    current_date = date.today()
    selected_year = st.selectbox("Năm", list(range(current_date.year - 2, current_date.year + 6)), index=2)
    selected_month = st.selectbox("Tháng", list(range(1, 13)), index=current_date.month - 1, format_func=lambda x: f"Tháng {x:02d}")
    
    st.divider()
    st.subheader("🤖 Gemini AI")
    if GEMINI_API_KEY: 
        st.success("🟢 Gemini API đã cấu hình")
    else: 
        st.warning("🔴 Chưa có Gemini API Key")
    
    st.subheader("📱 Telegram")
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        st.success("🟢 Telegram đã cấu hình")
    else: 
        st.warning("🔴 Chưa cấu hình Telegram")

# ============================================================
# 1. QUẢN LÝ FILE DỮ LIỆU
# ============================================================
st.header("1️⃣ File danh sách nhân sự")

uploaded_file = st.file_uploader("Cập nhật file Excel danh sách nhân sự (Nếu có nhân viên mới/cũ)", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Lưu đè file mới để hệ thống ghi nhớ
    with open(CURRENT_DB_FILE, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success("✅ Đã cập nhật cơ sở dữ liệu nhân sự mới nhất!")

# Nếu chưa upload bao giờ (file không tồn tại), hệ thống mới dừng lại
if not CURRENT_DB_FILE.exists():
    st.info("📁 Hệ thống trống. Hãy upload file Excel danh sách nhân sự lần đầu để hệ thống bắt đầu hoạt động.")
    st.stop()

# Đọc file dữ liệu đã được lưu
df_original = read_excel(CURRENT_DB_FILE)

if df_original is None:
    st.stop()

st.success(f"Đang sử dụng dữ liệu gồm {len(df_original):,} nhân sự.")

with st.expander("👁️ Xem dữ liệu Excel gốc đang lưu"):
    st.dataframe(df_original, use_container_width=True, height=300)

# ============================================================
# NHẬN DIỆN CỘT
# ============================================================
st.header("2️⃣ Nhận diện cấu trúc Excel")

if "column_mapping" not in st.session_state:
    st.session_state.column_mapping = None

if st.button("🤖 Phân tích cột bằng Gemini AI", type="primary"):
    if not GEMINI_API_KEY:
        st.error("Bạn chưa cấu hình GEMINI_API_KEY")
    else:
        with st.spinner("Gemini đang phân tích cấu trúc Excel..."):
            mapping = detect_columns_with_gemini(list(df_original.columns))
            st.session_state.column_mapping = mapping

# Nếu chưa chạy Gemini, sử dụng fallback
if st.session_state.column_mapping is None:
    mapping = detect_columns_fallback(list(df_original.columns))
else:
    mapping = st.session_state.column_mapping

# Kiểm tra mapping
mapping_names = {
    "ma_nv": "Mã nhân viên", 
    "ho_ten": "Họ và tên", 
    "ngay_sinh": "Ngày sinh", 
    "bo_phan": "Bộ phận", 
    "ngay_het_han_hop_dong": "Ngày hết hạn hợp đồng"
}

mapping_df = pd.DataFrame([{"Trường hệ thống": mapping_names[key], "Cột Excel": mapping.get(key)} for key in mapping_names])
st.dataframe(mapping_df, use_container_width=True, hide_index=True)

prepared_df, missing = prepare_dataframe(df_original, mapping)

if missing:
    st.error("Không tìm thấy các trường bắt buộc: " + ", ".join(mapping_names[x] for x in missing))
    st.stop()

# ============================================================
# LỌC & DASHBOARD
# ============================================================
birthday_df = get_birthday_report(prepared_df, selected_month)
contract_df = get_contract_report(prepared_df, selected_month, selected_year)

st.header(f"3️⃣ Báo cáo tháng {selected_month:02d}/{selected_year}")

col1, col2, col3 = st.columns(3)
with col1: st.metric("🎂 Sinh nhật", f"{len(birthday_df):,}")
with col2: st.metric("📄 Hợp đồng hết hạn", f"{len(contract_df):,}")
with col3: st.metric("👥 Tổng nhân sự", f"{len(prepared_df):,}")

# Bảng Sinh nhật
st.subheader(f"🎂 Danh sách sinh nhật tháng {selected_month:02d}")
if birthday_df.empty:
    st.info("Không có nhân viên sinh nhật.")
else:
    bday_disp = birthday_df.copy()
    bday_disp["Ngày tháng năm sinh"] = bday_disp["Ngày tháng năm sinh"].dt.strftime("%d/%m/%Y")
    st.dataframe(bday_disp, use_container_width=True, hide_index=True)

st.download_button("⬇️ Download danh sách sinh nhật", data=dataframe_to_excel(birthday_df, "Sinh nhật").getvalue(), file_name=f"danh_sach_sinh_nhat_{selected_month:02d}_{selected_year}.xlsx")

# Bảng Hợp đồng
st.subheader(f"📄 Danh sách hợp đồng hết hạn tháng {selected_month:02d}/{selected_year}")
if contract_df.empty:
    st.info("Không có hợp đồng hết hạn.")
else:
    cntr_disp = contract_df.copy()
    cntr_disp["Ngày tháng năm sinh"] = cntr_disp["Ngày tháng năm sinh"].dt.strftime("%d/%m/%Y")
    cntr_disp["Ngày hết hạn hợp đồng"] = cntr_disp["Ngày hết hạn hợp đồng"].dt.strftime("%d/%m/%Y")
    st.dataframe(cntr_disp, use_container_width=True, hide_index=True)

st.download_button("⬇️ Download danh sách hợp đồng", data=dataframe_to_excel(contract_df, "Hợp đồng").getvalue(), file_name=f"danh_sach_den_han_ky_hop_dong_{selected_month:02d}_{selected_year}.xlsx")

# ============================================================
# TELEGRAM & TỰ ĐỘNG GỬI
# ============================================================
st.divider()
st.header("4️⃣ Telegram Bot")

t_col1, t_col2 = st.columns(2)
with t_col1:
    st.write("### 📱 Trạng thái Bot")
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        ok, bot_name = telegram_test()
        if ok: 
            st.success(f"🟢 Bot đang hoạt động: @{bot_name}")
        else: 
            st.error(f"🔴 Bot lỗi: {bot_name}")
    else: 
        st.warning("Chưa cấu hình Telegram.")

with t_col2:
    st.write("### 📤 Gửi báo cáo thủ công")
    if st.button("📤 Gửi báo cáo ngay", type="primary", use_container_width=True):
        with st.spinner("Đang gửi báo cáo lên Telegram..."):
            success, message = send_monthly_report(birthday_df, contract_df, selected_month, selected_year)
            if success: 
                st.success("✅ Đã gửi thành công 2 file Excel lên Telegram.")
            else: 
                st.error(f"❌ Gửi thất bại: {message}")

st.divider()
st.header("5️⃣ Tự động gửi báo cáo hàng tháng")
st.info("🤖 Hệ thống sẽ tự động gửi báo cáo khi web được truy cập vào ngày 01 từ 10:00 trở đi. Dữ liệu lấy từ file bạn đã lưu.")

now = datetime.now()
st.write(f"🕐 Thời gian máy chủ hiện tại: **{now.strftime('%d/%m/%Y %H:%M:%S')}**")
last_sent = get_last_sent()
if last_sent: 
    st.write(f"📌 Báo cáo gần nhất đã tự động gửi: **{last_sent}**")
else: 
    st.write("📌 Chưa có báo cáo tự động nào được gửi.")

# Lệnh kích hoạt luồng tự động kiểm tra giờ
automatic_monthly_check(birthday_df, contract_df)
