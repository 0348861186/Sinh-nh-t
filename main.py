import io
import json
import hashlib
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from google import genai
from google.genai import types

# ============================================================
# CẤU HÌNH & KHỞI TẠO THƯ MỤC
# ============================================================

GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()
TELEGRAM_BOT_TOKEN = st.secrets.get("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = str(st.secrets.get("TELEGRAM_CHAT_ID", "")).strip()

REPORT_DIR = Path("reports")
REPORT_DIR.mkdir(exist_ok=True)
LAST_SENT_FILE = REPORT_DIR / "last_sent.txt"

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
CURRENT_DB_FILE = DATA_DIR / "current_hr_data.xlsx"
MAPPING_CACHE_FILE = DATA_DIR / "mapping_cache.json"

GEMINI_MODEL = "gemini-2.5-flash"

# ============================================================
# STREAMLIT CONFIG & CSS
# ============================================================

st.set_page_config(page_title="Quản lý sinh nhật & hợp đồng", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .main-title { font-size: 32px; font-weight: 700; margin-bottom: 10px; }
    .sub-title { color: #666; margin-bottom: 25px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# TẦNG AI (BỘ NÃO CỦA CODE A) - ĐỌC CẢ TIÊU ĐỀ LẪN DỮ LIỆU
# ============================================================

def get_gemini_client():
    if not GEMINI_API_KEY:
        return None
    try:
        return genai.Client(api_key=GEMINI_API_KEY)
    except Exception as e:
        print(f"Lỗi khởi tạo Gemini Client: {e}")
        return None

def detect_columns_with_gemini(df):
    """
    Dùng Gemini để đọc tiêu đề và 8 dòng dữ liệu mẫu.
    Trả về JSON mapping chuẩn xác 100%. Ưu tiên chạy trước tiên.
    """
    client = get_gemini_client()
    if client is None:
        return None

    # Trích xuất 8 dòng dữ liệu mẫu để AI "nhìn thấu" dữ liệu
    sample_data = []
    sample_df = df.head(8).fillna("")
    for _, row in sample_df.iterrows():
        sample_data.append({str(col): str(row[col])[:50] for col in df.columns})

    prompt = f"""
    Bạn là AI phân tích dữ liệu nhân sự.
    Nhiệm vụ: Phân tích danh sách cột và DỮ LIỆU MẪU bên dưới để map vào 5 trường chuẩn.
    
    DỮ LIỆU MẪU (8 dòng đầu):
    {json.dumps(sample_data, ensure_ascii=False, indent=2)}

    CÁC TRƯỜNG CẦN MAP:
    - ma_nv: Mã nhân viên (thường là chuỗi ngắn, ID)
    - ho_ten: Họ và tên (chứa chữ cái)
    - ngay_sinh: Ngày sinh (năm thường < 2010)
    - bo_phan: Bộ phận, phòng ban
    - ngay_het_han_hop_dong: Ngày hết hạn HĐ (thường ở tương lai hoặc năm gần đây)
    
    Tuyệt đối chỉ chọn tên cột CÓ THẬT trong dữ liệu. Nếu không chắc chắn, trả về null.
    """

    # Schema chuẩn cho cấu hình JSON output
    schema = {
        "type": "object",
        "properties": {
            "ma_nv": {"type": ["string", "null"]},
            "ho_ten": {"type": ["string", "null"]},
            "ngay_sinh": {"type": ["string", "null"]},
            "bo_phan": {"type": ["string", "null"]},
            "ngay_het_han_hop_dong": {"type": ["string", "null"]},
        },
        "required": ["ma_nv", "ho_ten", "ngay_sinh", "bo_phan", "ngay_het_han_hop_dong"]
    }

    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json",
                response_schema=schema
            )
        )
        
        if not response or not response.text:
            return None
            
        result = json.loads(response.text.strip())
        
        # Bảo vệ: Đảm bảo cột AI chọn thực sự tồn tại trong file Excel
        valid_columns = set(df.columns)
        for key in result:
            if result[key] not in valid_columns:
                result[key] = None
                
        # Kiểm tra xem AI có trả về ít nhất một trường hợp lệ không
        if not any(result.values()):
            return None
            
        return result
    except Exception as e:
        print(f"Lỗi khi gọi Gemini API: {e}")
        return None

# ============================================================
# TẦNG FALLBACK (PHAO CỨU SINH CỦA CODE B)
# ============================================================

def normalize_text(text):
    return str(text).strip().lower().replace("_", " ").replace("-", " ")

def find_column_by_keywords(columns, keywords):
    normalized = {col: normalize_text(col) for col in columns}
    for col, norm in normalized.items():
        if norm in keywords: return col
    for col, norm in normalized.items():
        for keyword in keywords:
            if keyword in norm: return col
    return None

def detect_columns_fallback(columns):
    """Quy tắc dự phòng nếu AI sập, hết Quota, hoặc không có mạng."""
    mapping = {}
    mapping["ma_nv"] = find_column_by_keywords(columns, ["mã nv", "ma nv", "mã nhân viên", "employee id"])
    mapping["ho_ten"] = find_column_by_keywords(columns, ["họ và tên", "ho va ten", "họ tên", "tên nhân viên", "full name"])
    mapping["ngay_sinh"] = find_column_by_keywords(columns, ["ngày sinh", "ngay sinh", "ns", "dob", "birthday"])
    mapping["bo_phan"] = find_column_by_keywords(columns, ["bộ phận", "bo phan", "phòng ban", "đơn vị", "department"])
    mapping["ngay_het_han_hop_dong"] = find_column_by_keywords(columns, ["ngày hết hạn hợp đồng", "ngay het han hop dong", "ngày hết hạn hđ", "contract end"])
    return mapping

# ============================================================
# HỆ THỐNG CACHE & CHUẨN HÓA DỮ LIỆU
# ============================================================

def get_file_hash(file_path):
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()

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

def get_birthday_report(df, month):
    res = df[df["Ngày tháng năm sinh"].dt.month == month].copy()
    return res[["Mã NV", "Họ và tên", "Ngày tháng năm sinh", "Bộ Phận"]].sort_values("Ngày tháng năm sinh")

def get_contract_report(df, month, year):
    res = df[(df["Ngày hết hạn hợp đồng"].dt.month == month) & (df["Ngày hết hạn hợp đồng"].dt.year == year)].copy()
    return res[["Mã NV", "Họ và tên", "Ngày tháng năm sinh", "Bộ Phận", "Ngày hết hạn hợp đồng"]].sort_values("Ngày hết hạn hợp đồng")

# ============================================================
# XUẤT EXCEL & TELEGRAM
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
            max_length = max((len(str(cell.value)) for cell in column_cells if cell.value), default=0)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 40)
    output.seek(0)
    return output

def save_excel_file(df, filename, sheet_name):
    path = REPORT_DIR / filename
    with open(path, "wb") as f:
        f.write(dataframe_to_excel(df, sheet_name).getvalue())
    return path

def telegram_send_message(message):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID: return False, "Thiếu Token/Chat ID"
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        resp = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "text": message}, timeout=30)
        return (True, "OK") if resp.ok else (False, resp.text)
    except Exception as e: return False, str(e)

def telegram_send_file(file_path, caption):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID: return False, "Thiếu Token/Chat ID"
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    try:
        with open(file_path, "rb") as f:
            resp = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption}, files={"document": f}, timeout=60)
        return (True, "OK") if resp.ok else (False, resp.text)
    except Exception as e: return False, str(e)

def send_monthly_report(birthday_df, contract_df, month, year):
    bday_file = save_excel_file(birthday_df, "sinh_nhat.xlsx", "Sinh nhật")
    cntr_file = save_excel_file(contract_df, "hop_dong.xlsx", "Hợp đồng")

    msg = f"📊 BÁO CÁO NHÂN SỰ {month:02d}/{year}\n🎂 Sinh nhật: {len(birthday_df)}\n📄 Hợp đồng đến hạn: {len(contract_df)}\n⏰ {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ok, err = telegram_send_message(msg)
    if not ok: return False, err

    telegram_send_file(bday_file, f"🎂 Sinh nhật tháng {month:02d}")
    telegram_send_file(cntr_file, f"📄 Hợp đồng tháng {month:02d}")
    return True, "Thành công"

def automatic_monthly_check(birthday_df, contract_df):
    now = datetime.now()
    if now.day != 1 or now.hour < 10: return
    month, year = now.month, now.year
    last_sent = LAST_SENT_FILE.read_text().strip() if LAST_SENT_FILE.exists() else ""
    if last_sent == f"{year}-{month:02d}": return

    if send_monthly_report(birthday_df, contract_df, month, year)[0]:
        LAST_SENT_FILE.write_text(f"{year}-{month:02d}")
        st.toast("🤖 Đã tự động gửi báo cáo Telegram thành công!")

# ============================================================
# UI CHÍNH - DASHBOARD & SIDEBAR
# ============================================================

st.markdown('<div class="main-title">📊 QUẢN LÝ SINH NHẬT & HỢP ĐỒNG (AI PRO)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Nhận diện thông minh dữ liệu, Dashboard tức thì và Tự động báo cáo</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Cấu hình")
    cur_date = date.today()
    selected_year = st.selectbox("Năm", list(range(cur_date.year - 2, cur_date.year + 6)), index=2)
    selected_month = st.selectbox("Tháng", list(range(1, 13)), index=cur_date.month - 1, format_func=lambda x: f"Tháng {x:02d}")
    
    st.divider()
    st.subheader("🤖 Gemini AI")
    if GEMINI_API_KEY: st.success("🟢 API AI đang kích hoạt")
    else: st.warning("🔴 Đang dùng thuật toán dự phòng")
    
    st.subheader("📱 Telegram")
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID: st.success("🟢 Telegram sẵn sàng")
    else: st.warning("🔴 Chưa cấu hình Telegram")

# ============================================================
# LUỒNG XỬ LÝ CHÍNH
# ============================================================

uploaded_file = st.file_uploader("Cập nhật file Excel danh sách nhân sự", type=["xlsx", "xls"])
if uploaded_file:
    with open(CURRENT_DB_FILE, "wb") as f: f.write(uploaded_file.getbuffer())
    if MAPPING_CACHE_FILE.exists(): MAPPING_CACHE_FILE.unlink() # Xóa cache cũ
    st.success("✅ Đã cập nhật cơ sở dữ liệu nhân sự mới nhất!")

if not CURRENT_DB_FILE.exists():
    st.info("📁 Hệ thống trống. Hãy upload file Excel lần đầu tiên.")
    st.stop()

# Đọc Data
try:
    df_original = pd.read_excel(CURRENT_DB_FILE).dropna(how="all")
    df_original.columns = [str(col).strip() for col in df_original.columns]
except Exception as e:
    st.error(f"Lỗi đọc file: {e}")
    st.stop()

# Kiểm tra Hash & Mapping
file_hash = get_file_hash(CURRENT_DB_FILE)
mapping_info = {}

if MAPPING_CACHE_FILE.exists():
    try: mapping_info = json.loads(MAPPING_CACHE_FILE.read_text())
    except: pass

# NẾU FILE MỚI -> Kích hoạt AI phân tích 1 lần duy nhất
if mapping_info.get("hash") != file_hash:
    with st.spinner("🤖 Trí tuệ nhân tạo đang đọc hiểu dữ liệu của bạn... (Chỉ tốn vài giây)"):
        # Thử AI trước
        mapped = detect_columns_with_gemini(df_original)
        method_used = "Gemini AI"
        
        # Nếu AI lỗi (Hết hạn, mất mạng, không trả về giá trị) -> Fallback
        if not mapped or not any(mapped.values()):
            mapped = detect_columns_fallback(df_original.columns)
            method_used = "Thuật toán dự phòng (Fallback)"
            
        mapping_info = {"hash": file_hash, "mapping": mapped, "method": method_used}
        MAPPING_CACHE_FILE.write_text(json.dumps(mapping_info))
        st.toast(f"Đã phân tích xong bằng {method_used}!")

# Chuẩn hóa DF
mapping = mapping_info.get("mapping", {})
prepared_df, missing = prepare_dataframe(df_original, mapping)

if missing:
    st.error(f"❌ File thiếu các cột quan trọng. Vui lòng kiểm tra lại. Cột thiếu: {', '.join(missing)}")
    st.stop()

# ============================================================
# HIỂN THỊ DASHBOARD & NÚT DOWNLOAD
# ============================================================

birthday_df = get_birthday_report(prepared_df, selected_month)
contract_df = get_contract_report(prepared_df, selected_month, selected_year)

st.header(f"Báo cáo tháng {selected_month:02d}/{selected_year}")
c1, c2, c3 = st.columns(3)
c1.metric("🎂 Sinh nhật", f"{len(birthday_df)} người")
c2.metric("📄 Hợp đồng hết hạn", f"{len(contract_df)} người")
c3.metric("👥 Tổng nhân sự", f"{len(prepared_df):,} người")

st.divider()

# DANH SÁCH SINH NHẬT
st.subheader(f"🎂 Danh sách sinh nhật tháng {selected_month:02d}")
if birthday_df.empty: st.info("Tháng này không có ai sinh nhật.")
else:
    display_bday = birthday_df.copy()
    display_bday["Ngày tháng năm sinh"] = display_bday["Ngày tháng năm sinh"].dt.strftime("%d/%m/%Y")
    st.dataframe(display_bday, use_container_width=True, hide_index=True)
st.download_button("⬇️ Download Excel Sinh Nhật", data=dataframe_to_excel(birthday_df, "Sinh nhật").getvalue(), file_name=f"Sinh_nhat_T{selected_month:02d}.xlsx")

st.write("---")

# DANH SÁCH HỢP ĐỒNG
st.subheader(f"📄 Hợp đồng đến hạn tháng {selected_month:02d}/{selected_year}")
if contract_df.empty: st.info("Tháng này không có hợp đồng nào hết hạn.")
else:
    display_cntr = contract_df.copy()
    display_cntr["Ngày tháng năm sinh"] = display_cntr["Ngày tháng năm sinh"].dt.strftime("%d/%m/%Y")
    display_cntr["Ngày hết hạn hợp đồng"] = display_cntr["Ngày hết hạn hợp đồng"].dt.strftime("%d/%m/%Y")
    st.dataframe(display_cntr, use_container_width=True, hide_index=True)
st.download_button("⬇️ Download Excel Hợp Đồng", data=dataframe_to_excel(contract_df, "Hợp đồng").getvalue(), file_name=f"Hop_dong_T{selected_month:02d}_{selected_year}.xlsx")

# ============================================================
# KHU VỰC TELEGRAM
# ============================================================
st.divider()
st.header("📤 Gửi Telegram")
if st.button("Gửi báo cáo thủ công lên Telegram ngay bây giờ", type="primary"):
    with st.spinner("Đang gửi..."):
        ok, msg = send_monthly_report(birthday_df, contract_df, selected_month, selected_year)
        if ok: st.success("✅ Đã gửi thành công!")
        else: st.error(f"❌ Lỗi: {msg}")

with st.expander("ℹ️ Thông tin trạng thái ngầm"):
    st.write(f"Đã map dữ liệu bằng: **{mapping_info.get('method', 'Không rõ')}**")
    last = LAST_SENT_FILE.read_text().strip() if LAST_SENT_FILE.exists() else "Chưa từng gửi"
    st.write(f"Đã tự động gửi báo cáo lần cuối vào tháng: **{last}**")

# Kích hoạt Check tự động hàng tháng (Yêu cầu Web được mở vào ngày mùng 1)
automatic_monthly_check(birthday_df, contract_df)
