import io
import json
import hashlib
from datetime import datetime, date, timedelta
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

# Sử dụng phiên bản mô hình mới theo yêu cầu
GEMINI_MODEL = "gemini-3.1-pro-preview"

# Khi Gemini đang ở trạng thái quota/error, chỉ thử lại sau khoảng thời gian này.
AI_RETRY_MINUTES = 30

REQUIRED_KEYS = [
    "ma_nv",
    "ho_ten",
    "ngay_sinh",
    "bo_phan",
    "ngay_het_han_hop_dong",
]

# ============================================================
# STREAMLIT CONFIG & CSS
# ============================================================

st.set_page_config(
    page_title="Quản lý sinh nhật & hợp đồng",
    page_icon="📊",
    layout="wide",
)

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
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# TẦNG AI (PRIMARY) - LUÔN ƯU TIÊN GEMINI
# ============================================================

def get_gemini_client():
    """Tạo Gemini client. Không có API key thì trả về None."""
    if not GEMINI_API_KEY:
        return None

    try:
        return genai.Client(api_key=GEMINI_API_KEY)
    except Exception:
        return None


def detect_columns_with_gemini(df):
    """
    Gọi Gemini để đọc tiêu đề + 8 dòng dữ liệu mẫu.
    """
    client = get_gemini_client()

    if client is None:
        return None, "NO_API", "Không có GEMINI_API_KEY hoặc không tạo được Gemini client."

    sample_data = []
    sample_df = df.head(8).fillna("")

    for _, row in sample_df.iterrows():
        sample_data.append(
            {str(col): str(row[col])[:50] for col in df.columns}
        )

    prompt = f"""
Bạn là AI phân tích dữ liệu nhân sự.

Nhiệm vụ: Phân tích danh sách cột và DỮ LIỆU MẪU bên dưới để map vào 5 trường chuẩn.

DỮ LIỆU MẪU (tối đa 8 dòng đầu):
{json.dumps(sample_data, ensure_ascii=False, indent=2)}

CÁC TRƯỜNG CẦN MAP:
- ma_nv: Mã nhân viên (thường là chuỗi ngắn, ID)
- ho_ten: Họ và tên (chứa chữ cái)
- ngay_sinh: Ngày sinh (năm thường < 2010)
- bo_phan: Bộ phận, phòng ban
- ngay_het_han_hop_dong: Ngày hết hạn HĐ (thường ở tương lai hoặc năm gần đây)

QUY TẮC BẮT BUỘC:
1. Chỉ chọn tên cột CÓ THẬT trong dữ liệu.
2. Không tự tạo tên cột.
3. Nếu không chắc chắn, trả về null.
4. Trả đúng JSON theo schema được cung cấp.
"""

    schema = {
        "type": "object",
        "properties": {
            "ma_nv": {"type": ["string", "null"]},
            "ho_ten": {"type": ["string", "null"]},
            "ngay_sinh": {"type": ["string", "null"]},
            "bo_phan": {"type": ["string", "null"]},
            "ngay_het_han_hop_dong": {"type": ["string", "null"]},
        },
        "required": REQUIRED_KEYS,
    }

    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json",
                response_schema=schema,
            ),
        )

        if not response or not getattr(response, "text", None):
            return None, "AI_ERROR", "Gemini không trả về nội dung."

        result = json.loads(response.text.strip())

        if not isinstance(result, dict):
            return None, "AI_ERROR", "Kết quả Gemini không phải JSON object."

        valid_columns = set(df.columns)
        clean_result = {}

        for key in REQUIRED_KEYS:
            value = result.get(key)
            clean_result[key] = value if value in valid_columns else None

        if not all(clean_result.get(key) for key in REQUIRED_KEYS):
            return (
                clean_result,
                "AI_ERROR",
                "Gemini trả về mapping nhưng chưa map đủ 5 trường bắt buộc.",
            )

        return clean_result, "AI_OK", ""

    except Exception as e:
        error_text = str(e)
        lower_error = error_text.lower()

        quota_markers = [
            "quota",
            "resource_exhausted",
            "rate limit",
            "rate_limit",
            "429",
            "too many requests",
            "exceeded",
        ]

        if any(marker in lower_error for marker in quota_markers):
            return None, "QUOTA", error_text

        return None, "AI_ERROR", error_text


# ============================================================
# TẦNG FALLBACK (CHỈ DÙNG KHI AI KHÔNG KHẢ DỤNG)
# ============================================================

def normalize_text(text):
    return (
        str(text)
        .strip()
        .lower()
        .replace("_", " ")
        .replace("-", " ")
    )


def find_column_by_keywords(columns, keywords):
    normalized = {col: normalize_text(col) for col in columns}
    normalized_keywords = [normalize_text(k) for k in keywords]

    for col, norm in normalized.items():
        if norm in normalized_keywords:
            return col

    for col, norm in normalized.items():
        for keyword in normalized_keywords:
            if keyword in norm:
                return col

    return None


def detect_columns_fallback(columns):
    """Thuật toán dự phòng khi Gemini không khả dụng."""
    mapping = {}

    mapping["ma_nv"] = find_column_by_keywords(
        columns,
        ["mã nv", "ma nv", "mã nhân viên", "ma nhan vien", "employee id"],
    )

    mapping["ho_ten"] = find_column_by_keywords(
        columns,
        ["họ và tên", "ho va ten", "họ tên", "ho ten", "tên nhân viên", "ten nhan vien", "full name"],
    )

    mapping["ngay_sinh"] = find_column_by_keywords(
        columns,
        ["ngày sinh", "ngay sinh", "ns", "dob", "birthday"],
    )

    mapping["bo_phan"] = find_column_by_keywords(
        columns,
        ["bộ phận", "bo phan", "phòng ban", "phong ban", "đơn vị", "don vi", "department"],
    )

    mapping["ngay_het_han_hop_dong"] = find_column_by_keywords(
        columns,
        [
            "ngày hết hạn hợp đồng",
            "ngay het han hop dong",
            "ngày hết hạn hđ",
            "ngay het han hd",
            "contract end",
        ],
    )

    return mapping


# ============================================================
# CACHE / TRẠNG THÁI AI
# ============================================================

def get_file_hash(file_path):
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load_mapping_cache():
    if not MAPPING_CACHE_FILE.exists():
        return {}

    try:
        data = json.loads(MAPPING_CACHE_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_mapping_cache(data):
    MAPPING_CACHE_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def parse_datetime(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def should_retry_ai(mapping_info, file_hash):
    if mapping_info.get("hash") != file_hash:
        return True

    if mapping_info.get("method") == "Gemini AI":
        return False

    next_retry = parse_datetime(mapping_info.get("next_ai_retry"))
    if next_retry is None:
        return True

    return datetime.now() >= next_retry


def process_mapping(df_original, file_hash, mapping_info):
    if (
        mapping_info.get("hash") == file_hash
        and mapping_info.get("method") == "Gemini AI"
        and mapping_info.get("mapping")
    ):
        return mapping_info, "CACHE_AI"

    if not should_retry_ai(mapping_info, file_hash):
        return mapping_info, "CACHE_FALLBACK"

    with st.spinner("🤖 Gemini AI đang phân tích dữ liệu... (AI là chế độ ưu tiên)"):
        mapped, ai_status, ai_error = detect_columns_with_gemini(df_original)

    if ai_status == "AI_OK":
        mapping_info = {
            "hash": file_hash,
            "mapping": mapped,
            "method": "Gemini AI",
            "ai_status": "AI_OK",
            "last_ai_try": datetime.now().isoformat(timespec="seconds"),
            "next_ai_retry": None,
            "last_ai_error": "",
        }
        save_mapping_cache(mapping_info)
        return mapping_info, "AI_SUCCESS"

    fallback_mapping = detect_columns_fallback(df_original.columns)
    next_retry = datetime.now() + timedelta(minutes=AI_RETRY_MINUTES)

    mapping_info = {
        "hash": file_hash,
        "mapping": fallback_mapping,
        "method": "Thuật toán dự phòng (Fallback)",
        "ai_status": ai_status,
        "last_ai_try": datetime.now().isoformat(timespec="seconds"),
        "next_ai_retry": next_retry.isoformat(timespec="seconds"),
        "last_ai_error": ai_error[:1000] if ai_error else "",
    }

    save_mapping_cache(mapping_info)

    return mapping_info, f"FALLBACK_{ai_status}"


# ============================================================
# CHUẨN HÓA DỮ LIỆU
# ============================================================

def prepare_dataframe(df, mapping):
    missing = [key for key in REQUIRED_KEYS if not mapping.get(key)]

    if missing:
        return None, missing

    result = pd.DataFrame()

    result["Mã NV"] = df[mapping["ma_nv"]]
    result["Họ và tên"] = df[mapping["ho_ten"]]
    result["Ngày tháng năm sinh"] = pd.to_datetime(
        df[mapping["ngay_sinh"]],
        errors="coerce",
        dayfirst=True,
    )
    result["Bộ Phận"] = df[mapping["bo_phan"]]
    result["Ngày hết hạn hợp đồng"] = pd.to_datetime(
        df[mapping["ngay_het_han_hop_dong"]],
        errors="coerce",
        dayfirst=True,
    )

    return result, []


def get_birthday_report(df, month):
    res = df[df["Ngày tháng năm sinh"].dt.month == month].copy()
    return res[
        ["Mã NV", "Họ và tên", "Ngày tháng năm sinh", "Bộ Phận"]
    ].sort_values("Ngày tháng năm sinh")


def get_contract_report(df, month, year):
    res = df[
        (df["Ngày hết hạn hợp đồng"].dt.month == month)
        & (df["Ngày hết hạn hợp đồng"].dt.year == year)
    ].copy()

    return res[
        [
            "Mã NV",
            "Họ và tên",
            "Ngày tháng năm sinh",
            "Bộ Phận",
            "Ngày hết hạn hợp đồng",
        ]
    ].sort_values("Ngày hết hạn hợp đồng")


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
            max_length = max(
                (
                    len(str(cell.value))
                    for cell in column_cells
                    if cell.value is not None
                ),
                default=0,
            )
            worksheet.column_dimensions[
                column_cells[0].column_letter
            ].width = min(max(max_length + 2, 12), 40)

    output.seek(0)
    return output


def save_excel_file(df, filename, sheet_name):
    path = REPORT_DIR / filename

    with open(path, "wb") as f:
        f.write(dataframe_to_excel(df, sheet_name).getvalue())

    return path


def telegram_send_message(message):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False, "Thiếu Token/Chat ID"

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"

    try:
        resp = requests.post(
            url,
            data={"chat_id": TELEGRAM_CHAT_ID, "text": message},
            timeout=30,
        )
        return (True, "OK") if resp.ok else (False, resp.text)
    except Exception as e:
        return False, str(e)


def telegram_send_file(file_path, caption):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False, "Thiếu Token/Chat ID"

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"

    try:
        with open(file_path, "rb") as f:
            resp = requests.post(
                url,
                data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption},
                files={"document": f},
                timeout=60,
            )
        return (True, "OK") if resp.ok else (False, resp.text)
    except Exception as e:
        return False, str(e)


def send_monthly_report(birthday_df, contract_df, month, year):
    bday_file = save_excel_file(
        birthday_df,
        "sinh_nhat.xlsx",
        "Sinh nhật",
    )
    cntr_file = save_excel_file(
        contract_df,
        "hop_dong.xlsx",
        "Hợp đồng",
    )

    msg = (
        f"📊 BÁO CÁO NHÂN SỰ {month:02d}/{year}\n"
        f"🎂 Sinh nhật: {len(birthday_df)}\n"
        f"📄 Hợp đồng đến hạn: {len(contract_df)}\n"
        f"⏰ {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    ok, err = telegram_send_message(msg)

    if not ok:
        return False, err

    file_ok_1, file_err_1 = telegram_send_file(
        bday_file,
        f"🎂 Sinh nhật tháng {month:02d}",
    )
    file_ok_2, file_err_2 = telegram_send_file(
        cntr_file,
        f"📄 Hợp đồng tháng {month:02d}",
    )

    if not file_ok_1:
        return False, file_err_1
    if not file_ok_2:
        return False, file_err_2

    return True, "Thành công"


def automatic_monthly_check(birthday_df, contract_df):
    now = datetime.now()

    if now.day != 1 or now.hour < 10:
        return

    month, year = now.month, now.year

    last_sent = (
        LAST_SENT_FILE.read_text(encoding="utf-8").strip()
        if LAST_SENT_FILE.exists()
        else ""
    )

    if last_sent == f"{year}-{month:02d}":
        return

    if send_monthly_report(birthday_df, contract_df, month, year)[0]:
        LAST_SENT_FILE.write_text(
            f"{year}-{month:02d}",
            encoding="utf-8",
        )
        st.toast("🤖 Đã tự động gửi báo cáo Telegram thành công!")


# ============================================================
# UI CHÍNH - DASHBOARD & SIDEBAR
# ============================================================

st.markdown(
    '<div class="main-title">📊 QUẢN LÝ SINH NHẬT & HỢP ĐỒNG (AI PRO)</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="sub-title">Nhận diện thông minh dữ liệu, Dashboard tức thì và Tự động báo cáo</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("⚙️ Cấu hình")

    cur_date = date.today()

    selected_year = st.selectbox(
        "Năm",
        list(range(cur_date.year - 2, cur_date.year + 6)),
        index=2,
    )

    selected_month = st.selectbox(
        "Tháng",
        list(range(1, 13)),
        index=cur_date.month - 1,
        format_func=lambda x: f"Tháng {x:02d}",
    )

    st.divider()

    st.subheader("🤖 Gemini AI")

    if GEMINI_API_KEY:
        st.success("🟢 AI PRIMARY đang kích hoạt")
    else:
        st.warning("🔴 Chưa có API key — chỉ có thể dùng Fallback")

    st.caption(
        f"Tự động thử lại AI sau lỗi/quota: {AI_RETRY_MINUTES} phút"
    )

    st.subheader("📱 Telegram")

    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        st.success("🟢 Telegram sẵn sàng")
    else:
        st.warning("🔴 Chưa cấu hình Telegram")


# ============================================================
# LUỒNG XỬ LÝ CHÍNH
# ============================================================

uploaded_file = st.file_uploader(
    "Cập nhật file Excel danh sách nhân sự",
    type=["xlsx", "xls"],
)

if uploaded_file:
    with open(CURRENT_DB_FILE, "wb") as f:
        f.write(uploaded_file.getbuffer())

    if MAPPING_CACHE_FILE.exists():
        MAPPING_CACHE_FILE.unlink()

    st.success("✅ Đã cập nhật cơ sở dữ liệu nhân sự mới nhất!")

if not CURRENT_DB_FILE.exists():
    st.info("📁 Hệ thống trống. Hãy upload file Excel lần đầu tiên.")
    st.stop()


# ============================================================
# ĐỌC DATA
# ============================================================

try:
    df_original = pd.read_excel(CURRENT_DB_FILE).dropna(how="all")
    df_original.columns = [str(col).strip() for col in df_original.columns]
except Exception as e:
    st.error(f"Lỗi đọc file: {e}")
    st.stop()


# ============================================================
# AI PRIMARY -> FALLBACK -> TỰ ĐỘNG QUAY LẠI AI
# ============================================================

file_hash = get_file_hash(CURRENT_DB_FILE)
mapping_info = load_mapping_cache()

mapping_info, process_status = process_mapping(
    df_original,
    file_hash,
    mapping_info,
)

method = mapping_info.get("method", "Không rõ")
ai_status = mapping_info.get("ai_status", "")
next_retry = parse_datetime(mapping_info.get("next_ai_retry"))

if process_status == "AI_SUCCESS":
    st.toast("🤖 Gemini AI đã phân tích thành công. AI PRIMARY đang được sử dụng.")

elif process_status.startswith("FALLBACK"):
    if ai_status == "QUOTA":
        st.warning(
            f"⚠️ Gemini đang hết quota/rate limit. Đang dùng Fallback tạm thời. "
            f"Hệ thống sẽ tự thử lại AI sau {AI_RETRY_MINUTES} phút."
        )
    elif ai_status == "NO_API":
        st.warning(
            "⚠️ Chưa có GEMINI_API_KEY. Đang dùng Fallback. "
            "Khi cấu hình API key và chạy lại ứng dụng, AI sẽ được ưu tiên."
        )
    else:
        err_detail = mapping_info.get("last_ai_error", "")
        st.warning(
            f"⚠️ Gemini không khả dụng ({ai_status}: {err_detail}). Đang dùng Fallback tạm thời."
        )

elif process_status == "CACHE_FALLBACK":
    if next_retry:
        remaining = max(
            0,
            int((next_retry - datetime.now()).total_seconds() / 60),
        )
        st.info(
            f"🔄 Đang dùng Fallback tạm thời. Gemini sẽ được tự động thử lại "
            f"sau khoảng {remaining} phút."
        )


mapping = mapping_info.get("mapping", {})
prepared_df, missing = prepare_dataframe(df_original, mapping)

if missing:
    st.error(
        "❌ File thiếu các cột quan trọng. Vui lòng kiểm tra lại. "
        f"Cột thiếu: {', '.join(missing)}"
    )
    st.stop()


# ============================================================
# HIỂN THỊ DASHBOARD & NÚT DOWNLOAD
# ============================================================

birthday_df = get_birthday_report(prepared_df, selected_month)
contract_df = get_contract_report(
    prepared_df,
    selected_month,
    selected_year,
)

st.header(f"Báo cáo tháng {selected_month:02d}/{selected_year}")

c1, c2, c3 = st.columns(3)

c1.metric("🎂 Sinh nhật", f"{len(birthday_df)} người")
c2.metric("📄 Hợp đồng hết hạn", f"{len(contract_df)} người")
c3.metric("👥 Tổng nhân sự", f"{len(prepared_df)::,} người")

st.divider()

# DANH SÁCH SINH NHẬT
st.subheader(f"🎂 Danh sách sinh nhật tháng {selected_month:02d}")

if birthday_df.empty:
    st.info("Tháng này không có ai sinh nhật.")
else:
    display_bday = birthday_df.copy()
    display_bday["Ngày tháng năm sinh"] = display_bday[
        "Ngày tháng năm sinh"
    ].dt.strftime("%d/%m/%Y")
    st.dataframe(
        display_bday,
        use_container_width=True,
        hide_index=True,
    )

st.download_button(
    "⬇️ Download Excel Sinh Nhật",
    data=dataframe_to_excel(birthday_df, "Sinh nhật").getvalue(),
    file_name=f"Sinh_nhat_T{selected_month:02d}.xlsx",
)

st.write("---")

# DANH SÁCH HỢP ĐỒNG
st.subheader(f"📄 Hợp đồng đến hạn tháng {selected_month:02d}/{selected_year}")

if contract_df.empty:
    st.info("Tháng này không có hợp đồng nào hết hạn.")
else:
    display_cntr = contract_df.copy()
    display_cntr["Ngày tháng năm sinh"] = display_cntr[
        "Ngày tháng năm sinh"
    ].dt.strftime("%d/%m/%Y")
    display_cntr["Ngày hết hạn hợp đồng"] = display_cntr[
        "Ngày hết hạn hợp đồng"
    ].dt.strftime("%d/%m/%Y")
    st.dataframe(
        display_cntr,
        use_container_width=True,
        hide_index=True,
    )

st.download_button(
    "⬇️ Download Excel Hợp Đồng",
    data=dataframe_to_excel(contract_df, "Hợp đồng").getvalue(),
    file_name=f"Hop_dong_T{selected_month:02d}_{selected_year}.xlsx",
)


# ============================================================
# KHU VỰC TELEGRAM
# ============================================================

st.divider()
st.header("📤 Gửi Telegram")

if st.button("Gửi báo cáo thủ công lên Telegram ngay bây giờ", type="primary"):
    with st.spinner("Đang gửi..."):
        ok, msg = send_monthly_report(
            birthday_df,
            contract_df,
            selected_month,
            selected_year,
        )

        if ok:
            st.success("✅ Đã gửi thành công!")
        else:
            st.error(f"❌ Lỗi: {msg}")


# ============================================================
# THÔNG TIN TRẠNG THÁI
# ============================================================

with st.expander("ℹ️ Thông tin trạng thái ngầm"):
    st.write(f"Đã map dữ liệu bằng: **{method}**")
    st.write(f"Trạng thái AI gần nhất: **{ai_status or 'Chưa có'}**")

    if mapping_info.get("last_ai_try"):
        st.write(f"Lần thử AI gần nhất: **{mapping_info['last_ai_try']}**")

    if next_retry:
        st.write(
            f"AI dự kiến thử lại: **{mapping_info.get('next_ai_retry')}**"
        )

    if mapping_info.get("last_ai_error"):
        st.caption(
            f"Lỗi AI gần nhất: {mapping_info['last_ai_error']}"
        )

    last = (
        LAST_SENT_FILE.read_text(encoding="utf-8").strip()
        if LAST_SENT_FILE.exists()
        else "Chưa từng gửi"
    )

    st.write(
        f"Đã tự động gửi báo cáo lần cuối vào tháng: **{last}**"
    )


automatic_monthly_check(birthday_df, contract_df)
