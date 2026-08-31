import io
import json
import hashlib
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from google import genai

# ============================================================
# CẤU HÌNH
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
# Khi AI lỗi/quota, chỉ thử lại sau khoảng thời gian này.
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
    .main-title { font-size: 32px; font-weight: 700; margin-bottom: 10px; }
    .sub-title { color: #666; margin-bottom: 25px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# AI - PRIMARY
# ============================================================

def get_gemini_client():
    if not GEMINI_API_KEY:
        return None, "NO_API_KEY"
    try:
        return genai.Client(api_key=GEMINI_API_KEY), "OK"
    except Exception as e:
        return None, f"CLIENT_ERROR: {e}"


def _classify_gemini_error(exc):
    """Phân loại lỗi để chỉ coi 429/quota là lỗi quota."""
    msg = str(exc)
    low = msg.lower()

    if (
        "429" in low
        or "resource_exhausted" in low
        or "quota" in low
        or "rate limit" in low
        or "too many requests" in low
    ):
        return "QUOTA", msg

    if "401" in low or "403" in low or "unauthorized" in low or "permission" in low:
        return "AUTH_ERROR", msg

    if (
        "timeout" in low
        or "timed out" in low
        or "connection" in low
        or "network" in low
        or "unavailable" in low
        or "503" in low
        or "500" in low
    ):
        return "NETWORK_ERROR", msg

    return "AI_ERROR", msg


def _extract_json_object(text):
    """Lấy JSON kể cả khi model trả về ```json ... ``` hoặc có text thừa."""
    if not text:
        raise ValueError("Gemini trả về nội dung rỗng.")

    text = text.strip()

    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start:end + 1])
        raise


def is_valid_mapping(mapping, df_columns):
    if not isinstance(mapping, dict):
        return False

    valid_columns = {str(c) for c in df_columns}
    for key in REQUIRED_KEYS:
        value = mapping.get(key)
        if not value or str(value) not in valid_columns:
            return False
    return True


def detect_columns_with_gemini(df):
    """
    AI PRIMARY.
    Trả về: (mapping, status, error_message)
    status: AI_OK / QUOTA / AUTH_ERROR / NETWORK_ERROR / AI_ERROR / NO_API_KEY
    """
    client, client_status = get_gemini_client()
    if client is None:
        if client_status == "NO_API_KEY":
            return None, "NO_API_KEY", "Chưa cấu hình GEMINI_API_KEY."
        return None, "AI_ERROR", client_status

    sample_data = []
    sample_df = df.head(8).fillna("")
    for _, row in sample_df.iterrows():
        sample_data.append({str(col): str(row[col])[:80] for col in df.columns})

    columns_text = json.dumps([str(c) for c in df.columns], ensure_ascii=False)
    sample_text = json.dumps(sample_data, ensure_ascii=False, indent=2)

    prompt = f"""
Bạn là AI phân tích dữ liệu nhân sự.

Nhiệm vụ: đọc TÊN CỘT và DỮ LIỆU MẪU để map chính xác vào 5 trường chuẩn.

DANH SÁCH CỘT THỰC TẾ:
{columns_text}

DỮ LIỆU MẪU (tối đa 8 dòng):
{sample_text}

CÁC TRƯỜNG CẦN MAP:
- ma_nv: Mã nhân viên (thường là chuỗi ngắn, ID)
- ho_ten: Họ và tên
- ngay_sinh: Ngày sinh
- bo_phan: Bộ phận/phòng ban/đơn vị
- ngay_het_han_hop_dong: Ngày hết hạn hợp đồng

YÊU CẦU BẮT BUỘC:
1. Chỉ chọn tên cột có thật trong DANH SÁCH CỘT THỰC TẾ.
2. Dựa cả vào tên cột và dữ liệu mẫu, không chỉ dựa vào tên.
3. Nếu không chắc chắn một trường, trả null cho trường đó.
4. Chỉ trả về MỘT object JSON, không markdown, không giải thích.

Định dạng JSON bắt buộc:
{{
  "ma_nv": "tên cột hoặc null",
  "ho_ten": "tên cột hoặc null",
  "ngay_sinh": "tên cột hoặc null",
  "bo_phan": "tên cột hoặc null",
  "ngay_het_han_hop_dong": "tên cột hoặc null"
}}
"""

    try:
        # Không dùng response_schema dạng type=["string", "null"] để tránh
        # lỗi tương thích giữa các phiên bản google-genai.
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config={"temperature": 0.1, "response_mime_type": "application/json"},
        )

        result = _extract_json_object(response.text)

        # Chỉ giữ các key chuẩn.
        result = {key: result.get(key) for key in REQUIRED_KEYS}

        valid_columns = {str(c) for c in df.columns}
        for key in REQUIRED_KEYS:
            value = result.get(key)
            if value is not None and str(value) not in valid_columns:
                result[key] = None

        if not is_valid_mapping(result, df.columns):
            return (
                result,
                "AI_ERROR",
                "Gemini trả mapping nhưng mapping chưa đủ 5 cột bắt buộc.",
            )

        return result, "AI_OK", ""

    except Exception as e:
        status, msg = _classify_gemini_error(e)
        return None, status, msg

# ============================================================
# FALLBACK
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
    """Fallback khi Gemini không khả dụng."""
    columns = list(columns)
    return {
        "ma_nv": find_column_by_keywords(
            columns, ["mã nv", "ma nv", "mã nhân viên", "employee id", "employee_id"]
        ),
        "ho_ten": find_column_by_keywords(
            columns, ["họ và tên", "ho va ten", "họ tên", "tên nhân viên", "full name", "fullname"]
        ),
        "ngay_sinh": find_column_by_keywords(
            columns, ["ngày sinh", "ngay sinh", "ns", "dob", "birthday"]
        ),
        "bo_phan": find_column_by_keywords(
            columns, ["bộ phận", "bo phan", "phòng ban", "phong ban", "đơn vị", "don vi", "department"]
        ),
        "ngay_het_han_hop_dong": find_column_by_keywords(
            columns,
            [
                "ngày hết hạn hợp đồng",
                "ngay het han hop dong",
                "ngày hết hạn hđ",
                "ngay het han hd",
                "contract end",
                "contract_end",
            ],
        ),
    }

# ============================================================
# CACHE / AI RETRY
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


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def parse_iso(value):
    try:
        return datetime.fromisoformat(value) if value else None
    except Exception:
        return None


def should_retry_ai(cache, file_hash):
    if cache.get("hash") != file_hash:
        return True

    # Nếu đã có AI thành công cho đúng file -> không gọi lại.
    if cache.get("method") == "Gemini AI" and is_valid_mapping(cache.get("mapping", {}), cache.get("columns", [])):
        return False

    # Fallback/AI lỗi: chỉ thử lại khi đến thời gian.
    retry_at = parse_iso(cache.get("ai_retry_at"))
    if retry_at is None:
        return True
    return datetime.now() >= retry_at


def set_next_ai_retry(cache):
    cache["ai_retry_at"] = (
        datetime.now() + timedelta(minutes=AI_RETRY_MINUTES)
    ).isoformat(timespec="seconds")


def process_mapping(df_original, file_hash):
    """
    Luồng chính:
    - File mới: AI ngay.
    - AI OK: lưu AI mapping và không gọi lại cho cùng file.
    - AI lỗi/quota: dùng fallback tạm thời + đặt lịch thử AI lại.
    - Khi tới thời gian retry: thử AI lại.
    - AI OK sau retry: chuyển hẳn về AI.
    """
    cache = load_mapping_cache()
    columns = [str(c) for c in df_original.columns]

    # Cache của file khác -> bắt buộc AI trước.
    if cache.get("hash") != file_hash:
        cache = {
            "hash": file_hash,
            "columns": columns,
            "mapping": {},
            "method": "AI_PENDING",
        }
        save_mapping_cache(cache)

        with st.spinner("🤖 Gemini AI đang đọc hiểu dữ liệu..."):
            mapped, status, error = detect_columns_with_gemini(df_original)

        if status == "AI_OK":
            cache.update(
                {
                    "hash": file_hash,
                    "columns": columns,
                    "mapping": mapped,
                    "method": "Gemini AI",
                    "ai_status": "AI_OK",
                    "last_ai_try": now_iso(),
                    "last_ai_error": "",
                    "ai_retry_at": None,
                }
            )
            save_mapping_cache(cache)
            st.success("🤖 Đã phân tích file bằng Gemini AI.")
            return mapped, "Gemini AI", ""

        # AI thất bại lần đầu -> fallback.
        fallback = detect_columns_fallback(df_original.columns)
        cache.update(
            {
                "hash": file_hash,
                "columns": columns,
                "mapping": fallback,
                "method": "Fallback",
                "ai_status": status,
                "last_ai_try": now_iso(),
                "last_ai_error": error,
            }
        )
        set_next_ai_retry(cache)
        save_mapping_cache(cache)
        return fallback, "Fallback", f"Gemini: {status} - {error}"

    # Cùng file, AI đã thành công -> dùng cache AI.
    if cache.get("method") == "Gemini AI" and is_valid_mapping(cache.get("mapping", {}), columns):
        return cache["mapping"], "Gemini AI", ""

    # Đang fallback. Chưa tới thời gian retry -> dùng fallback, KHÔNG gọi API.
    if not should_retry_ai(cache, file_hash):
        return cache.get("mapping", {}), "Fallback", cache.get("last_ai_error", "")

    # Đã tới thời gian retry -> AI lại.
    with st.spinner("🔄 Đang thử kết nối Gemini AI lại..."):
        mapped, status, error = detect_columns_with_gemini(df_original)

    if status == "AI_OK":
        cache.update(
            {
                "hash": file_hash,
                "columns": columns,
                "mapping": mapped,
                "method": "Gemini AI",
                "ai_status": "AI_OK",
                "last_ai_try": now_iso(),
                "last_ai_error": "",
                "ai_retry_at": None,
            }
        )
        save_mapping_cache(cache)
        st.success("🟢 Gemini AI đã hoạt động trở lại. Hệ thống chuyển về AI PRIMARY.")
        return mapped, "Gemini AI", ""

    fallback = cache.get("mapping") or detect_columns_fallback(df_original.columns)
    cache.update(
        {
            "hash": file_hash,
            "columns": columns,
            "mapping": fallback,
            "method": "Fallback",
            "ai_status": status,
            "last_ai_try": now_iso(),
            "last_ai_error": error,
        }
    )
    set_next_ai_retry(cache)
    save_mapping_cache(cache)
    return fallback, "Fallback", f"Gemini: {status} - {error}"

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
        df[mapping["ngay_sinh"]], errors="coerce", dayfirst=True
    )
    result["Bộ Phận"] = df[mapping["bo_phan"]]
    result["Ngày hết hạn hợp đồng"] = pd.to_datetime(
        df[mapping["ngay_het_han_hop_dong"]], errors="coerce", dayfirst=True
    )
    return result, []


def get_birthday_report(df, month):
    res = df[df["Ngày tháng năm sinh"].dt.month == month].copy()
    return res[["Mã NV", "Họ và tên", "Ngày tháng năm sinh", "Bộ Phận"]].sort_values(
        "Ngày tháng năm sinh"
    )


def get_contract_report(df, month, year):
    res = df[
        (df["Ngày hết hạn hợp đồng"].dt.month == month)
        & (df["Ngày hết hạn hợp đồng"].dt.year == year)
    ].copy()
    return res[
        ["Mã NV", "Họ và tên", "Ngày tháng năm sinh", "Bộ Phận", "Ngày hết hạn hợp đồng"]
    ].sort_values("Ngày hết hạn hợp đồng")

# ============================================================
# EXCEL & TELEGRAM
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
                (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                default=0,
            )
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                max(max_length + 2, 12), 40
            )

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
    bday_file = save_excel_file(birthday_df, "sinh_nhat.xlsx", "Sinh nhật")
    cntr_file = save_excel_file(contract_df, "hop_dong.xlsx", "Hợp đồng")

    msg = (
        f"📊 BÁO CÁO NHÂN SỰ {month:02d}/{year}\n"
        f"🎂 Sinh nhật: {len(birthday_df)}\n"
        f"📄 Hợp đồng đến hạn: {len(contract_df)}\n"
        f"⏰ {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    ok, err = telegram_send_message(msg)
    if not ok:
        return False, err

    ok1, err1 = telegram_send_file(bday_file, f"🎂 Sinh nhật tháng {month:02d}")
    if not ok1:
        return False, err1

    ok2, err2 = telegram_send_file(cntr_file, f"📄 Hợp đồng tháng {month:02d}")
    if not ok2:
        return False, err2

    return True, "Thành công"


def automatic_monthly_check(birthday_df, contract_df):
    now = datetime.now()
    if now.day != 1 or now.hour < 10:
        return

    month, year = now.month, now.year
    last_sent = LAST_SENT_FILE.read_text(encoding="utf-8").strip() if LAST_SENT_FILE.exists() else ""

    if last_sent == f"{year}-{month:02d}":
        return

    ok, _ = send_monthly_report(birthday_df, contract_df, month, year)
    if ok:
        LAST_SENT_FILE.write_text(f"{year}-{month:02d}", encoding="utf-8")
        st.toast("🤖 Đã tự động gửi báo cáo Telegram thành công!")

# ============================================================
# UI
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
        st.success(f"🟢 AI PRIMARY: {GEMINI_MODEL}")
        st.caption(f"Tự thử lại sau lỗi: {AI_RETRY_MINUTES} phút")
    else:
        st.warning("🔴 Chưa cấu hình GEMINI_API_KEY")

    st.subheader("📱 Telegram")
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        st.success("🟢 Telegram sẵn sàng")
    else:
        st.warning("🔴 Chưa cấu hình Telegram")

# ============================================================
# FILE UPLOAD - QUAN TRỌNG: KHÔNG XÓA CACHE MỖI STREAMLIT RERUN
# ============================================================

uploaded_file = st.file_uploader(
    "Cập nhật file Excel danh sách nhân sự",
    type=["xlsx", "xls"],
)

if uploaded_file:
    uploaded_bytes = uploaded_file.getvalue()
    uploaded_hash = hashlib.sha256(uploaded_bytes).hexdigest()

    current_hash = get_file_hash(CURRENT_DB_FILE) if CURRENT_DB_FILE.exists() else None

    if current_hash != uploaded_hash:
        with open(CURRENT_DB_FILE, "wb") as f:
            f.write(uploaded_bytes)

        # Chỉ xóa mapping khi THỰC SỰ có file mới.
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
# AI PRIMARY -> FALLBACK -> TỰ RETRY AI
# ============================================================

file_hash = get_file_hash(CURRENT_DB_FILE)
mapping, method_used, ai_error = process_mapping(df_original, file_hash)

# Hiển thị trạng thái nhưng không spam lỗi kỹ thuật.
if method_used == "Fallback":
    cache_now = load_mapping_cache()
    retry_at = parse_iso(cache_now.get("ai_retry_at"))
    retry_text = retry_at.strftime("%d/%m/%Y %H:%M:%S") if retry_at else "chưa xác định"

    st.warning(
        f"⚠️ Gemini tạm thời không khả dụng ({cache_now.get('ai_status', 'UNKNOWN')}). "
        f"Đang dùng Fallback tạm thời. AI sẽ được thử lại lúc {retry_text}."
    )

    # Chi tiết lỗi để debug.
    if ai_error:
        with st.expander("🔎 Chi tiết lỗi Gemini"):
            st.code(ai_error)

prepared_df, missing = prepare_dataframe(df_original, mapping)

if missing:
    st.error(
        "❌ File thiếu các cột quan trọng hoặc Fallback không nhận diện được: "
        + ", ".join(missing)
    )
    st.stop()

# ============================================================
# DASHBOARD
# ============================================================

birthday_df = get_birthday_report(prepared_df, selected_month)
contract_df = get_contract_report(prepared_df, selected_month, selected_year)

st.header(f"Báo cáo tháng {selected_month:02d}/{selected_year}")
c1, c2, c3 = st.columns(3)
c1.metric("🎂 Sinh nhật", f"{len(birthday_df)} người")
c2.metric("📄 Hợp đồng hết hạn", f"{len(contract_df)} người")
c3.metric("👥 Tổng nhân sự", f"{len(prepared_df):,} người")

st.divider()

st.subheader(f"🎂 Danh sách sinh nhật tháng {selected_month:02d}")
if birthday_df.empty:
    st.info("Tháng này không có ai sinh nhật.")
else:
    display_bday = birthday_df.copy()
    display_bday["Ngày tháng năm sinh"] = display_bday["Ngày tháng năm sinh"].dt.strftime("%d/%m/%Y")
    st.dataframe(display_bday, use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ Download Excel Sinh Nhật",
    data=dataframe_to_excel(birthday_df, "Sinh nhật").getvalue(),
    file_name=f"Sinh_nhat_T{selected_month:02d}.xlsx",
)

st.write("---")

st.subheader(f"📄 Hợp đồng đến hạn tháng {selected_month:02d}/{selected_year}")
if contract_df.empty:
    st.info("Tháng này không có hợp đồng nào hết hạn.")
else:
    display_cntr = contract_df.copy()
    display_cntr["Ngày tháng năm sinh"] = display_cntr["Ngày tháng năm sinh"].dt.strftime("%d/%m/%Y")
    display_cntr["Ngày hết hạn hợp đồng"] = display_cntr["Ngày hết hạn hợp đồng"].dt.strftime("%d/%m/%Y")
    st.dataframe(display_cntr, use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ Download Excel Hợp Đồng",
    data=dataframe_to_excel(contract_df, "Hợp đồng").getvalue(),
    file_name=f"Hop_dong_T{selected_month:02d}_{selected_year}.xlsx",
)

# ============================================================
# TELEGRAM
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

with st.expander("ℹ️ Thông tin trạng thái ngầm"):
    cache_info = load_mapping_cache()
    st.write(f"Đã map dữ liệu bằng: **{method_used}**")
    st.write(f"Trạng thái AI: **{cache_info.get('ai_status', 'Không rõ')}**")
    st.write(f"Lần thử AI gần nhất: **{cache_info.get('last_ai_try', 'Chưa có')}**")
    st.write(f"AI thử lại lúc: **{cache_info.get('ai_retry_at', 'Không có')}**")
    if cache_info.get("last_ai_error"):
        st.write(f"Lỗi AI gần nhất: `{cache_info.get('last_ai_error')}`")

    last = LAST_SENT_FILE.read_text(encoding="utf-8").strip() if LAST_SENT_FILE.exists() else "Chưa từng gửi"
    st.write(f"Đã tự động gửi báo cáo lần cuối vào tháng: **{last}**")

# Kích hoạt check tự động hàng tháng (yêu cầu Web được mở vào ngày mùng 1)
automatic_monthly_check(birthday_df, contract_df)
