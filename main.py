```python
import io
import json
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

# ------------------------------------------------------------
# Đọc API KEY từ Streamlit Secrets
# ------------------------------------------------------------

GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()
TELEGRAM_BOT_TOKEN = st.secrets.get("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = str(
    st.secrets.get("TELEGRAM_CHAT_ID", "")
).strip()


# ------------------------------------------------------------
# Gemini Model
# ------------------------------------------------------------

GEMINI_MODEL = "gemini-2.5-pro"


# ------------------------------------------------------------
# Thư mục báo cáo
# ------------------------------------------------------------

REPORT_DIR = Path("reports")
REPORT_DIR.mkdir(exist_ok=True)

LAST_SENT_FILE = REPORT_DIR / "last_sent.txt"


# ------------------------------------------------------------
# Thư mục dữ liệu
# ------------------------------------------------------------

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
# GEMINI
# ============================================================

def get_gemini_client():
    """
    Khởi tạo Gemini Client.
    """

    if not GEMINI_API_KEY:
        return None

    try:
        return genai.Client(
            api_key=GEMINI_API_KEY
        )

    except Exception as e:
        st.error(
            f"Không thể khởi tạo Gemini Client: {e}"
        )
        return None


# ============================================================
# NHẬN DIỆN CỘT BẰNG GEMINI 2.5 PRO
# ============================================================

def detect_columns_with_gemini(columns):

    client = get_gemini_client()

    if client is None:
        return None

    column_list = "\n".join(
        [
            f"- {str(col)}"
            for col in columns
        ]
    )

    prompt = f"""
Bạn là chuyên gia phân tích dữ liệu nhân sự và Excel.

Hãy phân tích danh sách tên cột Excel dưới đây.

DANH SÁCH CỘT THỰC TẾ:

{column_list}


NHIỆM VỤ:

Xác định cột nào tương ứng với 5 trường dữ liệu sau:


1. ma_nv

Có thể là:

- Mã NV
- Mã nhân viên
- Ma NV
- Ma nhan vien
- Employee ID
- Employee Code
- Staff ID
- Staff Code


2. ho_ten

Có thể là:

- Họ và tên
- Họ tên
- Ho va ten
- Ho ten
- Tên nhân viên
- Ten nhan vien
- Full Name
- Employee Name


3. ngay_sinh

Có thể là:

- Ngày sinh
- Ngày tháng năm sinh
- Ngay sinh
- Date of Birth
- DOB
- Birthday


4. bo_phan

Có thể là:

- Bộ phận
- Bo phan
- Phòng ban
- Phong ban
- Đơn vị
- Don vi
- Department
- Division


5. ngay_het_han_hop_dong

Có thể là:

- Ngày hết hạn hợp đồng
- Ngay het han hop dong
- Ngày hết hạn HĐ
- Ngay het han HD
- Ngày kết thúc hợp đồng
- Ngay ket thuc hop dong
- Contract End Date
- Contract Expiry Date
- Contract Expiration Date


QUY TẮC RẤT QUAN TRỌNG:

1. Chỉ được chọn tên cột thực tế có trong danh sách.
2. Không được tự tạo tên cột.
3. Không được sửa tên cột.
4. Nếu không tìm thấy thì trả về null.
5. Nếu có nhiều cột phù hợp, chọn cột chính xác nhất.
6. Không giải thích.
7. Chỉ trả về JSON.
8. JSON phải có đầy đủ 5 key.


FORMAT:

{{
    "ma_nv": "tên cột hoặc null",
    "ho_ten": "tên cột hoặc null",
    "ngay_sinh": "tên cột hoặc null",
    "bo_phan": "tên cột hoặc null",
    "ngay_het_han_hop_dong": "tên cột hoặc null"
}}
"""

    try:

        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,

            config=types.GenerateContentConfig(
                temperature=0.1,
                max_output_tokens=1000,
                response_mime_type="application/json",
            ),
        )

        text = response.text.strip()

        if not text:
            st.warning(
                "Gemini không trả về dữ liệu."
            )
            return None

        result = json.loads(text)

        required_keys = [
            "ma_nv",
            "ho_ten",
            "ngay_sinh",
            "bo_phan",
            "ngay_het_han_hop_dong",
        ]

        # ----------------------------------------------------
        # Đảm bảo đầy đủ key
        # ----------------------------------------------------

        for key in required_keys:

            if key not in result:
                result[key] = None


        # ----------------------------------------------------
        # Chỉ chấp nhận cột thực sự tồn tại
        # ----------------------------------------------------

        valid_columns = set(columns)

        for key in required_keys:

            value = result.get(key)

            if value is None:
                continue

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
# CHUẨN HÓA TEXT
# ============================================================

def normalize_text(text):

    return (
        str(text)
        .strip()
        .lower()
        .replace("_", " ")
        .replace("-", " ")
    )


# ============================================================
# TÌM CỘT THEO KEYWORD
# ============================================================

def find_column_by_keywords(columns, keywords):

    normalized = {
        col: normalize_text(col)
        for col in columns
    }


    # --------------------------------------------------------
    # Khớp chính xác
    # --------------------------------------------------------

    for col, norm in normalized.items():

        if norm in keywords:
            return col


    # --------------------------------------------------------
    # Khớp chứa keyword
    # --------------------------------------------------------

    for col, norm in normalized.items():

        for keyword in keywords:

            if keyword in norm:
                return col


    return None


# ============================================================
# FALLBACK
# ============================================================

def detect_columns_fallback(columns):

    mapping = {}


    mapping["ma_nv"] = find_column_by_keywords(
        columns,
        [
            "mã nv",
            "ma nv",
            "mã nhân viên",
            "ma nhan vien",
            "employee id",
            "employee code",
            "staff id",
            "staff code",
        ],
    )


    mapping["ho_ten"] = find_column_by_keywords(
        columns,
        [
            "họ và tên",
            "ho va ten",
            "họ tên",
            "ho ten",
            "tên nhân viên",
            "ten nhan vien",
            "full name",
            "employee name",
        ],
    )


    mapping["ngay_sinh"] = find_column_by_keywords(
        columns,
        [
            "ngày sinh",
            "ngay sinh",
            "date of birth",
            "dob",
            "birthday",
        ],
    )


    mapping["bo_phan"] = find_column_by_keywords(
        columns,
        [
            "bộ phận",
            "bo phan",
            "phòng ban",
            "phong ban",
            "đơn vị",
            "don vi",
            "department",
            "division",
        ],
    )


    mapping["ngay_het_han_hop_dong"] = find_column_by_keywords(
        columns,
        [
            "ngày hết hạn hợp đồng",
            "ngay het han hop dong",
            "ngày hết hạn hđ",
            "ngay het han hd",
            "ngày kết thúc hợp đồng",
            "ngay ket thuc hop dong",
            "contract end date",
            "contract expiry date",
            "contract expiration date",
        ],
    )


    return mapping


# ============================================================
# ĐỌC EXCEL
# ============================================================

def read_excel(file_path_or_buffer):

    try:

        df = pd.read_excel(
            file_path_or_buffer,
            engine="openpyxl",
        )

        df = df.dropna(
            how="all"
        ).copy()

        df.columns = [
            str(col).strip()
            for col in df.columns
        ]

        return df


    except Exception as e:

        st.error(
            f"Không thể đọc file Excel: {e}"
        )

        return None


# ============================================================
# CHUẨN HÓA DATAFRAME
# ============================================================

def prepare_dataframe(df, mapping):

    required = [
        "ma_nv",
        "ho_ten",
        "ngay_sinh",
        "bo_phan",
        "ngay_het_han_hop_dong",
    ]


    missing = [
        key
        for key in required
        if not mapping.get(key)
    ]


    if missing:
        return None, missing


    result = pd.DataFrame()


    result["Mã NV"] = df[
        mapping["ma_nv"]
    ]


    result["Họ và tên"] = df[
        mapping["ho_ten"]
    ]


    result["Ngày tháng năm sinh"] = pd.to_datetime(
        df[mapping["ngay_sinh"]],
        errors="coerce",
        dayfirst=True,
    )


    result["Bộ Phận"] = df[
        mapping["bo_phan"]
    ]


    result["Ngày hết hạn hợp đồng"] = pd.to_datetime(
        df[mapping["ngay_het_han_hop_dong"]],
        errors="coerce",
        dayfirst=True,
    )


    return result, []


# ============================================================
# BÁO CÁO SINH NHẬT
# ============================================================

def get_birthday_report(df, month):

    result = df[
        df["Ngày tháng năm sinh"].dt.month == month
    ].copy()


    result = result[
        [
            "Mã NV",
            "Họ và tên",
            "Ngày tháng năm sinh",
            "Bộ Phận",
        ]
    ]


    return result.sort_values(
        by="Ngày tháng năm sinh"
    )


# ============================================================
# BÁO CÁO HỢP ĐỒNG
# ============================================================

def get_contract_report(df, month, year):

    result = df[
        (
            df["Ngày hết hạn hợp đồng"].dt.month
            == month
        )
        &
        (
            df["Ngày hết hạn hợp đồng"].dt.year
            == year
        )
    ].copy()


    result = result[
        [
            "Mã NV",
            "Họ và tên",
            "Ngày tháng năm sinh",
            "Bộ Phận",
            "Ngày hết hạn hợp đồng",
        ]
    ]


    return result.sort_values(
        by="Ngày hết hạn hợp đồng"
    )


# ============================================================
# TẠO EXCEL
# ============================================================

def dataframe_to_excel(df, sheet_name):

    output = io.BytesIO()


    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name,
        )


        worksheet = writer.sheets[
            sheet_name
        ]


        # ----------------------------------------------------
        # Định dạng ngày
        # ----------------------------------------------------

        for row in worksheet.iter_rows():

            for cell in row:

                if hasattr(
                    cell.value,
                    "strftime",
                ):

                    cell.number_format = (
                        "dd/mm/yyyy"
                    )


        # ----------------------------------------------------
        # Tự động chỉnh độ rộng cột
        # ----------------------------------------------------

        for column_cells in worksheet.columns:

            max_length = 0

            column_letter = (
                column_cells[0].column_letter
            )


            for cell in column_cells:

                try:

                    max_length = max(
                        max_length,
                        len(str(cell.value)),
                    )

                except Exception:
                    pass


            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max(max_length + 2, 12),
                40,
            )


    output.seek(0)

    return output


# ============================================================
# LƯU EXCEL
# ============================================================

def save_excel_file(
    df,
    filename,
    sheet_name,
):

    path = REPORT_DIR / filename

    output = dataframe_to_excel(
        df,
        sheet_name,
    )


    with open(
        path,
        "wb",
    ) as f:

        f.write(
            output.getvalue()
        )


    return path


# ============================================================
# TELEGRAM
# ============================================================

def telegram_send_message(message):

    if (
        not TELEGRAM_BOT_TOKEN
        or not TELEGRAM_CHAT_ID
    ):

        return (
            False,
            "Thiếu Token hoặc Chat ID",
        )


    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/"
        f"sendMessage"
    )


    try:

        resp = requests.post(
            url,
            data={
                "chat_id": TELEGRAM_CHAT_ID,
                "text": message,
            },
            timeout=30,
        )


        if resp.ok:
            return True, "OK"

        return False, resp.text


    except Exception as e:

        return False, str(e)


# ============================================================
# TELEGRAM SEND FILE
# ============================================================

def telegram_send_file(
    file_path,
    caption,
):

    if (
        not TELEGRAM_BOT_TOKEN
        or not TELEGRAM_CHAT_ID
    ):

        return (
            False,
            "Thiếu Token hoặc Chat ID",
        )


    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/"
        f"sendDocument"
    )


    try:

        with open(
            file_path,
            "rb",
        ) as f:

            resp = requests.post(
                url,
                data={
                    "chat_id": TELEGRAM_CHAT_ID,
                    "caption": caption,
                },
                files={
                    "document": f
                },
                timeout=60,
            )


        if resp.ok:
            return True, "OK"

        return False, resp.text


    except Exception as e:

        return False, str(e)


# ============================================================
# TEST TELEGRAM
# ============================================================

def telegram_test():

    if not TELEGRAM_BOT_TOKEN:

        return (
            False,
            "Chưa nhập TELEGRAM_BOT_TOKEN",
        )


    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/"
        f"getMe"
    )


    try:

        resp = requests.get(
            url,
            timeout=15,
        )


        if (
            resp.ok
            and resp.json().get("ok")
        ):

            return (
                True,
                resp.json()["result"].get(
                    "username",
                    "",
                ),
            )


        return False, resp.text


    except Exception as e:

        return False, str(e)


# ============================================================
# GỬI BÁO CÁO THÁNG
# ============================================================

def send_monthly_report(
    birthday_df,
    contract_df,
    month,
    year,
):

    birthday_file = save_excel_file(
        birthday_df,
        "danh_sach_sinh_nhat.xlsx",
        "Sinh nhật",
    )


    contract_file = save_excel_file(
        contract_df,
        "danh_sach_den_han_ky_hop_dong.xlsx",
        "Hợp đồng",
    )


    message = f"""
📊 BÁO CÁO NHÂN SỰ THÁNG {month:02d}/{year}

🎂 Sinh nhật trong tháng: {len(birthday_df)} nhân viên

📄 Hợp đồng hết hạn trong tháng: {len(contract_df)} nhân viên

⏰ Thời gian báo cáo:
{datetime.now().strftime("%d/%m/%Y %H:%M")}
"""


    ok, error = telegram_send_message(
        message.strip()
    )


    if not ok:
        return False, error


    ok, error = telegram_send_file(
        birthday_file,
        f"🎂 Danh sách sinh nhật tháng {month:02d}/{year}",
    )


    if not ok:
        return False, error


    ok, error = telegram_send_file(
        contract_file,
        f"📄 Danh sách hợp đồng hết hạn tháng {month:02d}/{year}",
    )


    if not ok:
        return False, error


    return (
        True,
        "Đã gửi thành công",
    )


# ============================================================
# CHỐNG GỬI TRÙNG
# ============================================================

def get_last_sent():

    if not LAST_SENT_FILE.exists():
        return ""


    try:

        return LAST_SENT_FILE.read_text(
            encoding="utf-8"
        ).strip()


    except Exception:
        return ""


# ============================================================
# GHI LẠI LẦN GỬI
# ============================================================

def set_last_sent(
    month,
    year,
):

    LAST_SENT_FILE.write_text(
        f"{year:04d}-{month:02d}",
        encoding="utf-8",
    )


# ============================================================
# KIỂM TRA ĐÃ GỬI
# ============================================================

def already_sent(
    month,
    year,
):

    return (
        get_last_sent()
        == f"{year:04d}-{month:02d}"
    )


# ============================================================
# TỰ ĐỘNG GỬI BÁO CÁO
# ============================================================

def automatic_monthly_check(
    prepared_df,
):

    now = datetime.now()


    # Chỉ chạy ngày 01
    if now.day != 1:
        return


    # Chỉ chạy từ 10:00
    if now.hour < 10:
        return


    month = now.month
    year = now.year


    # Không gửi trùng
    if already_sent(
        month,
        year,
    ):
        return


    birthday_df = get_birthday_report(
        prepared_df,
        month,
    )


    contract_df = get_contract_report(
        prepared_df,
        month,
        year,
    )


    success, message = send_monthly_report(
        birthday_df,
        contract_df,
        month,
        year,
    )


    if success:

        set_last_sent(
            month,
            year,
        )

        st.success(
            "🤖 Đã tự động gửi báo cáo Telegram."
        )

    else:

        st.error(
            "Không thể tự động gửi Telegram: "
            + message
        )


# ============================================================
# GIAO DIỆN
# ============================================================

st.markdown(
    '<div class="main-title">'
    '📊 QUẢN LÝ SINH NHẬT & HỢP ĐỒNG'
    '</div>',
    unsafe_allow_html=True,
)


st.markdown(
    '<div class="sub-title">'
    'Theo dõi sinh nhật nhân viên và hợp đồng đến hạn ký'
    '</div>',
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Cấu hình")


    current_date = date.today()


    selected_year = st.selectbox(
        "Năm",
        list(
            range(
                current_date.year - 2,
                current_date.year + 6,
            )
        ),
        index=2,
    )


    selected_month = st.selectbox(
        "Tháng",
        list(range(1, 13)),
        index=current_date.month - 1,
        format_func=lambda x:
        f"Tháng {x:02d}",
    )


    st.divider()


    # --------------------------------------------------------
    # GEMINI
    # --------------------------------------------------------

    st.subheader("🤖 Gemini AI")


    if GEMINI_API_KEY:

        st.success(
            f"🟢 Gemini API đã cấu hình\n\n"
            f"Model: {GEMINI_MODEL}"
        )

    else:

        st.warning(
            "🔴 Chưa có Gemini API Key"
        )


    # --------------------------------------------------------
    # TELEGRAM
    # --------------------------------------------------------

    st.subheader("📱 Telegram")


    if (
        TELEGRAM_BOT_TOKEN
        and TELEGRAM_CHAT_ID
    ):

        st.success(
            "🟢 Telegram đã cấu hình"
        )

    else:

        st.warning(
            "🔴 Chưa cấu hình Telegram"
        )


# ============================================================
# 1. QUẢN LÝ FILE DỮ LIỆU
# ============================================================

st.header(
    "1️⃣ File danh sách nhân sự"
)


uploaded_file = st.file_uploader(
    "Cập nhật file Excel danh sách nhân sự "
    "(Nếu có nhân viên mới/cũ)",
    type=[
        "xlsx",
        "xls",
    ],
)


if uploaded_file is not None:

    try:

        with open(
            CURRENT_DB_FILE,
            "wb",
        ) as f:

            f.write(
                uploaded_file.getbuffer()
            )


        # ----------------------------------------------------
        # Reset mapping khi upload file mới
        # ----------------------------------------------------

        st.session_state.column_mapping = None


        st.success(
            "✅ Đã cập nhật cơ sở dữ liệu nhân sự mới nhất!"
        )


    except Exception as e:

        st.error(
            f"Không thể lưu file: {e}"
        )


# ============================================================
# KIỂM TRA FILE DATABASE
# ============================================================

if not CURRENT_DB_FILE.exists():

    st.info(
        "📁 Hệ thống trống. "
        "Hãy upload file Excel danh sách nhân sự "
        "lần đầu để hệ thống bắt đầu hoạt động."
    )

    st.stop()


# ============================================================
# ĐỌC DATABASE
# ============================================================

df_original = read_excel(
    CURRENT_DB_FILE
)


if df_original is None:
    st.stop()


st.success(
    f"Đang sử dụng dữ liệu gồm "
    f"{len(df_original):,} nhân sự."
)


# ============================================================
# XEM FILE GỐC
# ============================================================

with st.expander(
    "👁️ Xem dữ liệu Excel gốc đang lưu"
):

    st.dataframe(
        df_original,
        use_container_width=True,
        height=300,
    )


# ============================================================
# 2. NHẬN DIỆN CỘT
# ============================================================

st.header(
    "2️⃣ Nhận diện cấu trúc Excel"
)


if "column_mapping" not in st.session_state:

    st.session_state.column_mapping = None


# ============================================================
# NÚT GEMINI
# ============================================================

if st.button(
    "🤖 Phân tích cột bằng Gemini 2.5 Pro",
    type="primary",
):

    if not GEMINI_API_KEY:

        st.error(
            "Bạn chưa cấu hình GEMINI_API_KEY"
        )

    else:

        with st.spinner(
            "Gemini 2.5 Pro đang phân tích cấu trúc Excel..."
        ):

            mapping = detect_columns_with_gemini(
                list(df_original.columns)
            )


        if mapping is not None:

            st.session_state.column_mapping = mapping

            st.success(
                "✅ Gemini 2.5 Pro đã nhận diện xong cấu trúc Excel."
            )

        else:

            st.warning(
                "⚠️ Không thể nhận diện bằng Gemini. "
                "Hệ thống sẽ sử dụng quy tắc dự phòng."
            )


# ============================================================
# FALLBACK
# ============================================================

if (
    st.session_state.column_mapping
    is None
):

    mapping = detect_columns_fallback(
        list(df_original.columns)
    )

else:

    mapping = (
        st.session_state.column_mapping
    )


# ============================================================
# HIỂN THỊ MAPPING
# ============================================================

mapping_names = {

    "ma_nv":
        "Mã nhân viên",

    "ho_ten":
        "Họ và tên",

    "ngay_sinh":
        "Ngày sinh",

    "bo_phan":
        "Bộ phận",

    "ngay_het_han_hop_dong":
        "Ngày hết hạn hợp đồng",
}


mapping_df = pd.DataFrame(
    [
        {
            "Trường hệ thống":
                mapping_names[key],

            "Cột Excel":
                mapping.get(key),
        }

        for key in mapping_names
    ]
)


st.dataframe(
    mapping_df,
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# CHUẨN HÓA DATA
# ============================================================

prepared_df, missing = prepare_dataframe(
    df_original,
    mapping,
)


if missing:

    st.error(
        "Không tìm thấy các trường bắt buộc: "
        +
        ", ".join(
            mapping_names[x]
            for x in missing
        )
    )

    st.stop()


# ============================================================
# 3. BÁO CÁO
# ============================================================

birthday_df = get_birthday_report(
    prepared_df,
    selected_month,
)


contract_df = get_contract_report(
    prepared_df,
    selected_month,
    selected_year,
)


st.header(
    f"3️⃣ Báo cáo tháng "
    f"{selected_month:02d}/{selected_year}"
)


# ============================================================
# METRICS
# ============================================================

col1, col2, col3 = st.columns(3)


with col1:

    st.metric(
        "🎂 Sinh nhật",
        f"{len(birthday_df):,}",
    )


with col2:

    st.metric(
        "📄 Hợp đồng hết hạn",
        f"{len(contract_df):,}",
    )


with col3:

    st.metric(
        "👥 Tổng nhân sự",
        f"{len(prepared_df):,}",
    )


# ============================================================
# SINH NHẬT
# ============================================================

st.subheader(
    f"🎂 Danh sách sinh nhật tháng "
    f"{selected_month:02d}"
)


if birthday_df.empty:

    st.info(
        "Không có nhân viên sinh nhật."
    )

else:

    bday_disp = birthday_df.copy()


    bday_disp[
        "Ngày tháng năm sinh"
    ] = (
        bday_disp[
            "Ngày tháng năm sinh"
        ]
        .dt.strftime("%d/%m/%Y")
    )


    st.dataframe(
        bday_disp,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# DOWNLOAD SINH NHẬT
# ============================================================

birthday_excel = dataframe_to_excel(
    birthday_df,
    "Sinh nhật",
)


st.download_button(
    "⬇️ Download danh sách sinh nhật",

    data=birthday_excel.getvalue(),

    file_name=(
        f"danh_sach_sinh_nhat_"
        f"{selected_month:02d}_"
        f"{selected_year}.xlsx"
    ),

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)


# ============================================================
# HỢP ĐỒNG
# ============================================================

st.subheader(
    f"📄 Danh sách hợp đồng hết hạn tháng "
    f"{selected_month:02d}/{selected_year}"
)


if contract_df.empty:

    st.info(
        "Không có hợp đồng hết hạn."
    )

else:

    cntr_disp = contract_df.copy()


    cntr_disp[
        "Ngày tháng năm sinh"
    ] = (
        cntr_disp[
            "Ngày tháng năm sinh"
        ]
        .dt.strftime("%d/%m/%Y")
    )


    cntr_disp[
        "Ngày hết hạn hợp đồng"
    ] = (
        cntr_disp[
            "Ngày hết hạn hợp đồng"
        ]
        .dt.strftime("%d/%m/%Y")
    )


    st.dataframe(
        cntr_disp,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# DOWNLOAD HỢP ĐỒNG
# ============================================================

contract_excel = dataframe_to_excel(
    contract_df,
    "Hợp đồng",
)


st.download_button(
    "⬇️ Download danh sách hợp đồng",

    data=contract_excel.getvalue(),

    file_name=(
        f"danh_sach_den_han_ky_hop_dong_"
        f"{selected_month:02d}_"
        f"{selected_year}.xlsx"
    ),

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)


# ============================================================
# 4. TELEGRAM
# ============================================================

st.divider()

st.header(
    "4️⃣ Telegram Bot"
)


t_col1, t_col2 = st.columns(2)


# ============================================================
# TRẠNG THÁI BOT
# ============================================================

with t_col1:

    st.write(
        "### 📱 Trạng thái Bot"
    )


    if (
        TELEGRAM_BOT_TOKEN
        and TELEGRAM_CHAT_ID
    ):

        ok, bot_name = telegram_test()


        if ok:

            st.success(
                f"🟢 Bot đang hoạt động: @{bot_name}"
            )

        else:

            st.error(
                f"🔴 Bot lỗi: {bot_name}"
            )

    else:

        st.warning(
            "Chưa cấu hình Telegram."
        )


# ============================================================
# GỬI THỦ CÔNG
# ============================================================

with t_col2:

    st.write(
        "### 📤 Gửi báo cáo thủ công"
    )


    if st.button(
        "📤 Gửi báo cáo ngay",
        type="primary",
        use_container_width=True,
    ):

        if (
            not TELEGRAM_BOT_TOKEN
            or not TELEGRAM_CHAT_ID
        ):

            st.error(
                "❌ Chưa cấu hình Telegram."
            )

        else:

            with st.spinner(
                "Đang gửi báo cáo lên Telegram..."
            ):

                success, message = (
                    send_monthly_report(
                        birthday_df,
                        contract_df,
                        selected_month,
                        selected_year,
                    )
                )


            if success:

                st.success(
                    "✅ Đã gửi thành công "
                    "2 file Excel lên Telegram."
                )

            else:

                st.error(
                    f"❌ Gửi thất bại: {message}"
                )


# ============================================================
# 5. TỰ ĐỘNG GỬI
# ============================================================

st.divider()

st.header(
    "5️⃣ Tự động gửi báo cáo hàng tháng"
)


st.info(
    "🤖 Hệ thống sẽ tự động gửi báo cáo "
    "khi web được truy cập vào ngày 01 "
    "từ 10:00 trở đi. "
    "Dữ liệu lấy từ file bạn đã lưu."
)


now = datetime.now()


st.write(
    "🕐 Thời gian máy chủ hiện tại: "
    f"**{now.strftime('%d/%m/%Y %H:%M:%S')}**"
)


last_sent = get_last_sent()


if last_sent:

    st.write(
        "📌 Báo cáo gần nhất đã tự động gửi: "
        f"**{last_sent}**"
    )

else:

    st.write(
        "📌 Chưa có báo cáo tự động nào được gửi."
    )


# ============================================================
# KÍCH HOẠT TỰ ĐỘNG
# ============================================================

automatic_monthly_check(
    prepared_df
)
```
