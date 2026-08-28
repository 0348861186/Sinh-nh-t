import os
import io
import json
import calendar
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

from dotenv import load_dotenv
from google import genai


# ============================================================
# CẤU HÌNH
# ============================================================

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "").strip()

REPORT_DIR = Path("reports")
REPORT_DIR.mkdir(exist_ok=True)

LAST_SENT_FILE = REPORT_DIR / "last_sent.txt"


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
# HÀM GEMINI
# ============================================================

def get_gemini_client():

    if not GEMINI_API_KEY:
        return None

    try:
        return genai.Client(api_key=GEMINI_API_KEY)
    except Exception:
        return None


def detect_columns_with_gemini(columns):

    client = get_gemini_client()

    if client is None:
        return None

    column_list = "\n".join(
        [f"- {str(col)}" for col in columns]
    )

    prompt = f"""
Bạn là hệ thống AI chuyên phân tích file Excel nhân sự.

Danh sách tên cột thực tế:

{column_list}

Hãy xác định cột nào tương ứng với các trường sau:

1. ma_nv
2. ho_ten
3. ngay_sinh
4. bo_phan
5. ngay_het_han_hop_dong

Các cột có thể có tên tiếng Việt hoặc tiếng Anh,
viết tắt hoặc cách gọi tương đương.

Ví dụ:
"Mã NV", "Mã nhân viên", "Employee ID" -> ma_nv
"Họ tên", "Họ và tên", "Tên nhân viên" -> ho_ten
"Ngày sinh", "NS", "DOB" -> ngay_sinh
"Bộ phận", "Phòng ban", "Đơn vị" -> bo_phan
"Ngày hết hạn HĐ", "Ngày kết thúc hợp đồng",
"Contract End Date" -> ngay_het_han_hop_dong

Chỉ trả về JSON hợp lệ theo mẫu:

{{
    "ma_nv": "tên cột",
    "ho_ten": "tên cột",
    "ngay_sinh": "tên cột",
    "bo_phan": "tên cột",
    "ngay_het_han_hop_dong": "tên cột"
}}

Nếu không tìm thấy trường nào thì giá trị là null.
Không giải thích thêm.
"""

    try:

        response = client.models.generate_content(
            model="gemini-3-flash-preview",
            contents=prompt,
        )

        text = response.text.strip()

        # Loại markdown ```json nếu Gemini trả về
        if text.startswith("```"):
            text = text.replace("```json", "")
            text = text.replace("```", "")
            text = text.strip()

        result = json.loads(text)

        return result

    except Exception as e:

        st.warning(
            f"Gemini không thể nhận diện cột: {e}"
        )

        return None


# ============================================================
# NHẬN DIỆN CỘT BẰNG QUY TẮC DỰ PHÒNG
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

    normalized = {
        col: normalize_text(col)
        for col in columns
    }

    # Ưu tiên khớp chính xác
    for col, norm in normalized.items():

        if norm in keywords:
            return col

    # Sau đó kiểm tra chứa từ khóa
    for col, norm in normalized.items():

        for keyword in keywords:

            if keyword in norm:
                return col

    return None


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
            "id nhân viên",
            "id nhan vien",
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
            "name",
        ],
    )

    mapping["ngay_sinh"] = find_column_by_keywords(
        columns,
        [
            "ngày sinh",
            "ngay sinh",
            "ns",
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
            "dept",
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
            "ngày kết thúc hđ",
            "contract end date",
            "contract expiration",
        ],
    )

    return mapping


# ============================================================
# ĐỌC EXCEL
# ============================================================

def read_excel(uploaded_file):

    try:

        df = pd.read_excel(
            uploaded_file,
            engine="openpyxl"
        )

        # Xóa các dòng hoàn toàn trống
        df = df.dropna(
            how="all"
        ).copy()

        # Chuẩn hóa tên cột
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

    missing = []

    for key in required:

        if not mapping.get(key):

            missing.append(key)

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


# ============================================================
# LỌC SINH NHẬT
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

    result = result.sort_values(
        by="Ngày tháng năm sinh"
    )

    return result


# ============================================================
# LỌC HỢP ĐỒNG
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

    result = result.sort_values(
        by="Ngày hết hạn hợp đồng"
    )

    return result


# ============================================================
# TẠO EXCEL
# ============================================================

def dataframe_to_excel(df, sheet_name):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name=sheet_name,
        )

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Định dạng ngày
        for row in worksheet.iter_rows():

            for cell in row:

                if hasattr(cell.value, "strftime"):

                    cell.number_format = "dd/mm/yyyy"

        # Tự động độ rộng cột
        for column_cells in worksheet.columns:

            max_length = 0

            column_letter = (
                column_cells[0].column_letter
            )

            for cell in column_cells:

                try:

                    value_length = len(
                        str(cell.value)
                    )

                    max_length = max(
                        max_length,
                        value_length
                    )

                except Exception:
                    pass

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max(max_length + 2, 12),
                40
            )

    output.seek(0)

    return output


# ============================================================
# LƯU FILE
# ============================================================

def save_excel_file(
    df,
    filename,
    sheet_name
):

    path = REPORT_DIR / filename

    output = dataframe_to_excel(
        df,
        sheet_name
    )

    with open(path, "wb") as f:

        f.write(
            output.getvalue()
        )

    return path


# ============================================================
# TELEGRAM
# ============================================================

def telegram_send_message(message):

    if not TELEGRAM_BOT_TOKEN:
        return False, "Chưa có TELEGRAM_BOT_TOKEN"

    if not TELEGRAM_CHAT_ID:
        return False, "Chưa có TELEGRAM_CHAT_ID"

    url = (
        "https://api.telegram.org/bot"
        f"{TELEGRAM_BOT_TOKEN}/sendMessage"
    )

    try:

        response = requests.post(
            url,
            data={
                "chat_id": TELEGRAM_CHAT_ID,
                "text": message,
            },
            timeout=30,
        )

        if response.ok:

            return True, "OK"

        return False, response.text

    except Exception as e:

        return False, str(e)


def telegram_send_file(
    file_path,
    caption
):

    if not TELEGRAM_BOT_TOKEN:
        return False, "Chưa có TELEGRAM_BOT_TOKEN"

    if not TELEGRAM_CHAT_ID:
        return False, "Chưa có TELEGRAM_CHAT_ID"

    url = (
        "https://api.telegram.org/bot"
        f"{TELEGRAM_BOT_TOKEN}/sendDocument"
    )

    try:

        with open(file_path, "rb") as f:

            response = requests.post(
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

        if response.ok:

            return True, "OK"

        return False, response.text

    except Exception as e:

        return False, str(e)


# ============================================================
# KIỂM TRA TELEGRAM
# ============================================================

def telegram_test():

    if not TELEGRAM_BOT_TOKEN:

        return False, "Chưa nhập TELEGRAM_BOT_TOKEN"

    url = (
        "https://api.telegram.org/bot"
        f"{TELEGRAM_BOT_TOKEN}/getMe"
    )

    try:

        response = requests.get(
            url,
            timeout=15
        )

        if response.ok:

            data = response.json()

            if data.get("ok"):

                bot_name = data[
                    "result"
                ].get("username", "")

                return True, bot_name

        return False, response.text

    except Exception as e:

        return False, str(e)


# ============================================================
# GỬI BÁO CÁO TELEGRAM
# ============================================================

def send_monthly_report(
    birthday_df,
    contract_df,
    month,
    year
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

🎂 Sinh nhật trong tháng:
{len(birthday_df)} nhân viên

📄 Hợp đồng hết hạn trong tháng:
{len(contract_df)} nhân viên

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

    return True, "Đã gửi thành công"


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


def set_last_sent(month, year):

    LAST_SENT_FILE.write_text(
        f"{year:04d}-{month:02d}",
        encoding="utf-8",
    )


def already_sent(month, year):

    return (
        get_last_sent()
        == f"{year:04d}-{month:02d}"
    )


# ============================================================
# KIỂM TRA TỰ ĐỘNG GỬI
# ============================================================

def automatic_monthly_check(
    birthday_df,
    contract_df
):

    now = datetime.now()

    # Chỉ chạy ngày 01
    if now.day != 1:
        return

    # Chỉ chạy từ 10:00 trở đi
    if now.hour < 10:
        return

    month = now.month
    year = now.year

    # Không gửi trùng
    if already_sent(month, year):

        return

    success, message = send_monthly_report(
        birthday_df,
        contract_df,
        month,
        year,
    )

    if success:

        set_last_sent(
            month,
            year
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
                current_date.year + 6
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

    st.subheader("🤖 Gemini AI")

    if GEMINI_API_KEY:

        st.success(
            "🟢 Gemini API đã cấu hình"
        )

    else:

        st.warning(
            "🔴 Chưa có Gemini API Key"
        )

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
# UPLOAD EXCEL
# ============================================================

st.header("1️⃣ File danh sách nhân sự")

uploaded_file = st.file_uploader(
    "Chọn file Excel danh sách nhân sự",
    type=["xlsx", "xls"],
)


if uploaded_file is None:

    st.info(
        "📁 Hãy upload file Excel danh sách nhân sự để bắt đầu."
    )

    st.stop()


# ============================================================
# ĐỌC EXCEL
# ============================================================

df_original = read_excel(
    uploaded_file
)

if df_original is None:

    st.stop()


st.success(
    f"Đã đọc {len(df_original):,} dòng dữ liệu."
)


with st.expander(
    "👁️ Xem dữ liệu Excel gốc"
):

    st.dataframe(
        df_original,
        use_container_width=True,
        height=300,
    )


# ============================================================
# NHẬN DIỆN CỘT
# ============================================================

st.header("2️⃣ Nhận diện cấu trúc Excel")

if "column_mapping" not in st.session_state:

    st.session_state.column_mapping = None


if st.button(
    "🤖 Phân tích cột bằng Gemini AI",
    type="primary",
):

    if not GEMINI_API_KEY:

        st.error(
            "Bạn chưa cấu hình GEMINI_API_KEY trong file .env"
        )

    else:

        with st.spinner(
            "Gemini đang phân tích cấu trúc Excel..."
        ):

            mapping = detect_columns_with_gemini(
                list(df_original.columns)
            )

            st.session_state.column_mapping = mapping


# Nếu chưa chạy Gemini, sử dụng fallback
if st.session_state.column_mapping is None:

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
    "ma_nv": "Mã nhân viên",
    "ho_ten": "Họ và tên",
    "ngay_sinh": "Ngày sinh",
    "bo_phan": "Bộ phận",
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
# KIỂM TRA MAPPING
# ============================================================

prepared_df, missing = prepare_dataframe(
    df_original,
    mapping,
)

if missing:

    st.error(
        "Không tìm thấy các trường bắt buộc: "
        + ", ".join(
            mapping_names[x]
            for x in missing
        )
    )

    st.warning(
        "Hãy bấm 'Phân tích cột bằng Gemini AI' "
        "hoặc kiểm tra lại file Excel."
    )

    st.stop()


# ============================================================
# XỬ LÝ DỮ LIỆU
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


# ============================================================
# DASHBOARD
# ============================================================

st.header(
    f"3️⃣ Báo cáo tháng {selected_month:02d}/{selected_year}"
)


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
    f"🎂 Danh sách sinh nhật tháng {selected_month:02d}"
)

if birthday_df.empty:

    st.info(
        "Không có nhân viên sinh nhật trong tháng này."
    )

else:

    birthday_display = birthday_df.copy()

    birthday_display[
        "Ngày tháng năm sinh"
    ] = birthday_display[
        "Ngày tháng năm sinh"
    ].dt.strftime(
        "%d/%m/%Y"
    )

    st.dataframe(
        birthday_display,
        use_container_width=True,
        hide_index=True,
    )


birthday_excel = dataframe_to_excel(
    birthday_df,
    "Sinh nhật",
)

st.download_button(
    label="⬇️ Download danh sách sinh nhật",
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
    f"📄 Danh sách hợp đồng hết hạn "
    f"tháng {selected_month:02d}/{selected_year}"
)

if contract_df.empty:

    st.info(
        "Không có hợp đồng hết hạn trong tháng này."
    )

else:

    contract_display = contract_df.copy()

    contract_display[
        "Ngày tháng năm sinh"
    ] = contract_display[
        "Ngày tháng năm sinh"
    ].dt.strftime(
        "%d/%m/%Y"
    )

    contract_display[
        "Ngày hết hạn hợp đồng"
    ] = contract_display[
        "Ngày hết hạn hợp đồng"
    ].dt.strftime(
        "%d/%m/%Y"
    )

    st.dataframe(
        contract_display,
        use_container_width=True,
        hide_index=True,
    )


contract_excel = dataframe_to_excel(
    contract_df,
    "Hợp đồng",
)

st.download_button(
    label="⬇️ Download danh sách hợp đồng",
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
# TELEGRAM
# ============================================================

st.divider()

st.header("4️⃣ Telegram Bot")


telegram_col1, telegram_col2 = st.columns(2)


with telegram_col1:

    st.write("### 📱 Trạng thái Bot")

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


with telegram_col2:

    st.write("### 📤 Gửi báo cáo")

    st.write(
        f"Tháng đang chọn: "
        f"**{selected_month:02d}/{selected_year}**"
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
                "Chưa cấu hình Telegram."
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
                    "✅ Đã gửi thành công 2 file Excel lên Telegram."
                )

            else:

                st.error(
                    f"❌ Gửi thất bại: {message}"
                )


# ============================================================
# TỰ ĐỘNG GỬI NGÀY 01 LÚC 10:00
# ============================================================

st.divider()

st.header(
    "5️⃣ Tự động gửi báo cáo hàng tháng"
)

st.info(
    """
    🤖 Hệ thống sẽ tự động gửi báo cáo khi ứng dụng
    được chạy vào ngày 01 từ 10:00 trở đi.

    Báo cáo tự động lấy tháng hiện tại.

    Hệ thống lưu tháng đã gửi để tránh gửi trùng.
    """
)


now = datetime.now()

st.write(
    f"🕐 Thời gian máy chủ hiện tại: "
    f"**{now.strftime('%d/%m/%Y %H:%M:%S')}**"
)

last_sent = get_last_sent()

if last_sent:

    st.write(
        f"📌 Báo cáo gần nhất đã gửi: **{last_sent}**"
    )

else:

    st.write(
        "📌 Chưa có báo cáo tự động nào được gửi."
    )


# Chạy kiểm tra tự động
automatic_monthly_check(
    birthday_df,
    contract_df,
)


# ============================================================
# THÔNG TIN
# ============================================================

st.divider()

with st.expander(
    "ℹ️ Thông tin hệ thống"
):

    st.write(
        """
        **Chức năng:**

        - Đọc file Excel nhân sự.
        - Gemini AI nhận diện cột.
        - Thống kê sinh nhật theo tháng.
        - Thống kê hợp đồng hết hạn theo tháng.
        - Xuất Excel.
        - Download Excel từ Dashboard.
        - Gửi Excel qua Telegram.
        - Gửi thủ công bằng nút "Gửi báo cáo ngay".
        - Tự động gửi ngày 01 hàng tháng từ 10:00.
        - Chống gửi trùng báo cáo.

        **Lưu ý:**

        Máy tính/server phải đang chạy ứng dụng
        Streamlit thì chức năng tự động gửi mới được thực hiện.
        """
    )
