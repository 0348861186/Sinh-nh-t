import io
import json
import hashlib
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from google import genai


# ============================================================
# 5-TẦNG HỆ THỐNG QUẢN LÝ NHÂN SỰ EXCEL
#
# TẦNG 1 - EXCEL INTELLIGENCE
#   Đọc workbook, nhiều sheet, header + dữ liệu mẫu, AI hiểu cấu trúc.
#
# TẦNG 2 - SCHEMA ENGINE
#   Chuyển Excel bất kỳ về schema nhân sự chuẩn.
#
# TẦNG 3 - VALIDATION ENGINE
#   Kiểm tra mapping, kiểu dữ liệu, tỷ lệ lỗi, dữ liệu thiếu,
#   điểm tin cậy và chặn dữ liệu không đủ chắc chắn.
#
# TẦNG 4 - BUSINESS ENGINE
#   Sinh nhật, hợp đồng, dashboard, export.
#
# TẦNG 5 - AUTOMATION ENGINE
#   Telegram, cron, chống gửi trùng, log trạng thái.
#
# Gemini:
#   Interactions API + gemini-3.1-pro-preview.
#   API key đặt trong Streamlit Secrets:
#       GEMINI_API_KEY = "..."
#       TELEGRAM_BOT_TOKEN = "..."
#       TELEGRAM_CHAT_ID = "..."
#
# Cài:
#   pip install -r requirements.txt
#
# Chạy:
#   streamlit run app.py
#
# Cron:
#   https://YOUR-APP/?cron=1
# ============================================================


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Quản lý nhân sự AI",
    page_icon="📊",
    layout="wide",
)

GEMINI_API_KEY = str(st.secrets.get("GEMINI_API_KEY", "")).strip()
TELEGRAM_BOT_TOKEN = str(st.secrets.get("TELEGRAM_BOT_TOKEN", "")).strip()
TELEGRAM_CHAT_ID = str(st.secrets.get("TELEGRAM_CHAT_ID", "")).strip()

# Gemini 3.1 Pro Preview hiện là model phù hợp cho luồng AI
# phân tích cấu trúc + structured output.
GEMINI_MODEL = "gemini-3.1-pro-preview"

DATA_DIR = Path("data")
REPORT_DIR = Path("reports")
STATE_DIR = Path("state")

DATA_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)
STATE_DIR.mkdir(exist_ok=True)

CURRENT_DB_FILE = DATA_DIR / "current_hr_data.xlsx"
SCHEMA_FILE = STATE_DIR / "schema.json"
SOURCE_HASH_FILE = STATE_DIR / "source_hash.txt"
AUTOMATION_LOG_FILE = STATE_DIR / "automation_log.json"

REQUIRED_FIELDS = [
    "ma_nv",
    "ho_ten",
    "ngay_sinh",
    "bo_phan",
    "ngay_het_han_hop_dong",
]

FIELD_LABELS = {
    "ma_nv": "Mã nhân viên",
    "ho_ten": "Họ và tên",
    "ngay_sinh": "Ngày sinh",
    "bo_phan": "Bộ phận",
    "ngay_het_han_hop_dong": "Ngày hết hạn hợp đồng",
}

FIELD_DESCRIPTIONS = {
    "ma_nv": "Mã nhân viên, employee ID, staff ID, employee code.",
    "ho_ten": "Họ tên đầy đủ của nhân viên, full name, employee name.",
    "ngay_sinh": "Ngày sinh, date of birth, DOB, birthday.",
    "bo_phan": "Bộ phận, phòng ban, đơn vị, department, division.",
    "ngay_het_han_hop_dong": (
        "Ngày hết hạn/kết thúc hợp đồng lao động, contract end/expiry date. "
        "Không nhầm với ngày bắt đầu hợp đồng."
    ),
}


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .sub-title {
        color: #666;
        margin-bottom: 24px;
    }
    .small-note {
        color: #666;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# TẦNG 1 - EXCEL INTELLIGENCE
# ============================================================

def normalize_text(value):
    """Chuẩn hóa chuỗi để rule engine so sánh."""
    if value is None:
        return ""
    text = str(value).strip().lower()
    replacements = {
        "_": " ",
        "-": " ",
        "/": " ",
        "\\": " ",
        ".": " ",
        "(": " ",
        ")": " ",
        ":": " ",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return " ".join(text.split())


def safe_cell(value):
    """Biến dữ liệu mẫu thành dạng JSON/string ổn định."""
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return str(value)


def read_workbook(path_or_buffer):
    """
    Đọc workbook bằng pandas.
    Trả về:
        {
          "sheet_name": {
             "df": dataframe,
             "columns": [...]
          }
        }
    """
    try:
        excel = pd.ExcelFile(path_or_buffer)
        sheets = {}

        for sheet_name in excel.sheet_names:
            try:
                df = pd.read_excel(
                    excel,
                    sheet_name=sheet_name,
                    dtype=object,
                )
                df = df.dropna(how="all").copy()
                df.columns = [str(c).strip() for c in df.columns]

                # Loại các cột không có tên thực.
                if len(df.columns):
                    df = df.loc[:, [str(c).strip() not in ("", "nan", "None")
                                   for c in df.columns]]

                sheets[str(sheet_name)] = {
                    "df": df,
                    "columns": list(df.columns),
                }
            except Exception:
                # Bỏ qua sheet lỗi, không làm hỏng toàn workbook.
                continue

        return sheets, None

    except Exception as exc:
        return {}, str(exc)


def workbook_fingerprint(path):
    """Hash file để biết workbook đã thay đổi."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def build_sheet_intelligence(sheets):
    """
    Tạo hồ sơ cho từng sheet:
      - tên sheet
      - số dòng/cột
      - tên cột
      - kiểu dữ liệu suy luận
      - tối đa 8 dòng mẫu
    """
    profiles = []

    for sheet_name, info in sheets.items():
        df = info["df"]

        sample_rows = []
        if not df.empty:
            sample = df.head(8)
            for _, row in sample.iterrows():
                sample_rows.append({
                    str(col): safe_cell(row[col])
                    for col in df.columns
                })

        dtypes = {}
        for col in df.columns:
            series = df[col]
            non_null = series.dropna()
            if non_null.empty:
                inferred = "empty"
            else:
                date_probe = pd.to_datetime(
                    non_null.astype(str),
                    errors="coerce",
                    dayfirst=True,
                )
                date_ratio = float(date_probe.notna().mean())
                if date_ratio >= 0.75:
                    inferred = "date-like"
                else:
                    inferred = str(series.dtype)

            dtypes[str(col)] = inferred

        profiles.append({
            "sheet_name": sheet_name,
            "rows": int(len(df)),
            "columns_count": int(len(df.columns)),
            "columns": [str(c) for c in df.columns],
            "inferred_types": dtypes,
            "sample_rows": sample_rows,
        })

    return profiles


def get_gemini_client():
    if not GEMINI_API_KEY:
        return None, "Chưa cấu hình GEMINI_API_KEY."

    try:
        return genai.Client(api_key=GEMINI_API_KEY), None
    except Exception as exc:
        return None, f"Không thể khởi tạo Gemini Client: {exc}"


def ai_analyze_workbook(sheet_profiles):
    """
    AI không chỉ nhìn tên cột.
    Nó nhìn:
      - sheet
      - tên cột
      - kiểu dữ liệu
      - dữ liệu mẫu
    và trả về schema có confidence + lý do.
    """

    client, error = get_gemini_client()
    if client is None:
        return None, error

    schema = {
        "type": "object",
        "properties": {
            "selected_sheet": {
                "type": "string",
                "description": "Tên sheet chứa bảng nhân sự chính."
            },
            "confidence": {
                "type": "number",
                "minimum": 0,
                "maximum": 1,
                "description": "Độ tin cậy tổng thể từ 0 đến 1."
            },
            "reason": {
                "type": "string",
                "description": "Giải thích ngắn gọn."
            },
            "mapping": {
                "type": "object",
                "properties": {
                    "ma_nv": {"type": ["string", "null"]},
                    "ho_ten": {"type": ["string", "null"]},
                    "ngay_sinh": {"type": ["string", "null"]},
                    "bo_phan": {"type": ["string", "null"]},
                    "ngay_het_han_hop_dong": {"type": ["string", "null"]},
                },
                "required": REQUIRED_FIELDS,
            },
            "field_confidence": {
                "type": "object",
                "properties": {
                    "ma_nv": {"type": "number", "minimum": 0, "maximum": 1},
                    "ho_ten": {"type": "number", "minimum": 0, "maximum": 1},
                    "ngay_sinh": {"type": "number", "minimum": 0, "maximum": 1},
                    "bo_phan": {"type": "number", "minimum": 0, "maximum": 1},
                    "ngay_het_han_hop_dong": {
                        "type": "number",
                        "minimum": 0,
                        "maximum": 1,
                    },
                },
                "required": REQUIRED_FIELDS,
            },
        },
        "required": [
            "selected_sheet",
            "confidence",
            "reason",
            "mapping",
            "field_confidence",
        ],
    }

    prompt = f"""
Bạn là AI chuyên gia về dữ liệu nhân sự và Excel.

MỤC TIÊU:
Người dùng có thể upload BẤT KỲ file Excel nhân sự nào.
Bạn phải hiểu cấu trúc thực tế của workbook và chuyển nó về
5 trường chuẩn của hệ thống.

QUAN TRỌNG:
- Không được chỉ dựa vào tên cột.
- Phải xem cả dữ liệu mẫu và kiểu dữ liệu.
- Có thể có nhiều sheet; chọn sheet chứa bảng nhân sự chính.
- Chỉ được trả về tên sheet/cột thực sự tồn tại.
- Không được tự tạo tên cột.
- Nếu không đủ bằng chứng, trả null và confidence thấp.
- Không được nhầm "ngày bắt đầu hợp đồng" với
  "ngày hết hạn hợp đồng".
- Không được nhầm "ngày sinh" với ngày vào làm.
- Không được nhầm mã nhân viên với số thứ tự STT nếu có bằng chứng khác.
- Không dùng các dòng tổng cộng, tiêu đề phụ, ghi chú làm dữ liệu nhân viên.
- Mapping phải ưu tiên ý nghĩa dữ liệu thực tế.

5 TRƯỜNG BẮT BUỘC:
{json.dumps(FIELD_DESCRIPTIONS, ensure_ascii=False, indent=2)}

HỒ SƠ WORKBOOK:
{json.dumps(sheet_profiles, ensure_ascii=False, indent=2)}

Trả về JSON đúng schema.
"""

    try:
        interaction = client.interactions.create(
            model=GEMINI_MODEL,
            input=prompt,
            response_format={
                "type": "text",
                "mime_type": "application/json",
                "schema": schema,
            },
            generation_config={
                "thinking_level": "medium",
            },
        )

        text = getattr(interaction, "output_text", "") or ""
        text = text.strip()

        if not text:
            return None, "Gemini không trả về output_text."

        result = json.loads(text)
        return result, None

    except Exception as exc:
        return None, f"Gemini phân tích workbook thất bại: {exc}"


# ============================================================
# TẦNG 2 - SCHEMA ENGINE
# ============================================================

RULE_KEYWORDS = {
    "ma_nv": [
        "mã nv",
        "ma nv",
        "mã nhân viên",
        "ma nhan vien",
        "employee id",
        "employee code",
        "staff id",
        "staff code",
        "employee number",
        "id nhân viên",
    ],
    "ho_ten": [
        "họ và tên",
        "ho va ten",
        "họ tên",
        "ho ten",
        "tên nhân viên",
        "ten nhan vien",
        "full name",
        "employee name",
        "name",
    ],
    "ngay_sinh": [
        "ngày sinh",
        "ngay sinh",
        "ngày tháng năm sinh",
        "dob",
        "date of birth",
        "birth date",
        "birthday",
    ],
    "bo_phan": [
        "bộ phận",
        "bo phan",
        "phòng ban",
        "phong ban",
        "đơn vị",
        "don vi",
        "department",
        "division",
        "dept",
    ],
    "ngay_het_han_hop_dong": [
        "ngày hết hạn hợp đồng",
        "ngay het han hop dong",
        "ngày hết hạn hđ",
        "ngay het han hd",
        "ngày kết thúc hợp đồng",
        "ngay ket thuc hop dong",
        "contract end date",
        "contract expiry date",
        "contract expiry",
        "expiry date",
        "contract end",
        "hđ đến",
        "hd den",
    ],
}


def rule_score_column(df, column, field):
    """
    Rule engine độc lập AI.
    Trả score 0..1 dựa trên:
      - tên cột
      - kiểu dữ liệu
      - nội dung mẫu
    """
    norm = normalize_text(column)
    keywords = [normalize_text(x) for x in RULE_KEYWORDS[field]]

    name_score = 0.0

    if norm in keywords:
        name_score = 1.0
    else:
        for keyword in keywords:
            if keyword and keyword in norm:
                name_score = max(name_score, 0.72)

    series = df[column].dropna()
    if series.empty:
        return name_score

    text = series.astype(str).str.strip()

    if field in ("ngay_sinh", "ngay_het_han_hop_dong"):
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
        date_ratio = float(parsed.notna().mean())
        data_score = min(1.0, date_ratio)
    elif field == "ho_ten":
        alpha_ratio = float(
            text.str.contains(r"[A-Za-zÀ-ỹ]", regex=True).mean()
        )
        avg_len = float(text.str.len().mean()) if len(text) else 0
        data_score = min(
            1.0,
            0.6 * alpha_ratio + 0.4 * min(avg_len / 20.0, 1.0),
        )
    elif field == "bo_phan":
        unique_ratio = float(text.nunique() / max(len(text), 1))
        avg_len = float(text.str.len().mean()) if len(text) else 0
        data_score = min(
            1.0,
            0.6 * min(avg_len / 15.0, 1.0)
            + 0.4 * min(unique_ratio * 3, 1.0),
        )
    else:
        # Mã NV thường là chuỗi ngắn, nhiều giá trị khác nhau.
        unique_ratio = float(text.nunique() / max(len(text), 1))
        avg_len = float(text.str.len().mean()) if len(text) else 0
        data_score = min(
            1.0,
            0.55 * min(unique_ratio * 1.5, 1.0)
            + 0.45 * min(avg_len / 12.0, 1.0),
        )

    return round(0.65 * name_score + 0.35 * data_score, 4)


def rule_based_mapping(df):
    scores = {}

    for field in REQUIRED_FIELDS:
        field_scores = []
        for column in df.columns:
            score = rule_score_column(df, column, field)
            field_scores.append((column, score))

        field_scores.sort(key=lambda x: x[1], reverse=True)
        scores[field] = field_scores

    mapping = {}
    confidence = {}

    used = set()

    for field in REQUIRED_FIELDS:
        selected = None
        selected_score = 0.0

        for column, score in scores[field]:
            if column not in used:
                selected = column
                selected_score = score
                break

        mapping[field] = selected
        confidence[field] = selected_score

        if selected is not None:
            used.add(selected)

    return mapping, confidence, scores


def validate_ai_mapping(ai_result, sheets):
    """
    Bảo vệ tầng schema:
    AI chỉ được chọn sheet/cột tồn tại thật.
    """
    if not isinstance(ai_result, dict):
        return None, "AI result không phải object."

    selected_sheet = ai_result.get("selected_sheet")
    mapping = ai_result.get("mapping") or {}

    if selected_sheet not in sheets:
        return None, f"AI chọn sheet không tồn tại: {selected_sheet}"

    actual_columns = set(sheets[selected_sheet]["df"].columns)

    clean = {}
    for field in REQUIRED_FIELDS:
        value = mapping.get(field)

        if value is None:
            clean[field] = None
            continue

        value = str(value).strip()
        if value not in actual_columns:
            clean[field] = None
        else:
            clean[field] = value

    result = {
        "selected_sheet": selected_sheet,
        "mapping": clean,
        "confidence": float(ai_result.get("confidence", 0)),
        "field_confidence": {
            key: float((ai_result.get("field_confidence") or {}).get(key, 0))
            for key in REQUIRED_FIELDS
        },
        "reason": str(ai_result.get("reason", "")),
    }

    return result, None


def merge_ai_and_rules(ai_result, sheets):
    """
    AI là lớp hiểu ngữ nghĩa.
    Rule engine là lớp kiểm chứng độc lập.
    Chỉ tự động chấp nhận khi đủ bằng chứng.
    """
    if not ai_result:
        return None

    checked, error = validate_ai_mapping(ai_result, sheets)
    if checked is None:
        return None

    sheet_name = checked["selected_sheet"]
    df = sheets[sheet_name]["df"]

    rule_mapping, rule_confidence, rule_details = rule_based_mapping(df)

    final_mapping = {}
    final_confidence = {}
    decisions = {}

    for field in REQUIRED_FIELDS:
        ai_col = checked["mapping"].get(field)
        ai_score = checked["field_confidence"].get(field, 0.0)

        rule_col = rule_mapping.get(field)
        rule_score = rule_confidence.get(field, 0.0)

        if ai_col and rule_col == ai_col:
            final_mapping[field] = ai_col
            final_confidence[field] = round(
                0.60 * ai_score + 0.40 * rule_score, 4
            )
            decisions[field] = "AI + Rule đồng thuận"

        elif ai_col and ai_score >= 0.90:
            final_mapping[field] = ai_col
            final_confidence[field] = round(
                0.70 * ai_score + 0.30 * rule_score, 4
            )
            decisions[field] = "AI mạnh, Rule khác"

        elif rule_col and rule_score >= 0.90:
            final_mapping[field] = rule_col
            final_confidence[field] = round(
                0.70 * rule_score + 0.30 * ai_score, 4
            )
            decisions[field] = "Rule mạnh, AI yếu"

        else:
            final_mapping[field] = ai_col or rule_col
            final_confidence[field] = round(
                max(ai_score, rule_score), 4
            )
            decisions[field] = "Cần xác minh"

    return {
        "selected_sheet": sheet_name,
        "mapping": final_mapping,
        "confidence": round(
            sum(final_confidence.values()) / len(REQUIRED_FIELDS), 4
        ),
        "field_confidence": final_confidence,
        "decisions": decisions,
        "ai_reason": checked["reason"],
        "rule_details": rule_details,
    }


def schema_is_safe(schema_result, threshold=0.82):
    """
    Chỉ tự động chạy nếu:
      - đủ 5 trường
      - confidence tổng >= threshold
      - từng trường >= threshold
    """
    if not schema_result:
        return False, "Không có schema."

    mapping = schema_result.get("mapping", {})
    confidence = schema_result.get("field_confidence", {})

    missing = [f for f in REQUIRED_FIELDS if not mapping.get(f)]
    if missing:
        return False, "Thiếu: " + ", ".join(FIELD_LABELS[x] for x in missing)

    weak = [
        FIELD_LABELS[f]
        for f in REQUIRED_FIELDS
        if float(confidence.get(f, 0)) < threshold
    ]

    if weak:
        return False, (
            "Độ tin cậy chưa đủ cho: "
            + ", ".join(weak)
            + f". Ngưỡng tự động: {threshold:.0%}."
        )

    return True, "Schema đạt ngưỡng tự động."


def prepare_standard_dataframe(df, mapping):
    """
    Từ Excel bất kỳ -> DataFrame chuẩn.
    """
    result = pd.DataFrame(index=df.index)

    result["Mã NV"] = df[mapping["ma_nv"]].astype("string").str.strip()
    result["Họ và tên"] = df[mapping["ho_ten"]].astype("string").str.strip()

    result["Ngày tháng năm sinh"] = pd.to_datetime(
        df[mapping["ngay_sinh"]],
        errors="coerce",
        dayfirst=True,
    )

    result["Bộ Phận"] = df[mapping["bo_phan"]].astype("string").str.strip()

    result["Ngày hết hạn hợp đồng"] = pd.to_datetime(
        df[mapping["ngay_het_han_hop_dong"]],
        errors="coerce",
        dayfirst=True,
    )

    # Xóa dòng hoàn toàn không có tên.
    result = result[
        result["Họ và tên"].notna()
        & (result["Họ và tên"].astype(str).str.strip() != "")
    ].copy()

    return result.reset_index(drop=True)


# ============================================================
# TẦNG 3 - VALIDATION ENGINE
# ============================================================

def validate_standard_dataframe(df):
    checks = []

    total = len(df)
    if total == 0:
        return {
            "ok": False,
            "score": 0.0,
            "checks": [{
                "name": "Số dòng",
                "ok": False,
                "detail": "Không có dữ liệu nhân viên."
            }],
        }

    def add(name, ok, detail):
        checks.append({
            "name": name,
            "ok": bool(ok),
            "detail": detail,
        })

    # Tên.
    name_valid = (
        df["Họ và tên"].notna()
        & (df["Họ và tên"].astype(str).str.strip() != "")
    )
    name_ratio = float(name_valid.mean())
    add(
        "Họ và tên",
        name_ratio >= 0.98,
        f"{name_ratio:.1%} dòng có họ tên.",
    )

    # Mã NV.
    id_valid = (
        df["Mã NV"].notna()
        & (df["Mã NV"].astype(str).str.strip() != "")
    )
    id_ratio = float(id_valid.mean())
    add(
        "Mã nhân viên",
        id_ratio >= 0.90,
        f"{id_ratio:.1%} dòng có mã nhân viên.",
    )

    # Ngày sinh.
    dob_ratio = float(df["Ngày tháng năm sinh"].notna().mean())
    add(
        "Ngày sinh",
        dob_ratio >= 0.85,
        f"{dob_ratio:.1%} ngày sinh đọc được.",
    )

    # Bộ phận.
    dept_ratio = float(df["Bộ Phận"].notna().mean())
    add(
        "Bộ phận",
        dept_ratio >= 0.85,
        f"{dept_ratio:.1%} dòng có bộ phận.",
    )

    # Hợp đồng.
    contract_ratio = float(
        df["Ngày hết hạn hợp đồng"].notna().mean()
    )
    add(
        "Ngày hết hạn hợp đồng",
        contract_ratio >= 0.70,
        f"{contract_ratio:.1%} ngày hết hạn đọc được.",
    )

    # Ngày sinh hợp lý.
    current_year = datetime.now().year
    dob = df["Ngày tháng năm sinh"]
    dob_reasonable = dob.isna() | (
        (dob.dt.year >= 1900)
        & (dob.dt.year <= current_year)
    )
    reasonable_ratio = float(dob_reasonable.mean())
    add(
        "Ngày sinh hợp lý",
        reasonable_ratio >= 0.98,
        f"{reasonable_ratio:.1%} giá trị nằm trong khoảng hợp lý.",
    )

    # Trùng mã NV.
    ids = df["Mã NV"].dropna().astype(str).str.strip()
    duplicate_count = int(ids.duplicated(keep=False).sum())
    duplicate_ratio = duplicate_count / max(total, 1)
    add(
        "Trùng mã nhân viên",
        duplicate_ratio <= 0.05,
        f"{duplicate_count} dòng thuộc nhóm mã bị trùng.",
    )

    passed = sum(1 for x in checks if x["ok"])
    score = passed / len(checks)

    # Các kiểm tra lõi phải đạt.
    core_ok = (
        name_ratio >= 0.98
        and id_ratio >= 0.90
        and dob_ratio >= 0.85
        and dept_ratio >= 0.85
        and contract_ratio >= 0.70
        and reasonable_ratio >= 0.98
        and duplicate_ratio <= 0.05
    )

    return {
        "ok": core_ok,
        "score": round(score, 4),
        "checks": checks,
    }


# ============================================================
# LƯU / TẢI SCHEMA
# ============================================================

def save_schema(schema_result, source_hash):
    payload = {
        "source_hash": source_hash,
        "saved_at": datetime.now().isoformat(),
        "schema": schema_result,
    }
    SCHEMA_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    SOURCE_HASH_FILE.write_text(source_hash, encoding="utf-8")


def load_saved_schema(source_hash):
    if not SCHEMA_FILE.exists():
        return None

    try:
        payload = json.loads(
            SCHEMA_FILE.read_text(encoding="utf-8")
        )
        if payload.get("source_hash") != source_hash:
            return None
        return payload.get("schema")
    except Exception:
        return None


# ============================================================
# TẦNG 4 - BUSINESS ENGINE
# ============================================================

def get_birthday_report(df, month):
    result = df[
        df["Ngày tháng năm sinh"].notna()
        & (df["Ngày tháng năm sinh"].dt.month == month)
    ].copy()

    columns = [
        "Mã NV",
        "Họ và tên",
        "Ngày tháng năm sinh",
        "Bộ Phận",
    ]
    return result[columns].sort_values(
        by="Ngày tháng năm sinh"
    )


def get_contract_report(df, month, year):
    result = df[
        df["Ngày hết hạn hợp đồng"].notna()
        & (df["Ngày hết hạn hợp đồng"].dt.month == month)
        & (df["Ngày hết hạn hợp đồng"].dt.year == year)
    ].copy()

    columns = [
        "Mã NV",
        "Họ và tên",
        "Ngày tháng năm sinh",
        "Bộ Phận",
        "Ngày hết hạn hợp đồng",
    ]
    return result[columns].sort_values(
        by="Ngày hết hạn hợp đồng"
    )


def get_upcoming_contract_report(df, days=30):
    today = pd.Timestamp(date.today())
    end = today + pd.Timedelta(days=days)

    result = df[
        df["Ngày hết hạn hợp đồng"].notna()
        & (df["Ngày hết hạn hợp đồng"] >= today)
        & (df["Ngày hết hạn hợp đồng"] <= end)
    ].copy()

    result["Còn lại (ngày)"] = (
        result["Ngày hết hạn hợp đồng"] - today
    ).dt.days

    columns = [
        "Mã NV",
        "Họ và tên",
        "Bộ Phận",
        "Ngày hết hạn hợp đồng",
        "Còn lại (ngày)",
    ]

    return result[columns].sort_values(
        by="Ngày hết hạn hợp đồng"
    )


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


def save_excel_file(df, filename, sheet_name):
    path = REPORT_DIR / filename
    output = dataframe_to_excel(df, sheet_name)

    with open(path, "wb") as f:
        f.write(output.getvalue())

    return path


# ============================================================
# TẦNG 5 - AUTOMATION ENGINE / TELEGRAM
# ============================================================

def telegram_request(method, data=None, files=None, timeout=30):
    if not TELEGRAM_BOT_TOKEN:
        return False, "Thiếu TELEGRAM_BOT_TOKEN."

    url = (
        f"https://api.telegram.org/"
        f"bot{TELEGRAM_BOT_TOKEN}/{method}"
    )

    try:
        response = requests.post(
            url,
            data=data,
            files=files,
            timeout=timeout,
        )

        if response.ok:
            payload = response.json()
            if payload.get("ok"):
                return True, payload

        return False, response.text

    except Exception as exc:
        return False, str(exc)


def telegram_send_message(message):
    if not TELEGRAM_CHAT_ID:
        return False, "Thiếu TELEGRAM_CHAT_ID."

    return telegram_request(
        "sendMessage",
        data={
            "chat_id": TELEGRAM_CHAT_ID,
            "text": message,
        },
        timeout=30,
    )


def telegram_send_file(file_path, caption):
    if not TELEGRAM_CHAT_ID:
        return False, "Thiếu TELEGRAM_CHAT_ID."

    try:
        with open(file_path, "rb") as f:
            return telegram_request(
                "sendDocument",
                data={
                    "chat_id": TELEGRAM_CHAT_ID,
                    "caption": caption,
                },
                files={
                    "document": f,
                },
                timeout=60,
            )
    except Exception as exc:
        return False, str(exc)


def telegram_test():
    ok, result = telegram_request(
        "getMe",
        data={},
        timeout=15,
    )

    if not ok:
        return False, result

    username = (
        result.get("result", {}).get("username", "")
        if isinstance(result, dict)
        else ""
    )

    return True, username


def send_monthly_report(
    birthday_df,
    contract_df,
    month,
    year,
    source_name="Excel nhân sự",
):
    birthday_file = save_excel_file(
        birthday_df,
        f"danh_sach_sinh_nhat_{month:02d}_{year}.xlsx",
        "Sinh nhật",
    )

    contract_file = save_excel_file(
        contract_df,
        f"danh_sach_den_han_ky_hop_dong_{month:02d}_{year}.xlsx",
        "Hợp đồng",
    )

    message = f"""
📊 BÁO CÁO NHÂN SỰ THÁNG {month:02d}/{year}

📁 Nguồn dữ liệu: {source_name}

🎂 Sinh nhật: {len(birthday_df)} nhân viên
📄 Hợp đồng hết hạn: {len(contract_df)} nhân viên

⏰ Thời gian: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
""".strip()

    ok, error = telegram_send_message(message)
    if not ok:
        return False, f"Gửi message thất bại: {error}"

    ok, error = telegram_send_file(
        birthday_file,
        f"🎂 Danh sách sinh nhật {month:02d}/{year}",
    )
    if not ok:
        return False, f"Gửi file sinh nhật thất bại: {error}"

    ok, error = telegram_send_file(
        contract_file,
        f"📄 Danh sách hợp đồng {month:02d}/{year}",
    )
    if not ok:
        return False, f"Gửi file hợp đồng thất bại: {error}"

    return True, "Đã gửi thành công."


def load_automation_log():
    if not AUTOMATION_LOG_FILE.exists():
        return {}

    try:
        return json.loads(
            AUTOMATION_LOG_FILE.read_text(
                encoding="utf-8"
            )
        )
    except Exception:
        return {}


def save_automation_log(log):
    AUTOMATION_LOG_FILE.write_text(
        json.dumps(
            log,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def automation_key(year, month, source_hash):
    return f"{year:04d}-{month:02d}:{source_hash}"


def already_sent(year, month, source_hash):
    log = load_automation_log()
    key = automation_key(
        year,
        month,
        source_hash,
    )
    return key in log


def mark_sent(year, month, source_hash, detail=""):
    log = load_automation_log()

    key = automation_key(
        year,
        month,
        source_hash,
    )

    log[key] = {
        "sent_at": datetime.now().isoformat(),
        "detail": detail,
    }

    save_automation_log(log)


def run_automation(
    prepared_df,
    source_hash,
    source_name,
    force=False,
):
    """
    Luồng cron thực sự:
      file đã lưu
      -> schema
      -> validation
      -> business report
      -> Telegram
      -> mark sent

    Không phụ thuộc vào Dashboard.
    """
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False, "Chưa cấu hình Telegram."

    now = datetime.now()
    month = now.month
    year = now.year

    # Mặc định chỉ gửi vào ngày 01 từ 10:00.
    if not force:
        if now.day != 1 or now.hour < 10:
            return False, "Chưa đến thời điểm gửi tự động."

    if already_sent(year, month, source_hash):
        return True, "Báo cáo tháng này đã được gửi."

    validation = validate_standard_dataframe(prepared_df)
    if not validation["ok"]:
        return False, "Dữ liệu chưa đạt validation."

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
        source_name=source_name,
    )

    if success:
        mark_sent(
            year,
            month,
            source_hash,
            detail=message,
        )

    return success, message


# ============================================================
# CRON ENTRY POINT
#
# KHÔNG st.stop() trước khi chạy automation.
# ?cron=1 chỉ chạy automation rồi kết thúc.
# ============================================================

def cron_entrypoint():
    if "cron" not in st.query_params:
        return False

    if not CURRENT_DB_FILE.exists():
        st.error("Cron: chưa có file Excel nhân sự.")
        return True

    df_hash = workbook_fingerprint(CURRENT_DB_FILE)

    sheets, error = read_workbook(CURRENT_DB_FILE)
    if error:
        st.error(f"Cron: không đọc được Excel: {error}")
        return True

    saved_schema = load_saved_schema(df_hash)

    if not saved_schema:
        st.error(
            "Cron: chưa có schema AI đã xác minh cho file hiện tại. "
            "Hãy mở Dashboard và phân tích file trước."
        )
        return True

    safe, reason = schema_is_safe(saved_schema)
    if not safe:
        st.error(f"Cron: schema không an toàn: {reason}")
        return True

    sheet_name = saved_schema["selected_sheet"]
    mapping = saved_schema["mapping"]

    if sheet_name not in sheets:
        st.error("Cron: sheet đã lưu không còn tồn tại.")
        return True

    prepared = prepare_standard_dataframe(
        sheets[sheet_name]["df"],
        mapping,
    )

    validation = validate_standard_dataframe(prepared)

    if not validation["ok"]:
        st.error(
            "Cron: validation thất bại. "
            "Không gửi Telegram để tránh báo cáo sai."
        )
        return True

    success, message = run_automation(
        prepared_df=prepared,
        source_hash=df_hash,
        source_name=CURRENT_DB_FILE.name,
        force=False,
    )

    if success:
        st.success(f"CRON OK: {message}")
    else:
        st.warning(f"CRON: {message}")

    return True


# ============================================================
# APP START
# ============================================================

st.markdown(
    '<div class="main-title">📊 QUẢN LÝ NHÂN SỰ AI</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">'
    'Upload Excel bất kỳ → AI hiểu cấu trúc → kiểm tra → Dashboard → Telegram → tự động'
    '</div>',
    unsafe_allow_html=True,
)


# Cron phải được xử lý SAU khi các hàm đã được định nghĩa.
if cron_entrypoint():
    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Hệ thống")

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
        format_func=lambda x: f"Tháng {x:02d}",
    )

    st.divider()

    st.subheader("🤖 Gemini")
    if GEMINI_API_KEY:
        st.success(
            f"🟢 API đã cấu hình\n\n"
            f"Model: `{GEMINI_MODEL}`"
        )
    else:
        st.warning("🔴 Chưa có GEMINI_API_KEY")

    st.subheader("📱 Telegram")
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        st.success("🟢 Telegram đã cấu hình")
    else:
        st.warning("🔴 Chưa cấu hình Telegram")


# ============================================================
# 1. UPLOAD
# ============================================================

st.header("1️⃣ Upload file Excel nhân sự")

uploaded_file = st.file_uploader(
    "Chọn file Excel (.xlsx hoặc .xls)",
    type=["xlsx", "xls"],
    help=(
        "Không yêu cầu tên cột cố định. "
        "Hệ thống sẽ đọc sheet, header và dữ liệu mẫu."
    ),
)

if uploaded_file is not None:
    try:
        with open(CURRENT_DB_FILE, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # File mới phải được AI phân tích lại.
        if SCHEMA_FILE.exists():
            SCHEMA_FILE.unlink()

        if SOURCE_HASH_FILE.exists():
            SOURCE_HASH_FILE.unlink()

        st.session_state.pop("analysis_result", None)
        st.session_state.pop("prepared_df", None)

        st.success(
            f"✅ Đã lưu file: {uploaded_file.name}"
        )

    except Exception as exc:
        st.error(f"Không thể lưu file: {exc}")


if not CURRENT_DB_FILE.exists():
    st.info(
        "📁 Chưa có dữ liệu. Hãy upload file Excel nhân sự."
    )
    st.stop()


# ============================================================
# ĐỌC WORKBOOK
# ============================================================

sheets, workbook_error = read_workbook(
    CURRENT_DB_FILE
)

if workbook_error:
    st.error(
        f"Không thể đọc workbook: {workbook_error}"
    )
    st.stop()

if not sheets:
    st.error("Workbook không có sheet dữ liệu hợp lệ.")
    st.stop()

source_hash = workbook_fingerprint(
    CURRENT_DB_FILE
)

st.success(
    f"📁 File hiện tại: **{CURRENT_DB_FILE.name}** | "
    f"Sheets: **{len(sheets)}**"
)


# ============================================================
# 2. EXCEL INTELLIGENCE + SCHEMA ENGINE
# ============================================================

st.header("2️⃣ AI hiểu cấu trúc Excel")

sheet_profiles = build_sheet_intelligence(
    sheets
)

with st.expander("🔎 Xem AI sẽ phân tích gì"):
    st.json(sheet_profiles)


if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = None


if st.button(
    "🤖 PHÂN TÍCH TOÀN BỘ EXCEL BẰNG AI",
    type="primary",
    use_container_width=True,
):
    if not GEMINI_API_KEY:
        st.error(
            "Bạn chưa cấu hình GEMINI_API_KEY."
        )
    else:
        with st.spinner(
            "AI đang đọc sheet + cột + dữ liệu mẫu..."
        ):
            ai_result, ai_error = ai_analyze_workbook(
                sheet_profiles
            )

            if ai_error:
                st.error(ai_error)
                st.session_state.analysis_result = None
            else:
                merged = merge_ai_and_rules(
                    ai_result,
                    sheets,
                )

                if merged is None:
                    st.error(
                        "Không thể tạo schema an toàn từ AI."
                    )
                    st.session_state.analysis_result = None
                else:
                    st.session_state.analysis_result = merged

                    # Chỉ lưu schema nếu đủ an toàn.
                    safe, reason = schema_is_safe(
                        merged
                    )

                    if safe:
                        save_schema(
                            merged,
                            source_hash,
                        )
                        st.success(
                            "✅ Schema đạt ngưỡng và đã được lưu "
                            "cho Dashboard + Cron."
                        )
                    else:
                        st.warning(
                            "⚠️ Schema chưa đạt ngưỡng tự động: "
                            + reason
                        )


schema_result = st.session_state.analysis_result

# Nếu chưa phân tích trong phiên hiện tại, thử lấy schema đã lưu.
if schema_result is None:
    schema_result = load_saved_schema(
        source_hash
    )


if schema_result is None:
    st.info(
        "Hãy nhấn **PHÂN TÍCH TOÀN BỘ EXCEL BẰNG AI** "
        "để hệ thống hiểu file."
    )
    st.stop()


# ============================================================
# HIỂN THỊ SCHEMA
# ============================================================

st.subheader("🧠 Schema nhân sự chuẩn")

schema_rows = []

for field in REQUIRED_FIELDS:
    schema_rows.append({
        "Trường chuẩn": FIELD_LABELS[field],
        "Cột Excel": schema_result["mapping"].get(field),
        "Độ tin cậy": (
            f'{schema_result["field_confidence"].get(field, 0):.1%}'
        ),
        "Quyết định": schema_result.get(
            "decisions",
            {},
        ).get(field, ""),
    })

schema_df = pd.DataFrame(schema_rows)

st.dataframe(
    schema_df,
    use_container_width=True,
    hide_index=True,
)

st.caption(
    f'Độ tin cậy tổng: '
    f'{schema_result.get("confidence", 0):.1%}'
)

if schema_result.get("ai_reason"):
    st.info(
        "AI: " + schema_result["ai_reason"]
    )


# ============================================================
# CHỌN SHEET + CHUẨN HÓA
# ============================================================

selected_sheet = schema_result["selected_sheet"]

if selected_sheet not in sheets:
    st.error(
        "Sheet trong schema không tồn tại trong file hiện tại."
    )
    st.stop()

source_df = sheets[selected_sheet]["df"]

prepared_df = prepare_standard_dataframe(
    source_df,
    schema_result["mapping"],
)


# ============================================================
# 3. VALIDATION ENGINE
# ============================================================

st.header("3️⃣ Kiểm tra độ chính xác dữ liệu")

validation = validate_standard_dataframe(
    prepared_df
)

v1, v2, v3 = st.columns(3)

with v1:
    if validation["ok"]:
        st.success("🟢 VALIDATION ĐẠT")
    else:
        st.error("🔴 VALIDATION KHÔNG ĐẠT")

with v2:
    st.metric(
        "Điểm validation",
        f'{validation["score"]:.1%}',
    )

with v3:
    st.metric(
        "Số nhân sự",
        f"{len(prepared_df):,}",
    )

validation_df = pd.DataFrame(
    validation["checks"]
)

st.dataframe(
    validation_df,
    use_container_width=True,
    hide_index=True,
)


if not validation["ok"]:
    st.error(
        "Hệ thống CHẶN Dashboard/Telegram tự động "
        "vì dữ liệu chưa đạt validation."
    )

    with st.expander(
        "👁️ Xem dữ liệu chuẩn hóa để sửa mapping"
    ):
        st.dataframe(
            prepared_df,
            use_container_width=True,
            height=350,
        )

    st.stop()


# ============================================================
# 4. BUSINESS ENGINE / DASHBOARD
# ============================================================

st.header(
    f"4️⃣ Dashboard tháng "
    f"{selected_month:02d}/{selected_year}"
)

birthday_df = get_birthday_report(
    prepared_df,
    selected_month,
)

contract_df = get_contract_report(
    prepared_df,
    selected_month,
    selected_year,
)

upcoming_df = get_upcoming_contract_report(
    prepared_df,
    days=30,
)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "👥 Tổng nhân sự",
        f"{len(prepared_df):,}",
    )

with c2:
    st.metric(
        "🎂 Sinh nhật",
        f"{len(birthday_df):,}",
    )

with c3:
    st.metric(
        "📄 HĐ hết hạn",
        f"{len(contract_df):,}",
    )

with c4:
    st.metric(
        "⚠️ HĐ trong 30 ngày",
        f"{len(upcoming_df):,}",
    )


# Sinh nhật
st.subheader(
    f"🎂 Sinh nhật tháng {selected_month:02d}"
)

if birthday_df.empty:
    st.info(
        "Không có nhân viên sinh nhật trong tháng."
    )
else:
    display = birthday_df.copy()
    display["Ngày tháng năm sinh"] = (
        display["Ngày tháng năm sinh"]
        .dt.strftime("%d/%m/%Y")
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
    )

st.download_button(
    "⬇️ Download danh sách sinh nhật",
    data=dataframe_to_excel(
        birthday_df,
        "Sinh nhật",
    ).getvalue(),
    file_name=(
        f"danh_sach_sinh_nhat_"
        f"{selected_month:02d}_{selected_year}.xlsx"
    ),
)


# Hợp đồng
st.subheader(
    f"📄 Hợp đồng hết hạn "
    f"tháng {selected_month:02d}/{selected_year}"
)

if contract_df.empty:
    st.info(
        "Không có hợp đồng hết hạn trong tháng."
    )
else:
    display = contract_df.copy()

    display["Ngày tháng năm sinh"] = (
        display["Ngày tháng năm sinh"]
        .dt.strftime("%d/%m/%Y")
    )

    display["Ngày hết hạn hợp đồng"] = (
        display["Ngày hết hạn hợp đồng"]
        .dt.strftime("%d/%m/%Y")
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
    )

st.download_button(
    "⬇️ Download danh sách hợp đồng",
    data=dataframe_to_excel(
        contract_df,
        "Hợp đồng",
    ).getvalue(),
    file_name=(
        f"danh_sach_den_han_ky_hop_dong_"
        f"{selected_month:02d}_{selected_year}.xlsx"
    ),
)


# HĐ sắp hết hạn
with st.expander(
    "⚠️ Hợp đồng sẽ hết hạn trong 30 ngày"
):
    if upcoming_df.empty:
        st.info(
            "Không có hợp đồng hết hạn trong 30 ngày tới."
        )
    else:
        display = upcoming_df.copy()
        display["Ngày hết hạn hợp đồng"] = (
            display["Ngày hết hạn hợp đồng"]
            .dt.strftime("%d/%m/%Y")
        )

        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# 5. TELEGRAM + AUTOMATION
# ============================================================

st.divider()
st.header("5️⃣ Telegram & tự động hàng tháng")

t1, t2 = st.columns(2)

with t1:
    st.subheader("📱 Trạng thái Telegram")

    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        ok, bot_name = telegram_test()

        if ok:
            st.success(
                f"🟢 Bot hoạt động: @{bot_name}"
            )
        else:
            st.error(
                f"🔴 Telegram lỗi: {bot_name}"
            )
    else:
        st.warning(
            "Chưa cấu hình TELEGRAM_BOT_TOKEN / "
            "TELEGRAM_CHAT_ID."
        )


with t2:
    st.subheader("📤 Gửi thủ công")

    if st.button(
        "📤 GỬI BÁO CÁO THÁNG NÀY",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner(
            "Đang tạo và gửi báo cáo..."
        ):
            success, message = send_monthly_report(
                birthday_df,
                contract_df,
                selected_month,
                selected_year,
                source_name=CURRENT_DB_FILE.name,
            )

        if success:
            st.success(
                "✅ " + message
            )
        else:
            st.error(
                "❌ " + message
            )


st.subheader("🤖 Cron tự động")

st.info(
    "Cron nên gọi URL của ứng dụng với `?cron=1` "
    "vào ngày 01 hàng tháng từ 10:00 trở đi. "
    "Cron sẽ đọc file + schema đã xác minh + validation "
    "rồi mới gửi Telegram. Nếu đã gửi, hệ thống không gửi trùng."
)

now = datetime.now()

st.write(
    f"🕐 Máy chủ: "
    f"**{now.strftime('%d/%m/%Y %H:%M:%S')}**"
)

st.write(
    f"📁 File hash: `{source_hash[:16]}...`"
)

st.write(
    f"🧠 Sheet đang dùng: **{selected_sheet}**"
)

st.write(
    f"🎯 Confidence schema: "
    f"**{schema_result.get('confidence', 0):.1%}**"
)

log = load_automation_log()

if log:
    latest_key = list(log.keys())[-1]
    latest = log[latest_key]

    st.write(
        f"📌 Lần gửi gần nhất: "
        f"**{latest.get('sent_at', '')}**"
    )
else:
    st.write(
        "📌 Chưa có lịch sử gửi tự động."
    )


# Nút test cron.
if st.button(
    "🧪 TEST LUỒNG AUTOMATION NGAY",
    use_container_width=True,
):
    with st.spinner(
        "Đang test toàn bộ tầng Automation..."
    ):
        success, message = run_automation(
            prepared_df=prepared_df,
            source_hash=source_hash,
            source_name=CURRENT_DB_FILE.name,
            force=True,
        )

    if success:
        st.success(
            "🟢 Automation: " + message
        )
    else:
        st.error(
            "🔴 Automation: " + message
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "5 tầng: Excel Intelligence → Schema Engine → "
    "Validation Engine → Business Engine → Automation Engine"
)
