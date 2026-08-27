import io
import re
import math
import urllib.request
from pathlib import Path
from functools import lru_cache

import streamlit as st
import openpyxl
import numpy as np

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import argostranslate.package
import argostranslate.translate

from PIL import Image, ImageEnhance, ImageFilter, ImageOps

try:
    import easyocr
except Exception:
    easyocr = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None


# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Dịch Bảng Chấm Công Chuyên Sâu",
    page_icon="🌐",
    layout="wide"
)

st.title("🌐 Dịch & Xuất Bảng Chấm Công Song Ngữ")
st.caption(
    "OCR nâng cao + Glossary chuyên ngành + Translation Memory + Argos Translate offline"
)


# ============================================================
# BỘ TỪ ĐIỂN CHUYÊN NGÀNH
# ============================================================

GLOSSARY_ZH_VI = {
    # --------------------------------------------------------
    # NHÂN SỰ / CHẤM CÔNG
    # --------------------------------------------------------
    "正式工": "Công nhân chính thức",
    "临时工": "Công nhân thời vụ",
    "合同工": "Công nhân hợp đồng",
    "员工": "Nhân viên",
    "工人": "Công nhân",
    "职员": "Nhân viên",
    "部门": "Bộ phận",
    "姓名": "Họ tên",
    "名字": "Tên",
    "工号": "Mã nhân viên",
    "编号": "Mã số",
    "班次": "Ca làm việc",
    "早班": "Ca sáng",
    "中班": "Ca chiều",
    "晚班": "Ca đêm",
    "白班": "Ca ngày",
    "夜班": "Ca đêm",

    # --------------------------------------------------------
    # CHẤM CÔNG
    # --------------------------------------------------------
    "出勤": "Đi làm",
    "出勤天数": "Số ngày đi làm",
    "考勤": "Chấm công",
    "考勤表": "Bảng chấm công",
    "打卡": "Chấm công",
    "上班": "Đi làm",
    "下班": "Tan ca",
    "休息": "Nghỉ",
    "休假": "Nghỉ phép",
    "请假": "Xin nghỉ",
    "病假": "Nghỉ bệnh",
    "事假": "Nghỉ việc riêng",
    "年假": "Nghỉ phép năm",
    "旷工": "Nghỉ không phép",
    "迟到": "Đi muộn",
    "早退": "Về sớm",
    "加班": "Tăng ca",
    "加班时间": "Thời gian tăng ca",
    "加班小时": "Số giờ tăng ca",
    "正常工时": "Giờ làm việc bình thường",
    "工作时间": "Thời gian làm việc",

    # --------------------------------------------------------
    # XƯỞNG / BỘ PHẬN
    # --------------------------------------------------------
    "车间": "Xưởng sản xuất",
    "生产车间": "Xưởng sản xuất",
    "一车间": "Xưởng 1",
    "二车间": "Xưởng 2",
    "三车间": "Xưởng 3",
    "四车间": "Xưởng 4",
    "1车间": "Xưởng 1",
    "2车间": "Xưởng 2",
    "3车间": "Xưởng 3",
    "4车间": "Xưởng 4",

    "品管部": "Phòng Quản lý chất lượng (QC)",
    "品质部": "Phòng Quản lý chất lượng (QC)",
    "质量部": "Phòng Quản lý chất lượng (QC)",
    "货仓": "Kho hàng",
    "仓库": "Kho hàng",
    "仓储部": "Bộ phận kho",
    "包装部": "Bộ phận đóng gói",
    "裁断部": "Bộ phận cắt",
    "生产部": "Bộ phận sản xuất",
    "采购部": "Bộ phận mua hàng",
    "销售部": "Bộ phận kinh doanh",
    "人事部": "Bộ phận nhân sự",
    "行政部": "Bộ phận hành chính",
    "财务部": "Bộ phận tài chính",
    "工程部": "Bộ phận kỹ thuật",
    "技术部": "Bộ phận kỹ thuật",
    "维修部": "Bộ phận bảo trì",
    "设备部": "Bộ phận thiết bị",

    # --------------------------------------------------------
    # MÁY MÓC
    # --------------------------------------------------------
    "开几台机": "Số máy chạy",
    "开机": "Chạy máy",
    "机器": "Máy",
    "机械": "Máy móc",
    "设备": "Thiết bị",
    "台": "máy",
    "机器数": "Số máy",
    "设备数量": "Số lượng thiết bị",

    # --------------------------------------------------------
    # BẢNG / TỔNG KẾT
    # --------------------------------------------------------
    "备注": "Ghi chú",
    "说明": "Ghi chú",
    "总计": "Tổng cộng",
    "合计": "Tổng cộng",
    "小计": "Tổng phụ",
    "总人数": "Tổng số người",
    "人数": "Số người",
    "总数": "Tổng số",

    # --------------------------------------------------------
    # NGÀY THÁNG
    # --------------------------------------------------------
    "年": "năm",
    "月": "tháng",
    "日": "ngày",
    "日期": "Ngày",
    "日期：": "Ngày:",
}


# ============================================================
# TRANSLATION MEMORY
# ============================================================

TRANSLATION_MEMORY_ZH_VI = {
    "品管部": "Phòng Quản lý chất lượng (QC)",
    "品质部": "Phòng Quản lý chất lượng (QC)",
    "裁断部": "Bộ phận cắt",
    "包装部": "Bộ phận đóng gói",
    "货仓": "Kho hàng",
    "正式工": "Công nhân chính thức",
    "临时工": "Công nhân thời vụ",
    "开几台机": "Số máy chạy",
    "考勤表": "Bảng chấm công",
    "加班": "Tăng ca",
    "旷工": "Nghỉ không phép",
    "迟到": "Đi muộn",
    "早退": "Về sớm",
}


# ============================================================
# GLOSSARY VIỆT -> TRUNG
# ============================================================

GLOSSARY_VI_ZH = {
    v: k
    for k, v in GLOSSARY_ZH_VI.items()
}

TRANSLATION_MEMORY_VI_ZH = {
    v: k
    for k, v in TRANSLATION_MEMORY_ZH_VI.items()
}


# ============================================================
# ARGOS MODEL
# ============================================================

ARGOS_PACKAGES = {
    "zh_en": "translate-zh_en-1_9.argosmodel",
    "en_zh": "translate-en_zh-1_9.argosmodel",
    "vi_en": "translate-vi_en-1_9.argosmodel",
    "en_vi": "translate-en_vi-1_9.argosmodel",
}

ARGOS_BASE_URL = (
    "https://data.argosopentech.com/argospm/v1/"
)


# ============================================================
# TEXT NORMALIZATION
# ============================================================

def normalize_text(text):
    if text is None:
        return ""

    text = str(text)

    text = (
        text.replace("\r\n", "\n")
        .replace("\r", "\n")
        .replace("\u3000", " ")
    )

    text = re.sub(r"[ \t]+", " ", text)

    return text.strip()


def normalize_chinese(text):
    text = normalize_text(text)

    replacements = {
        "，": ",",
        "。": ".",
        "：": ":",
        "；": ";",
        "（": "(",
        "）": ")",
        "【": "[",
        "】": "]",
    }

    for a, b in replacements.items():
        text = text.replace(a, b)

    return text.strip()


# ============================================================
# PHÁT HIỆN DỮ LIỆU KHÔNG CẦN DỊCH
# ============================================================

def is_number(text):
    text = normalize_text(text)

    if not text:
        return False

    return bool(
        re.fullmatch(
            r"[-+]?\d+(?:[.,]\d+)?%?",
            text
        )
    )


def is_date(text):
    text = normalize_text(text)

    patterns = [
        r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$",
        r"^\d{1,2}[-/.]\d{1,2}[-/.]\d{4}$",
        r"^\d{4}年\d{1,2}月\d{1,2}日?$",
        r"^\d{1,2}月\d{1,2}日$",
    ]

    return any(re.fullmatch(p, text) for p in patterns)


def is_code(text):
    text = normalize_text(text)

    if not text:
        return False

    if len(text) > 30:
        return False

    # Mã kiểu A001 / A-001 / EMP001 / AB12
    if re.fullmatch(
        r"[A-Za-z]{1,8}[-_]?\d{1,10}",
        text
    ):
        return True

    # Mã chỉ gồm số và ký tự đặc biệt
    if re.fullmatch(
        r"[\dA-Za-z_\-/\.]+",
        text
    ):
        if any(c.isdigit() for c in text):
            return True

    return False


def should_preserve(text):
    text = normalize_text(text)

    if not text:
        return True

    if is_number(text):
        return True

    if is_date(text):
        return True

    if is_code(text):
        return True

    return False


# ============================================================
# ARGOS LANGUAGE CACHE
# ============================================================

@st.cache_resource
def get_argos_languages():
    try:
        return argostranslate.translate.get_installed_languages()
    except Exception:
        return []


def find_argos_translation(from_code, to_code):

    languages = get_argos_languages()

    source = None
    target = None

    for lang in languages:
        if lang.code == from_code:
            source = lang

        if lang.code == to_code:
            target = lang

    if source is None or target is None:
        return None

    try:
        return source.get_translation(target)
    except Exception:
        return None


# ============================================================
# DOWNLOAD MODEL
# ============================================================

def download_argos_package(package_name):

    cache_dir = Path(".argos_models")
    cache_dir.mkdir(exist_ok=True)

    target = cache_dir / package_name

    if target.exists() and target.stat().st_size > 0:
        return target

    try:

        url = ARGOS_BASE_URL + package_name

        urllib.request.urlretrieve(
            url,
            str(target)
        )

        return target

    except Exception as e:

        if target.exists():
            try:
                target.unlink()
            except Exception:
                pass

        raise RuntimeError(
            f"Không thể tải model Argos: {e}"
        )


def install_argos_package(from_code, to_code):

    key = f"{from_code}_{to_code}"

    if key not in ARGOS_PACKAGES:
        return

    package_name = ARGOS_PACKAGES[key]

    path = download_argos_package(package_name)

    try:
        argostranslate.package.install_from_path(
            str(path)
        )
    except Exception:
        pass

    try:
        get_argos_languages.clear()
    except Exception:
        pass


# ============================================================
# KHỞI TẠO MODEL
# ============================================================

@st.cache_resource
def initialize_translation_models():

    pairs = [
        ("zh", "en"),
        ("en", "zh"),
        ("vi", "en"),
        ("en", "vi"),
    ]

    installed = []

    for from_code, to_code in pairs:

        if find_argos_translation(
            from_code,
            to_code
        ) is None:

            try:

                install_argos_package(
                    from_code,
                    to_code
                )

                installed.append(
                    f"{from_code}->{to_code}"
                )

            except Exception:
                pass

    return installed


# ============================================================
# ARGOS DIRECT TRANSLATION
# ============================================================

def translate_direct(
    text,
    from_code,
    to_code
):

    text = normalize_text(text)

    if not text:
        return ""

    if from_code == to_code:
        return text

    tr = find_argos_translation(
        from_code,
        to_code
    )

    if tr is None:

        try:
            install_argos_package(
                from_code,
                to_code
            )
        except Exception:
            return text

        tr = find_argos_translation(
            from_code,
            to_code
        )

    if tr is None:
        return text

    try:

        result = tr.translate(text)

        if result:
            return normalize_text(result)

    except Exception:
        pass

    return text


# ============================================================
# TÌM GLOSSARY EXACT
# ============================================================

def glossary_exact(
    text,
    mode
):

    text = normalize_text(text)

    if mode == "Trung ➔ Việt":

        if text in TRANSLATION_MEMORY_ZH_VI:
            return TRANSLATION_MEMORY_ZH_VI[text]

        if text in GLOSSARY_ZH_VI:
            return GLOSSARY_ZH_VI[text]

    else:

        if text in TRANSLATION_MEMORY_VI_ZH:
            return TRANSLATION_MEMORY_VI_ZH[text]

        if text in GLOSSARY_VI_ZH:
            return GLOSSARY_VI_ZH[text]

    return None


# ============================================================
# TÁCH THUẬT NGỮ TRONG CÂU
# ============================================================

def replace_glossary_phrases(
    text,
    mode
):

    text = normalize_text(text)

    if not text:
        return text, False

    glossary = (
        GLOSSARY_ZH_VI
        if mode == "Trung ➔ Việt"
        else GLOSSARY_VI_ZH
    )

    changed = False

    # Cụm dài ưu tiên trước
    terms = sorted(
        glossary.items(),
        key=lambda x: len(x[0]),
        reverse=True
    )

    result = text

    for source, target in terms:

        if source in result:

            result = result.replace(
                source,
                f"[[TERM:{target}]]"
            )

            changed = True

    return result, changed


# ============================================================
# XỬ LÝ PLACEHOLDER THUẬT NGỮ
# ============================================================

def protect_glossary_terms(
    text,
    mode
):

    text = normalize_text(text)

    glossary = (
        GLOSSARY_ZH_VI
        if mode == "Trung ➔ Việt"
        else GLOSSARY_VI_ZH
    )

    terms = sorted(
        glossary.items(),
        key=lambda x: len(x[0]),
        reverse=True
    )

    placeholders = {}

    result = text

    for index, (source, target) in enumerate(terms):

        if source not in result:
            continue

        token = f"ZXTERM{index}ZX"

        result = result.replace(
            source,
            token
        )

        placeholders[token] = target

    return result, placeholders


def restore_glossary_terms(
    text,
    placeholders
):

    result = text

    for token, value in placeholders.items():

        result = result.replace(
            token,
            value
        )

    return result


# ============================================================
# HẬU XỬ LÝ DỊCH THUẬT
# ============================================================

def postprocess_translation(
    translated,
    mode
):

    translated = normalize_text(translated)

    if not translated:
        return translated

    glossary = (
        GLOSSARY_ZH_VI
        if mode == "Trung ➔ Việt"
        else GLOSSARY_VI_ZH
    )

    # Chuẩn hóa dấu câu
    translated = re.sub(
        r"\s+([,.!?;:])",
        r"\1",
        translated
    )

    # Hậu xử lý thuật ngữ
    for source, target in sorted(
        glossary.items(),
        key=lambda x: len(x[0]),
        reverse=True
    ):

        translated = translated.replace(
            source,
            target
        )

    # Xóa khoảng trắng thừa
    translated = re.sub(
        r"[ \t]+",
        " ",
        translated
    ).strip()

    return translated


# ============================================================
# DỊCH THÔNG MINH
# ============================================================

@lru_cache(maxsize=5000)
def smart_translate_cached(
    text,
    mode
):

    text = normalize_text(text)

    if not text:
        return ""

    # --------------------------------------------------------
    # 1. KHÔNG DỊCH SỐ / NGÀY / MÃ
    # --------------------------------------------------------

    if should_preserve(text):
        return text

    # --------------------------------------------------------
    # 2. EXACT TRANSLATION MEMORY / GLOSSARY
    # --------------------------------------------------------

    exact = glossary_exact(
        text,
        mode
    )

    if exact:
        return exact

    # --------------------------------------------------------
    # 3. PROTECT TERM
    # --------------------------------------------------------

    protected, placeholders = protect_glossary_terms(
        text,
        mode
    )

    # --------------------------------------------------------
    # 4. ARGOS TRỰC TIẾP
    # --------------------------------------------------------

    if mode == "Trung ➔ Việt":

        direct = translate_direct(
            protected,
            "zh",
            "vi"
        )

        confidence = 0.82

    else:

        direct = translate_direct(
            protected,
            "vi",
            "zh"
        )

        confidence = 0.82

    # --------------------------------------------------------
    # 5. NẾU ARGOS KHÔNG DỊCH ĐƯỢC
    # --------------------------------------------------------

    if not direct or direct == protected:

        if mode == "Trung ➔ Việt":

            english = translate_direct(
                protected,
                "zh",
                "en"
            )

            direct = translate_direct(
                english,
                "en",
                "vi"
            )

        else:

            english = translate_direct(
                protected,
                "vi",
                "en"
            )

            direct = translate_direct(
                english,
                "en",
                "zh"
            )

        confidence = 0.60

    # --------------------------------------------------------
    # 6. RESTORE TERM
    # --------------------------------------------------------

    result = restore_glossary_terms(
        direct,
        placeholders
    )

    # --------------------------------------------------------
    # 7. HẬU XỬ LÝ
    # --------------------------------------------------------

    result = postprocess_translation(
        result,
        mode
    )

    if not result:
        return text

    return result


def smart_translate(
    text,
    mode
):

    try:

        return smart_translate_cached(
            normalize_text(text),
            mode
        )

    except Exception:

        return normalize_text(text)


# ============================================================
# DỊCH KÈM THÔNG TIN ĐỘ TIN CẬY
# ============================================================

def translate_with_info(
    text,
    mode
):

    text = normalize_text(text)

    if not text:
        return {
            "text": "",
            "method": "empty",
            "confidence": 1.0
        }

    if should_preserve(text):
        return {
            "text": text,
            "method": "preserve",
            "confidence": 1.0
        }

    exact = glossary_exact(
        text,
        mode
    )

    if exact:
        return {
            "text": exact,
            "method": "glossary",
            "confidence": 1.0
        }

    result = smart_translate(
        text,
        mode
    )

    return {
        "text": result,
        "method": "argos",
        "confidence": 0.82
    }


# ============================================================
# IMAGE PREPROCESSING
# ============================================================

def preprocess_image(
    image,
    method=1
):

    if image.mode != "RGB":
        image = image.convert("RGB")

    image = ImageOps.exif_transpose(image)

    w, h = image.size

    # Upscale
    if w < 1800:
        scale = 2.5
    elif w < 2800:
        scale = 1.5
    else:
        scale = 1

    if scale != 1:

        image = image.resize(
            (
                int(w * scale),
                int(h * scale)
            ),
            Image.Resampling.LANCZOS
        )

    if method == 1:

        image = ImageEnhance.Contrast(
            image
        ).enhance(1.5)

        image = ImageEnhance.Sharpness(
            image
        ).enhance(1.5)

    elif method == 2:

        gray = ImageOps.grayscale(image)

        gray = ImageEnhance.Contrast(
            gray
        ).enhance(2.0)

        gray = gray.filter(
            ImageFilter.SHARPEN
        )

        image = gray.convert("RGB")

    elif method == 3:

        gray = ImageOps.grayscale(image)

        gray = ImageEnhance.Contrast(
            gray
        ).enhance(2.5)

        # Threshold nhẹ
        gray = gray.point(
            lambda p: 255 if p > 170 else 0
        )

        image = gray.convert("RGB")

    return image


# ============================================================
# OCR ENGINE
# ============================================================

@st.cache_resource
def get_ocr():

    if easyocr is None:
        raise RuntimeError(
            "Chưa cài EasyOCR. "
            "Hãy thêm easyocr vào requirements.txt"
        )

    return easyocr.Reader(
        ["ch_sim", "en"],
        gpu=False
    )


# ============================================================
# OCR MỘT LẦN
# ============================================================

def run_ocr_once(
    image
):

    reader = get_ocr()

    results = reader.readtext(
        np.array(image),
        detail=1,
        paragraph=False,
        width_ths=0.7,
        link_threshold=0.3,
        text_threshold=0.6,
        low_text=0.3,
    )

    items = []

    for result in results:

        if len(result) != 3:
            continue

        bbox, text, prob = result

        text = normalize_text(text)

        if not text:
            continue

        items.append({
            "text": text,
            "score": float(prob),
            "box": bbox
        })

    return items


# ============================================================
# OCR NHIỀU PASS
# ============================================================

def ocr_image(image):

    all_results = []

    for method in [1, 2]:

        try:

            processed = preprocess_image(
                image,
                method
            )

            result = run_ocr_once(
                processed
            )

            all_results.extend(result)

        except Exception:
            continue

    if not all_results:
        return []

    # Loại bỏ kết quả trùng
    unique = []

    seen = set()

    for item in all_results:

        text = item["text"]

        key = text.lower()

        if key in seen:
            continue

        seen.add(key)

        unique.append(item)

    # Ưu tiên confidence cao
    unique.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    return unique


# ============================================================
# BOX GEOMETRY
# ============================================================

def get_box_geometry(box):

    if not box:
        return None

    try:

        xs = [
            float(p[0])
            for p in box
        ]

        ys = [
            float(p[1])
            for p in box
        ]

        return (
            min(xs),
            min(ys),
            max(xs),
            max(ys)
        )

    except Exception:

        return None


# ============================================================
# GROUP OCR LINES
# ============================================================

def group_ocr_lines(items):

    prepared = []

    for item in items:

        geo = get_box_geometry(
            item.get("box")
        )

        if not geo:
            continue

        x1, y1, x2, y2 = geo

        prepared.append({
            **item,
            "x1": x1,
            "y1": y1,
            "x2": x2,
            "y2": y2,
            "cx": (x1 + x2) / 2,
            "cy": (y1 + y2) / 2,
            "height": max(y2 - y1, 1)
        })

    prepared.sort(
        key=lambda x: (
            x["cy"],
            x["x1"]
        )
    )

    lines = []

    for item in prepared:

        placed = False

        for line in lines:

            avg_y = sum(
                x["cy"]
                for x in line
            ) / len(line)

            avg_h = sum(
                x["height"]
                for x in line
            ) / len(line)

            tolerance = max(
                14,
                avg_h * 0.65
            )

            if abs(
                item["cy"] - avg_y
            ) <= tolerance:

                line.append(item)
                placed = True
                break

        if not placed:
            lines.append([item])

    for line in lines:
        line.sort(
            key=lambda x: x["x1"]
        )

    return lines


def line_to_text(line):

    return " ".join(
        item["text"].strip()
        for item in line
        if item["text"].strip()
    )


# ============================================================
# DATE
# ============================================================

def extract_date(text):

    patterns = [

        r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})",

        r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})",

        r"(\d{4})年(\d{1,2})月(\d{1,2})日",

        r"(\d{4})年(\d{1,2})月",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text
        )

        if not match:
            continue

        groups = match.groups()

        try:

            if len(groups[0]) == 4:

                return (
                    f"{int(groups[0]):04d}-"
                    f"{int(groups[1]):02d}-"
                    f"{int(groups[2]):02d}"
                )

            return (
                f"{int(groups[2]):04d}-"
                f"{int(groups[1]):02d}-"
                f"{int(groups[0]):02d}"
            )

        except Exception:
            continue

    return ""


# ============================================================
# CLEAN NUMBER
# ============================================================

def clean_number(text):

    if text is None:
        return ""

    text = str(text).strip()

    # Trung Quốc dùng dấu phẩy
    text = text.replace(
        "，",
        "."
    )

    # Chỉ đổi comma nếu có số
    if re.fullmatch(
        r"-?\d+,\d+",
        text
    ):
        text = text.replace(
            ",",
            "."
        )

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text
    )

    if not match:
        return text

    try:

        value = float(
            match.group(0)
        )

        if value.is_integer():
            return int(value)

        return value

    except Exception:

        return text


# ============================================================
# HEADER DETECTION
# ============================================================

def detect_header_line(lines):

    keywords = [
        "部门",
        "开几台",
        "正式",
        "临时",
        "备注",
        "姓名",
        "工号",
        "班次",
        "bộ phận",
        "số máy",
        "chính thức",
        "thời vụ",
        "ghi chú",
        "stt",
        "tên",
        "ca",
    ]

    best_idx = None
    best_score = 0

    for idx, line in enumerate(lines):

        text = line_to_text(
            line
        ).lower()

        score = 0

        for keyword in keywords:

            if keyword in text:
                score += 1

        if score > best_score:

            best_score = score
            best_idx = idx

    return best_idx


# ============================================================
# CLASSIFY COLUMN
# ============================================================

def classify_column_dynamic(name):

    name = normalize_text(
        name
    ).lower()

    if any(
        k in name
        for k in [
            "部门",
            "bộ phận",
            "phòng",
            "xưởng",
            "tên",
            "姓名",
        ]
    ):
        return "dept_src"

    if any(
        k in name
        for k in [
            "机器",
            "台",
            "开机",
            "开几台",
            "máy",
            "số máy",
        ]
    ):
        return "machines"

    if any(
        k in name
        for k in [
            "正式",
            "chính thức",
        ]
    ):
        return "formal"

    if any(
        k in name
        for k in [
            "临时",
            "thời vụ",
        ]
    ):
        return "temp"

    if any(
        k in name
        for k in [
            "备注",
            "ghi chú",
            "说明",
        ]
    ):
        return "remark"

    return "unknown"


# ============================================================
# PARSE ATTENDANCE ROWS
# ============================================================

def parse_attendance_rows(
    lines,
    header_index
):

    if header_index is None:
        header_index = 0

    if header_index >= len(lines):
        return []

    columns = []

    for item in lines[header_index]:

        columns.append({
            "name": item["text"].lower(),
            "x": item["cx"]
        })

    rows = []

    for line in lines[
        header_index + 1:
    ]:

        if not line:
            continue

        line_text = line_to_text(
            line
        )

        if not line_text:
            continue

        lower_text = line_text.lower()

        # Bỏ dòng tổng
        if any(
            kw in lower_text
            for kw in [
                "一共",
                "总计",
                "合计",
                "tổng cộng",
                "total"
            ]
        ):
            continue

        stt = len(rows) + 1

        if line:

            first_text = (
                line[0]["text"]
                .strip()
            )

            match_stt = re.fullmatch(
                r"(\d+)[\.\)\、]?",
                first_text
            )

            if match_stt:

                stt = int(
                    match_stt.group(1)
                )

        row = {
            "stt": stt,
            "dept_src": "",
            "dept_tgt": "",
            "machines": "",
            "formal": "",
            "temp": "",
            "remark": "",
        }

        for item in line:

            if not columns:
                continue

            col = min(
                columns,
                key=lambda c: abs(
                    c["x"] - item["cx"]
                )
            )

            c_type = classify_column_dynamic(
                col["name"]
            )

            txt = item["text"].strip()

            if c_type == "dept_src":

                row["dept_src"] = (
                    row["dept_src"]
                    + " "
                    + txt
                ).strip()

            elif c_type == "machines":

                row["machines"] = clean_number(
                    txt
                )

            elif c_type == "formal":

                row["formal"] = clean_number(
                    txt
                )

            elif c_type == "temp":

                row["temp"] = clean_number(
                    txt
                )

            elif c_type == "remark":

                row["remark"] = (
                    row["remark"]
                    + " "
                    + txt
                ).strip()

        # Fallback tìm bộ phận
        if not row["dept_src"]:

            for item in line:

                txt = item["text"].strip()

                if not txt:
                    continue

                if re.fullmatch(
                    r"\d+(?:\.\d+)?",
                    txt
                ):
                    continue

                if any(
                    k in txt
                    for k in [
                        "总计",
                        "合计"
                    ]
                ):
                    continue

                row["dept_src"] = txt
                break

        rows.append(row)

    return rows


# ============================================================
# PARSE IMAGE
# ============================================================

def parse_attendance_image(
    image,
    mode
):

    items = ocr_image(
        image
    )

    if not items:
        raise RuntimeError(
            "Không đọc được chữ từ ảnh."
        )

    lines = group_ocr_lines(
        items
    )

    if not lines:
        raise RuntimeError(
            "Không tạo được dòng OCR."
        )

    header_index = detect_header_line(
        lines
    )

    all_text = "\n".join(
        line_to_text(line)
        for line in lines
    )

    date_str = extract_date(
        all_text
    )

    if header_index is not None:

        title_lines = [
            line_to_text(line)
            for line in lines[
                :header_index
            ]
        ]

    else:

        title_lines = []

    title_src = " ".join(
        x
        for x in title_lines
        if x
    ).strip()

    rows = parse_attendance_rows(
        lines,
        header_index
    )

    title_tgt = (
        smart_translate(
            title_src,
            mode
        )
        if title_src
        else ""
    )

    for row in rows:

        if row.get("dept_src"):

            row["dept_tgt"] = smart_translate(
                row["dept_src"],
                mode
            )

    return {
        "title_src": title_src,
        "title_tgt": title_tgt,
        "date_str": date_str,
        "rows": rows,
    }


# ============================================================
# PDF -> IMAGE
# ============================================================

def pdf_to_images(
    pdf_bytes
):

    if fitz is None:

        raise RuntimeError(
            "Chưa cài PyMuPDF. "
            "Hãy thêm pymupdf vào requirements.txt"
        )

    doc = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    images = []

    try:

        for page in doc:

            pix = page.get_pixmap(
                matrix=fitz.Matrix(
                    3.0,
                    3.0
                ),
                alpha=False
            )

            image = Image.frombytes(
                "RGB",
                [
                    pix.width,
                    pix.height
                ],
                pix.samples
            )

            images.append(image)

    finally:

        doc.close()

    return images


# ============================================================
# MERGE DOCUMENTS
# ============================================================

def merge_parsed_documents(
    documents
):

    if not documents:

        return {
            "title_src": "",
            "title_tgt": "",
            "date_str": "",
            "rows": []
        }

    merged = {
        "title_src": documents[0].get(
            "title_src",
            ""
        ),
        "title_tgt": documents[0].get(
            "title_tgt",
            ""
        ),
        "date_str": documents[0].get(
            "date_str",
            ""
        ),
        "rows": []
    }

    next_stt = 1

    for document in documents:

        for row in document.get(
            "rows",
            []
        ):

            new_row = dict(row)

            new_row["stt"] = next_stt

            merged["rows"].append(
                new_row
            )

            next_stt += 1

    return merged


# ============================================================
# EXCEL
# ============================================================

def build_excel_from_json(
    data,
    mode
):

    wb = Workbook()

    ws = wb.active

    ws.title = "Bảng chấm công"

    font_name = "Microsoft YaHei"

    orange_fill = PatternFill(
        fill_type="solid",
        fgColor="ED7D00"
    )

    light_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )

    thin = Side(
        style="thin",
        color="000000"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    title_src = data.get(
        "title_src",
        ""
    )

    title_tgt = data.get(
        "title_tgt",
        ""
    )

    date_str = data.get(
        "date_str",
        ""
    )

    rows = data.get(
        "rows",
        []
    )

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    full_title = (
        f"{date_str} {title_src}\n"
        f"{title_tgt}"
    ).strip()

    ws.merge_cells(
        "A1:F1"
    )

    ws["A1"] = full_title

    ws["A1"].font = Font(
        name=font_name,
        size=13,
        bold=True
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    ws["A1"].border = border

    ws.row_dimensions[1].height = 48

    # --------------------------------------------------------
    # HEADERS
    # --------------------------------------------------------

    if mode == "Trung ➔ Việt":

        headers = [
            ("STT", "STT"),
            ("部门", "Bộ phận"),
            ("开几台机", "Số máy chạy"),
            ("正式工", "Công nhân chính thức"),
            ("临时工", "Công nhân thời vụ"),
            ("备注", "Ghi chú"),
        ]

    else:

        headers = [
            ("STT", "STT"),
            ("Bộ phận", "部门"),
            ("Số máy chạy", "开几台机"),
            ("Công nhân chính thức", "正式工"),
            ("Công nhân thời vụ", "临时工"),
            ("Ghi chú", "备注"),
        ]

    for col_idx, (
        source_header,
        target_header
    ) in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            row=2,
            column=col_idx
        )

        if source_header == target_header:

            cell.value = source_header

        else:

            cell.value = (
                f"{source_header}\n"
                f"{target_header}"
            )

        cell.font = Font(
            name=font_name,
            size=10,
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.fill = orange_fill

        cell.border = border

    ws.row_dimensions[2].height = 52

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    current_row = 3

    for row in rows:

        values = [
            row.get("stt", ""),
            (
                f"{row.get('dept_src', '')}\n"
                f"{row.get('dept_tgt', '')}"
            ).strip(),
            row.get("machines", ""),
            row.get("formal", ""),
            row.get("temp", ""),
            row.get("remark", ""),
        ]

        for col_idx, value in enumerate(
            values,
            start=1
        ):

            cell = ws.cell(
                row=current_row,
                column=col_idx,
                value=value
            )

            cell.font = Font(
                name=font_name,
                size=10
            )

            cell.alignment = Alignment(
                horizontal=(
                    "center"
                    if col_idx in [
                        1, 3, 4, 5
                    ]
                    else "left"
                ),
                vertical="center",
                wrap_text=True
            )

            cell.border = border

        ws.row_dimensions[
            current_row
        ].height = 38

        current_row += 1

    # --------------------------------------------------------
    # COLUMN WIDTH
    # --------------------------------------------------------

    widths = {
        "A": 8,
        "B": 34,
        "C": 14,
        "D": 20,
        "E": 20,
        "F": 24,
    }

    for col, width in widths.items():

        ws.column_dimensions[
            col
        ].width = width

    # --------------------------------------------------------
    # FREEZE
    # --------------------------------------------------------

    ws.freeze_panes = "A3"

    # --------------------------------------------------------
    # FILTER
    # --------------------------------------------------------

    if current_row > 3:

        ws.auto_filter.ref = (
            f"A2:F{current_row - 1}"
        )

    # --------------------------------------------------------
    # PAGE
    # --------------------------------------------------------

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ============================================================
# THỐNG KÊ
# ============================================================

def get_translation_statistics(
    data,
    mode
):

    total = 0
    glossary_count = 0
    normal_count = 0
    preserve_count = 0

    for row in data.get(
        "rows",
        []
    ):

        text = row.get(
            "dept_src",
            ""
        )

        if not text:
            continue

        total += 1

        if should_preserve(text):

            preserve_count += 1
            continue

        if glossary_exact(
            text,
            mode
        ):

            glossary_count += 1

        else:

            normal_count += 1

    return {
        "total": total,
        "glossary": glossary_count,
        "argos": normal_count,
        "preserve": preserve_count
    }


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header(
    "⚙️ Cấu hình"
)

mode = st.sidebar.selectbox(
    "Chọn chiều dịch",
    [
        "Trung ➔ Việt",
        "Việt ➔ Trung"
    ]
)

st.sidebar.markdown(
    "---"
)

st.sidebar.subheader(
    "📚 Bộ từ điển"
)

st.sidebar.write(
    f"Thuật ngữ Trung–Việt: "
    f"**{len(GLOSSARY_ZH_VI)}**"
)

st.sidebar.write(
    f"Translation Memory: "
    f"**{len(TRANSLATION_MEMORY_ZH_VI)}**"
)

st.sidebar.info(
    "Hệ thống không sử dụng AI/API. "
    "Argos Translate chạy local."
)


# ============================================================
# UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📤 Tải lên tệp bảng chấm công",
    type=[
        "png",
        "jpg",
        "jpeg",
        "pdf"
    ]
)


# ============================================================
# MAIN
# ============================================================

if uploaded_file is not None:

    try:

        # ----------------------------------------------------
        # INIT
        # ----------------------------------------------------

        with st.spinner(
            "Đang khởi tạo model dịch offline..."
        ):

            initialize_translation_models()

        # ----------------------------------------------------
        # LOAD FILE
        # ----------------------------------------------------

        if uploaded_file.type == "application/pdf":

            images = pdf_to_images(
                uploaded_file.getvalue()
            )

        else:

            image = Image.open(
                uploaded_file
            )

            images = [image]

        st.info(
            f"Đã phát hiện {len(images)} trang/ảnh."
        )

        # ----------------------------------------------------
        # PROCESS
        # ----------------------------------------------------

        documents = []

        progress = st.progress(
            0
        )

        total_images = len(images)

        for index, image in enumerate(
            images
        ):

            with st.spinner(
                f"Đang OCR & phân tích trang "
                f"{index + 1}/{total_images}..."
            ):

                document = parse_attendance_image(
                    image,
                    mode
                )

                documents.append(
                    document
                )

            progress.progress(
                int(
                    ((index + 1) / total_images)
                    * 100
                )
            )

        # ----------------------------------------------------
        # MERGE
        # ----------------------------------------------------

        final_data = merge_parsed_documents(
            documents
        )

        # ----------------------------------------------------
        # SUCCESS
        # ----------------------------------------------------

        st.success(
            "✅ Xử lý thành công!"
        )

        # ----------------------------------------------------
        # STATISTICS
        # ----------------------------------------------------

        stats = get_translation_statistics(
            final_data,
            mode
        )

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(
                "Số dòng",
                stats["total"]
            )

        with col2:
            st.metric(
                "Glossary",
                stats["glossary"]
            )

        with col3:
            st.metric(
                "Argos",
                stats["argos"]
            )

        with col4:
            st.metric(
                "Giữ nguyên",
                stats["preserve"]
            )

        # ----------------------------------------------------
        # PREVIEW
        # ----------------------------------------------------

        st.subheader(
            "👁️ Kết quả nhận diện"
        )

        if final_data["title_src"]:

            st.write(
                "**Tiêu đề gốc:**",
                final_data["title_src"]
            )

            st.write(
                "**Tiêu đề dịch:**",
                final_data["title_tgt"]
            )

        if final_data["date_str"]:

            st.write(
                "**Ngày:**",
                final_data["date_str"]
            )

        # ----------------------------------------------------
        # TABLE
        # ----------------------------------------------------

        preview_rows = []

        for row in final_data["rows"]:

            preview_rows.append({

                "STT": row.get(
                    "stt",
                    ""
                ),

                "Bộ phận gốc": row.get(
                    "dept_src",
                    ""
                ),

                "Bộ phận dịch": row.get(
                    "dept_tgt",
                    ""
                ),

                "Số máy": row.get(
                    "machines",
                    ""
                ),

                "Chính thức": row.get(
                    "formal",
                    ""
                ),

                "Thời vụ": row.get(
                    "temp",
                    ""
                ),

                "Ghi chú": row.get(
                    "remark",
                    ""
                ),
            })

        if preview_rows:

            st.dataframe(
                preview_rows,
                use_container_width=True,
                hide_index=True
            )

        else:

            st.warning(
                "Không tìm thấy dòng dữ liệu."
            )

        # ----------------------------------------------------
        # DEBUG JSON
        # ----------------------------------------------------

        with st.expander(
            "🔍 Xem dữ liệu JSON nội bộ"
        ):

            st.json(
                final_data
            )

        # ----------------------------------------------------
        # EXCEL
        # ----------------------------------------------------

        excel_data = build_excel_from_json(
            final_data,
            mode
        )

        st.download_button(
            label="📥 Tải xuống Excel Song Ngữ",
            data=excel_data,
            file_name=(
                "bang_cham_cong_song_ngu.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    except Exception as e:

        st.error(
            f"❌ Lỗi xử lý: {e}"
        )

        with st.expander(
            "Chi tiết lỗi"
        ):

            st.exception(e)
