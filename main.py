import io
import re
import json
import time
import math
import urllib.request
from pathlib import Path

import streamlit as st
import openpyxl
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

# ============================================================
# THƯ VIỆN DỊCH LOCAL + OCR
# ============================================================
import argostranslate.package
import argostranslate.translate

try:
    import easyocr
except Exception:
    easyocr = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None


# ============================================================
# CẤU HÌNH
# ============================================================
st.set_page_config(
    page_title="Dịch bảng chấm công Local",
    page_icon="🌐",
    layout="wide"
)

APP_DIR = Path(".attendance_translator")
APP_DIR.mkdir(exist_ok=True)

ARGOS_DIR = APP_DIR / "argos_models"
ARGOS_DIR.mkdir(exist_ok=True)

TM_FILE = APP_DIR / "translation_memory.json"
GLOSSARY_FILE = APP_DIR / "custom_glossary.json"

ARGOS_BASE_URL = "https://data.argosopentech.com/argospm/v1/"

# Chỉ dùng các model chắc chắn theo cấu hình gốc.
ARGOS_PACKAGES = {
    "zh_en": "translate-zh_en-1_9.argosmodel",
    "en_zh": "translate-en_zh-1_9.argosmodel",
    "vi_en": "translate-vi_en-1_9.argosmodel",
    "en_vi": "translate-en_vi-1_9.argosmodel",
}

# ============================================================
# TỪ ĐIỂN CHUYÊN NGÀNH MẶC ĐỊNH
# priority càng cao càng được ưu tiên.
# ============================================================
DEFAULT_GLOSSARY_ZH_VI = {
    "正式工": {"translation": "Công nhân chính thức", "type": "attendance", "priority": 100},
    "临时工": {"translation": "Công nhân thời vụ", "type": "attendance", "priority": 100},
    "部门": {"translation": "Bộ phận", "type": "department", "priority": 100},
    "开几台机": {"translation": "Số máy mở", "type": "machine", "priority": 100},
    "备注": {"translation": "Ghi chú", "type": "remark", "priority": 100},
    "总计": {"translation": "Tổng cộng", "type": "total", "priority": 100},
    "合计": {"translation": "Tổng cộng", "type": "total", "priority": 100},
    "姓名": {"translation": "Họ tên", "type": "name", "priority": 100},
    "班次": {"translation": "Ca làm việc", "type": "shift", "priority": 100},
    "早班": {"translation": "Ca sáng", "type": "shift", "priority": 100},
    "晚班": {"translation": "Ca đêm", "type": "shift", "priority": 100},
    "白班": {"translation": "Ca ngày", "type": "shift", "priority": 100},
    "夜班": {"translation": "Ca đêm", "type": "shift", "priority": 100},
    "休息": {"translation": "Nghỉ", "type": "attendance", "priority": 100},
    "一车间": {"translation": "Xưởng 1", "type": "department", "priority": 100},
    "二车间": {"translation": "Xưởng 2", "type": "department", "priority": 100},
    "三车间": {"translation": "Xưởng 3", "type": "department", "priority": 100},
    "生产部": {"translation": "Bộ phận sản xuất", "type": "department", "priority": 100},
    "品管部": {"translation": "Bộ phận quản lý chất lượng", "type": "department", "priority": 100},
    "品管": {"translation": "Quản lý chất lượng", "type": "department", "priority": 100},
    "仓库": {"translation": "Kho", "type": "department", "priority": 100},
    "裁断组": {"translation": "Tổ cắt", "type": "department", "priority": 100},
    "裁断": {"translation": "Cắt", "type": "department", "priority": 100},
    "针车组": {"translation": "Tổ may", "type": "department", "priority": 100},
    "针车": {"translation": "May", "type": "department", "priority": 100},
    "成型": {"translation": "Định hình", "type": "department", "priority": 100},
    "包装": {"translation": "Đóng gói", "type": "department", "priority": 100},
    "原料": {"translation": "Nguyên liệu", "type": "department", "priority": 100},
    "材料": {"translation": "Vật liệu", "type": "department", "priority": 100},
    "人事部": {"translation": "Bộ phận nhân sự", "type": "department", "priority": 100},
    "财务部": {"translation": "Bộ phận tài vụ", "type": "department", "priority": 100},
    "经理": {"translation": "Quản lý", "type": "title", "priority": 80},
    "主管": {"translation": "Chủ quản", "type": "title", "priority": 80},
    "组长": {"translation": "Tổ trưởng", "type": "title", "priority": 80},
    "厂长": {"translation": "Quản đốc", "type": "title", "priority": 80},
}

DEFAULT_GLOSSARY_VI_ZH = {
    value["translation"]: {
        "translation": key,
        "type": value.get("type", "general"),
        "priority": value.get("priority", 100),
    }
    for key, value in DEFAULT_GLOSSARY_ZH_VI.items()
}


# ============================================================
# FILE JSON
# ============================================================
def load_json_file(path, default):
    try:
        if not path.exists():
            return default.copy() if isinstance(default, dict) else default
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, type(default)) else default
    except Exception:
        return default.copy() if isinstance(default, dict) else default


def save_json_file(path, data):
    tmp = path.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def get_custom_glossary():
    return load_json_file(GLOSSARY_FILE, {})


def get_translation_memory():
    return load_json_file(TM_FILE, {})


# ============================================================
# NORMALIZE
# ============================================================
def normalize_text(text):
    if text is None:
        return ""

    text = str(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def normalize_lookup(text):
    text = normalize_text(text)
    text = text.replace(" ", "")
    return text.lower()


# ============================================================
# GLOSSARY
# ============================================================
def merged_glossary(mode):
    custom = get_custom_glossary()

    if mode == "Trung ➔ Việt":
        result = dict(DEFAULT_GLOSSARY_ZH_VI)
    else:
        result = dict(DEFAULT_GLOSSARY_VI_ZH)

    for key, value in custom.items():
        if isinstance(value, str):
            result[key] = {
                "translation": value,
                "type": "custom",
                "priority": 200,
            }
        elif isinstance(value, dict) and value.get("translation"):
            result[key] = {
                "translation": str(value["translation"]),
                "type": value.get("type", "custom"),
                "priority": int(value.get("priority", 200)),
            }

    return result


def apply_glossary_exact(text, mode):
    glossary = merged_glossary(mode)
    key = normalize_lookup(text)

    for src, item in glossary.items():
        if normalize_lookup(src) == key:
            return item["translation"]

    return None


def apply_glossary_phrases(text, mode):
    """
    Thay các thuật ngữ theo cụm dài nhất trước.
    Không phá cấu trúc câu.
    """
    glossary = merged_glossary(mode)
    result = normalize_text(text)

    entries = []
    for src, item in glossary.items():
        src = normalize_text(src)
        if not src:
            continue
        entries.append((
            int(item.get("priority", 100)),
            len(src),
            src,
            str(item["translation"])
        ))

    entries.sort(key=lambda x: (x[0], x[1]), reverse=True)

    for _, _, src, tgt in entries:
        if re.search(r"[\u3400-\u9fff]", src):
            pattern = re.escape(src)
            result = re.sub(pattern, tgt, result, flags=re.IGNORECASE)
        else:
            pattern = r"(?<!\w)" + re.escape(src) + r"(?!\w)"
            result = re.sub(pattern, tgt, result, flags=re.IGNORECASE)

    return result


# ============================================================
# BẢO VỆ SỐ / MÃ / NGÀY TRƯỚC KHI DỊCH
# ============================================================
PROTECT_PATTERNS = [
    r"\b\d{4}[-/.]\d{1,2}[-/.]\d{1,2}\b",
    r"\b\d{1,2}[-/.]\d{1,2}[-/.]\d{4}\b",
    r"\b\d{1,2}:\d{2}(?::\d{2})?\b",
    r"\b\d+(?:[.,]\d+)?%\b",
    r"\b\d+(?:[.,]\d+)?\b",
    r"\b[A-Z]{1,8}[-_/]?\d{1,8}\b",
    r"\b\d+[A-Za-z]+[A-Za-z0-9_-]*\b",
    r"\b[A-Za-z]+\d+[A-Za-z0-9_-]*\b",
    r"\b\d+#\b",
]


def protect_tokens(text):
    protected = {}
    counter = 0

    def repl(match):
        nonlocal counter
        token = match.group(0)
        marker = f"ZXQPROTECT{counter}QXZ"
        protected[marker] = token
        counter += 1
        return marker

    result = text
    for pattern in PROTECT_PATTERNS:
        result = re.sub(pattern, repl, result)

    return result, protected


def restore_tokens(text, protected):
    result = text
    for marker, original in protected.items():
        result = result.replace(marker, original)
        result = result.replace(marker.lower(), original)
    return result


def looks_like_non_translatable(text):
    text = normalize_text(text)
    if not text:
        return True

    if re.fullmatch(r"[\d\s.,:/\\%#()+\-]+", text):
        return True

    if re.fullmatch(r"[A-Za-z0-9_\-/#. ]+", text):
        if re.search(r"\d", text):
            return True

    return False


# ============================================================
# ARGOS
# ============================================================
@st.cache_resource
def get_argos_languages():
    try:
        return argostranslate.translate.get_installed_languages()
    except Exception:
        return []


def find_argos_translation(from_code, to_code):
    languages = get_argos_languages()

    from_lang = None
    to_lang = None

    for lang in languages:
        if lang.code == from_code:
            from_lang = lang
        if lang.code == to_code:
            to_lang = lang

    if from_lang is None or to_lang is None:
        return None

    try:
        return from_lang.get_translation(to_lang)
    except Exception:
        return None


def download_argos_package(package_name):
    target = ARGOS_DIR / package_name

    if target.exists() and target.stat().st_size > 0:
        return target

    url = ARGOS_BASE_URL + package_name

    try:
        urllib.request.urlretrieve(url, str(target))
        return target
    except Exception as e:
        if target.exists():
            try:
                target.unlink()
            except Exception:
                pass

        raise RuntimeError(
            f"Không thể tải model Argos '{package_name}'. Lỗi: {e}"
        )


def install_argos_package(from_code, to_code):
    key = f"{from_code}_{to_code}"

    if key not in ARGOS_PACKAGES:
        raise RuntimeError(
            f"Chưa cấu hình model Argos {from_code} -> {to_code}"
        )

    package_path = download_argos_package(ARGOS_PACKAGES[key])

    try:
        argostranslate.package.install_from_path(str(package_path))
    except Exception as e:
        if "already installed" not in str(e).lower():
            raise RuntimeError(f"Lỗi cài model Argos: {e}")

    try:
        get_argos_languages.clear()
    except Exception:
        pass


def ensure_argos_pair(from_code, to_code):
    translation = find_argos_translation(from_code, to_code)

    if translation is not None:
        return translation

    install_argos_package(from_code, to_code)
    return find_argos_translation(from_code, to_code)


@st.cache_resource
def initialize_translation_models():
    required_pairs = [
        ("zh", "en"),
        ("en", "zh"),
        ("vi", "en"),
        ("en", "vi"),
    ]

    errors = []

    for from_code, to_code in required_pairs:
        try:
            ensure_argos_pair(from_code, to_code)
        except Exception as e:
            errors.append(f"{from_code}->{to_code}: {e}")

    return errors


def argos_translate_once(text, from_code, to_code):
    text = normalize_text(text)

    if not text or from_code == to_code:
        return text

    translation = find_argos_translation(from_code, to_code)

    if translation is None:
        try:
            translation = ensure_argos_pair(from_code, to_code)
        except Exception:
            return text

    if translation is None:
        return text

    try:
        result = translation.translate(text)
        return normalize_text(result) if result else text
    except Exception:
        return text


# ============================================================
# TRANSLATION MEMORY
# ============================================================
def tm_key(text, mode, context):
    return f"{mode}|{context}|{normalize_lookup(text)}"


def tm_lookup(text, mode, context):
    memory = get_translation_memory()
    key = tm_key(text, mode, context)

    item = memory.get(key)

    if isinstance(item, dict):
        return item.get("translation", "")

    if isinstance(item, str):
        return item

    return None


def tm_save(text, translation, mode, context="general"):
    text = normalize_text(text)
    translation = normalize_text(translation)

    if not text or not translation:
        return

    memory = get_translation_memory()
    key = tm_key(text, mode, context)

    memory[key] = {
        "source": text,
        "translation": translation,
        "mode": mode,
        "context": context,
        "updated": time.strftime("%Y-%m-%d %H:%M:%S"),
    }

    save_json_file(TM_FILE, memory)


# ============================================================
# DỊCH THEO NGỮ CẢNH
# ============================================================
def context_hint(context):
    hints = {
        "title": "tiêu đề bảng chấm công",
        "department": "tên bộ phận hoặc xưởng trong nhà máy",
        "machine": "số lượng máy",
        "formal": "số công nhân chính thức",
        "temp": "số công nhân thời vụ",
        "remark": "ghi chú trong bảng chấm công",
        "shift": "ca làm việc",
        "general": "bảng chấm công nhà máy",
    }
    return hints.get(context, hints["general"])


def postprocess_translation(source, translated, mode):
    translated = normalize_text(translated)

    if not translated:
        return normalize_text(source)

    source_s = normalize_text(source)

    if source_s.endswith(":") and not translated.endswith(":"):
        translated += ":"

    return translated


def translate_text(text, mode, context="general", remember=True):
    text = normalize_text(text)

    if not text:
        return ""

    if looks_like_non_translatable(text):
        return text

    cached = tm_lookup(text, mode, context)
    if cached:
        return cached

    exact = apply_glossary_exact(text, mode)
    if exact is not None:
        if remember:
            tm_save(text, exact, mode, context)
        return exact

    glossary_result = apply_glossary_phrases(text, mode)

    if glossary_result != text and not re.search(r"[\u3400-\u9fff]", glossary_result):
        result = glossary_result
        if remember:
            tm_save(text, result, mode, context)
        return result

    protected_text, protected = protect_tokens(glossary_result)

    if mode == "Trung ➔ Việt":
        direct = find_argos_translation("zh", "vi")

        if direct is not None:
            result = argos_translate_once(
                protected_text, "zh", "vi"
            )
        else:
            english = argos_translate_once(
                protected_text, "zh", "en"
            )
            result = argos_translate_once(
                english, "en", "vi"
            )

    else:
        direct = find_argos_translation("vi", "zh")

        if direct is not None:
            result = argos_translate_once(
                protected_text, "vi", "zh"
            )
        else:
            english = argos_translate_once(
                protected_text, "vi", "en"
            )
            result = argos_translate_once(
                english, "en", "zh"
            )

    result = restore_tokens(result, protected)
    result = postprocess_translation(text, result, mode)

    if remember and result and result != text:
        tm_save(text, result, mode, context)

    return result


# ============================================================
# OCR
# ============================================================
def preprocess_variants(image):
    if image.mode != "RGB":
        image = image.convert("RGB")

    width, height = image.size
    scale = 2 if width < 1800 else 1

    if scale > 1:
        image = image.resize(
            (width * scale, height * scale),
            Image.Resampling.LANCZOS
        )

    rgb = image

    gray = ImageOps.grayscale(rgb)
    gray = ImageEnhance.Contrast(gray).enhance(1.5)
    gray = ImageEnhance.Sharpness(gray).enhance(1.4)

    sharp = ImageEnhance.Contrast(rgb).enhance(1.7)
    sharp = ImageEnhance.Sharpness(sharp).enhance(1.8)

    denoise = rgb.filter(ImageFilter.MedianFilter(size=3))
    denoise = ImageEnhance.Contrast(denoise).enhance(1.4)

    return [
        ("original", rgb),
        ("gray", gray.convert("RGB")),
        ("sharp", sharp),
        ("denoise", denoise),
    ]


@st.cache_resource
def get_ocr():
    if easyocr is None:
        raise RuntimeError(
            "Chưa cài EasyOCR. Hãy cài easyocr trong requirements.txt."
        )

    return easyocr.Reader(
        ["ch_sim", "en", "vi"],
        gpu=False,
        verbose=False
    )


def raw_ocr(image):
    reader = get_ocr()

    results = reader.readtext(
        np.array(image),
        detail=1,
        paragraph=False,
        text_threshold=0.55,
        low_text=0.25,
        link_threshold=0.35,
        mag_ratio=1.0
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


def ocr_quality_score(items):
    if not items:
        return 0.0

    scores = [max(0.0, min(1.0, x["score"])) for x in items]
    avg = sum(scores) / len(scores)

    useful = sum(
        1 for x in items
        if re.search(r"[\u3400-\u9fffA-Za-zÀ-ỹ0-9]", x["text"])
    )

    return avg * math.log1p(useful)


def ocr_image(image, multi_pass=True):
    variants = preprocess_variants(image)

    if not multi_pass:
        variants = variants[:1]

    candidates = []

    for name, variant in variants:
        try:
            items = raw_ocr(variant)
            score = ocr_quality_score(items)

            candidates.append({
                "name": name,
                "items": items,
                "score": score
            })
        except Exception:
            continue

    if not candidates:
        return []

    candidates.sort(key=lambda x: x["score"], reverse=True)

    return candidates[0]["items"]


# ============================================================
# HÌNH HỌC OCR
# ============================================================
def get_box_geometry(box):
    if not box:
        return None

    try:
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]

        return (
            min(xs),
            min(ys),
            max(xs),
            max(ys)
        )
    except Exception:
        return None


def group_ocr_lines(items):
    prepared = []

    for item in items:
        geo = get_box_geometry(item.get("box"))

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

    prepared.sort(key=lambda x: (x["cy"], x["x1"]))

    lines = []

    for item in prepared:
        placed = False

        for line in lines:
            avg_y = sum(x["cy"] for x in line) / len(line)
            avg_h = sum(x["height"] for x in line) / len(line)

            tolerance = max(10, avg_h * 0.65)

            if abs(item["cy"] - avg_y) <= tolerance:
                line.append(item)
                placed = True
                break

        if not placed:
            lines.append([item])

    for line in lines:
        line.sort(key=lambda x: x["x1"])

    return lines


def line_to_text(line):
    return " ".join(
        item["text"].strip()
        for item in line
        if item["text"].strip()
    )


# ============================================================
# NGÀY / SỐ
# ============================================================
def extract_date(text):
    patterns = [
        r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})",
        r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})",
        r"(\d{4})年(\d{1,2})月(\d{1,2})日",
    ]

    for pattern in patterns:
        m = re.search(pattern, text)

        if not m:
            continue

        groups = m.groups()

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
            pass

    return ""


def clean_number(text):
    if text is None:
        return ""

    text = str(text).strip()
    text = text.replace(",", ".").replace("，", ".")

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text
    )

    if not match:
        return text

    try:
        value = float(match.group(0))

        if value.is_integer():
            return int(value)

        return value
    except Exception:
        return text


# ============================================================
# NHẬN DIỆN HEADER
# ============================================================
HEADER_KEYWORDS = [
    "部门",
    "开几台",
    "正式",
    "临时",
    "备注",
    "姓名",
    "班次",
    "bộ phận",
    "số máy",
    "chính thức",
    "thời vụ",
    "ghi chú",
    "họ tên",
    "ca",
    "stt",
    "tên"
]


def detect_header_line(lines):
    best_idx = None
    best_score = 0

    for idx, line in enumerate(lines):
        text = line_to_text(line).lower()
        score = 0

        for kw in HEADER_KEYWORDS:
            if kw.lower() in text:
                score += 1

        if len(line) >= 3:
            score += 0.5

        if score > best_score:
            best_score = score
            best_idx = idx

    return best_idx


def classify_column_dynamic(column_name):
    text = normalize_text(column_name).lower()

    if any(k in text for k in [
        "部门", "bộ phận", "phòng", "xưởng", "tên"
    ]):
        return "dept_src"

    if any(k in text for k in [
        "机器", "机械", "máy", "台", "开机", "số máy"
    ]):
        return "machines"

    if any(k in text for k in [
        "正式", "chính thức", "cố định"
    ]):
        return "formal"

    if any(k in text for k in [
        "临时", "thời vụ", "phụ"
    ]):
        return "temp"

    if any(k in text for k in [
        "备注", "ghi chú", "chú thích"
    ]):
        return "remark"

    return "unknown"


# ============================================================
# PARSE BẢNG
# ============================================================
def parse_attendance_rows(lines, header_index):
    if not lines:
        return []

    if header_index is None:
        header_index = 0

    header_line = lines[header_index]
    columns = []

    for item in header_line:
        name = normalize_text(item["text"]).lower()

        if not name:
            continue

        columns.append({
            "name": name,
            "x": item["cx"]
        })

    rows = []

    for line in lines[header_index + 1:]:
        if not line:
            continue

        line_text = line_to_text(line)

        if not line_text:
            continue

        lower_text = line_text.lower()

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

        stt = ""
        first_text = normalize_text(line[0]["text"])

        stt_match = re.match(
            r"^\s*(\d+)[.\\)]?\s*$",
            first_text
        )

        if stt_match:
            stt = int(stt_match.group(1))
        else:
            match = re.match(
                r"^\s*(\d+)",
                line_text
            )

            if match:
                stt = int(match.group(1))
            else:
                stt = len(rows) + 1

        row = {
            "stt": stt,
            "dept_src": "",
            "dept_tgt": "",
            "machines": "",
            "formal": "",
            "temp": "",
            "remark": ""
        }

        for item in line:
            if not columns:
                continue

            nearest_col = min(
                columns,
                key=lambda c: abs(c["x"] - item["cx"])
            )

            col_type = classify_column_dynamic(
                nearest_col["name"]
            )

            text = normalize_text(item["text"])

            if not text:
                continue

            if col_type == "dept_src":
                row["dept_src"] = (
                    row["dept_src"] + " " + text
                ).strip()

            elif col_type == "machines":
                row["machines"] = clean_number(text)

            elif col_type == "formal":
                row["formal"] = clean_number(text)

            elif col_type == "temp":
                row["temp"] = clean_number(text)

            elif col_type == "remark":
                row["remark"] = (
                    row["remark"] + " " + text
                ).strip()

        if not row["dept_src"]:
            for item in line:
                txt = normalize_text(item["text"])

                if not txt:
                    continue

                if not re.fullmatch(
                    r"\d+(?:\.\d+)?",
                    txt
                ):
                    row["dept_src"] = txt
                    break

        rows.append(row)

    return rows


# ============================================================
# PARSE ẢNH
# ============================================================
def parse_attendance_image(image, mode, multi_pass=True):
    items = ocr_image(
        image,
        multi_pass=multi_pass
    )

    if not items:
        raise RuntimeError(
            "Không đọc được chữ từ ảnh."
        )

    lines = group_ocr_lines(items)

    if not lines:
        raise RuntimeError(
            "OCR có kết quả nhưng không thể gom thành dòng."
        )

    header_index = detect_header_line(lines)

    all_text = "\n".join(
        line_to_text(line)
        for line in lines
    )

    date_str = extract_date(all_text)

    if header_index is not None:
        title_lines = [
            line_to_text(line)
            for line in lines[:header_index]
        ]
    else:
        title_lines = []

    title_src = " ".join(
        t for t in title_lines if t
    ).strip()

    rows = parse_attendance_rows(
        lines,
        header_index
    )

    title_tgt = ""

    if title_src:
        title_tgt = translate_text(
            title_src,
            mode,
            context="title"
        )

    for row in rows:
        if row.get("dept_src"):
            row["dept_tgt"] = translate_text(
                row["dept_src"],
                mode,
                context="department"
            )

        if row.get("remark"):
            row["remark_tgt"] = translate_text(
                row["remark"],
                mode,
                context="remark"
            )
        else:
            row["remark_tgt"] = ""

    return {
        "title_src": title_src,
        "title_tgt": title_tgt,
        "date_str": date_str,
        "rows": rows
    }


# ============================================================
# PDF
# ============================================================
def pdf_to_images(pdf_bytes, dpi_scale=2.5):
    if fitz is None:
        raise RuntimeError(
            "Chưa cài PyMuPDF."
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
                    dpi_scale,
                    dpi_scale
                ),
                alpha=False
            )

            image = Image.frombytes(
                "RGB",
                [pix.width, pix.height],
                pix.samples
            )

            images.append(image)

    finally:
        doc.close()

    return images


# ============================================================
# MERGE NHIỀU TRANG
# ============================================================
def merge_parsed_documents(documents):
    if not documents:
        return {
            "title_src": "",
            "title_tgt": "",
            "date_str": "",
            "rows": []
        }

    merged = {
        "title_src": documents[0].get("title_src", ""),
        "title_tgt": documents[0].get("title_tgt", ""),
        "date_str": documents[0].get("date_str", ""),
        "rows": []
    }

    next_stt = 1

    for doc in documents:
        for row in doc.get("rows", []):
            new_row = dict(row)
            new_row["stt"] = next_stt
            merged["rows"].append(new_row)
            next_stt += 1

        if not merged["date_str"] and doc.get("date_str"):
            merged["date_str"] = doc.get("date_str")

    return merged


# ============================================================
# XUẤT EXCEL
# ============================================================
def export_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng Chấm Công"

    font_title = Font(name="Arial", size=14, bold=True)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Arial", size=10)
    
    fill_header = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws.append([data.get("title_tgt") or data.get("title_src", "Bảng Chấm Công")])
    ws.cell(row=1, column=1).font = font_title
    ws.append([])

    headers = ["STT", "Bộ phận (Gốc)", "Bộ phận (Dịch)", "Số máy", "Chính thức", "Thời vụ", "Ghi chú"]
    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    for row_idx, row in enumerate(data.get("rows", []), start=4):
        ws.append([
            row.get("stt", ""),
            row.get("dept_src", ""),
            row.get("dept_tgt", ""),
            row.get("machines", ""),
            row.get("formal", ""),
            row.get("temp", ""),
            row.get("remark_tgt") or row.get("remark", "")
        ])
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = font_data
            cell.border = thin_border
            if col_idx in [1, 4, 5, 6]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# GIAO DIỆN STREAMLIT
# ============================================================
def main():
    st.title("🌐 Dịch Bảng Chấm Công Local")
    st.write("Ứng dụng hỗ trợ dịch bảng chấm công từ tiếng Trung sang tiếng Việt và ngược lại sử dụng EasyOCR và Argos Translate.")

    with st.spinner("Đang khởi tạo các mô hình dịch..."):
        initialize_translation_models()

    mode = st.sidebar.selectbox("Chọn chiều dịch", ["Trung ➔ Việt", "Việt ➔ Trung"])
    
    uploaded_files = st.file_uploader(
        "Tải lên ảnh hoặc file PDF bảng chấm công", 
        type=["png", "jpg", "jpeg", "pdf"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("Bắt đầu xử lý và dịch"):
            all_documents = []
            
            for uploaded_file in uploaded_files:
                file_bytes = uploaded_file.read()
                
                if uploaded_file.type == "application/pdf":
                    try:
                        images = pdf_to_images(file_bytes)
                        for img in images:
                            res = parse_attendance_image(img, mode)
                            all_documents.append(res)
                    except Exception as e:
                        st.error(f"Lỗi khi đọc file PDF {uploaded_file.name}: {e}")
                else:
                    try:
                        image = Image.open(io.BytesIO(file_bytes))
                        res = parse_attendance_image(image, mode)
                        all_documents.append(res)
                    except Exception as e:
                        st.error(f"Lỗi khi đọc ảnh {uploaded_file.name}: {e}")

            if all_documents:
                merged_data = merge_parsed_documents(all_documents)
                st.success("Xử lý thành công!")
                
                st.subheader("Kết quả dịch")
                st.write(f"**Tiêu đề:** {merged_data.get('title_tgt')}")
                st.write(f"**Ngày:** {merged_data.get('date_str')}")

                rows = merged_data.get("rows", [])
                if rows:
                    import pandas as pd
                    df = pd.DataFrame(rows)
                    st.dataframe(df, use_container_width=True)

                    excel_data = export_to_excel(merged_data)
                    st.download_button(
                        label="📥 Tải xuống file Excel",
                        data=excel_data,
                        file_name="bang_cham_cong_da_dich.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("Không tìm thấy dữ liệu dòng nào trong tài liệu.")

if __name__ == "__main__":
    main()
