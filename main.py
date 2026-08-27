from pathlib import Path

code = r'''import io
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
        # Với tiếng Trung có thể bỏ khoảng trắng trong OCR.
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
        # Mã/ký hiệu thường không nên dịch.
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
    """
    Chỉ cài 4 model cần cho đường vòng:
    ZH -> EN -> VI
    VI -> EN -> ZH

    Không ép cài zh_vi/vi_zh vì không phải môi trường Argos nào
    cũng có đúng package đó.
    """
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

    # Khôi phục dấu câu cơ bản bị mất ở đầu/cuối.
    source_s = normalize_text(source)

    if source_s.endswith(":") and not translated.endswith(":"):
        translated += ":"

    # Nếu bản dịch vẫn còn quá nhiều ký tự Trung khi Trung -> Việt,
    # không tự ý thay đổi vì có thể là tên riêng/mã.
    return translated


def translate_text(text, mode, context="general", remember=True):
    text = normalize_text(text)

    if not text:
        return ""

    if looks_like_non_translatable(text):
        return text

    # 1. Translation Memory.
    cached = tm_lookup(text, mode, context)
    if cached:
        return cached

    # 2. Exact glossary.
    exact = apply_glossary_exact(text, mode)
    if exact is not None:
        if remember:
            tm_save(text, exact, mode, context)
        return exact

    # 3. Cụm thuật ngữ.
    glossary_result = apply_glossary_phrases(text, mode)

    # Nếu glossary đã chuyển toàn bộ câu, dùng luôn.
    if glossary_result != text and not re.search(r"[\u3400-\u9fff]", glossary_result):
        result = glossary_result
        if remember:
            tm_save(text, result, mode, context)
        return result

    # 4. Bảo vệ số/mã/ngày.
    protected_text, protected = protect_tokens(glossary_result)

    # 5. Nếu còn tiếng Trung/Vietnamese cần dịch.
    if mode == "Trung ➔ Việt":
        # Ưu tiên direct ZH -> VI nếu môi trường đã có model.
        direct = find_argos_translation("zh", "vi")

        if direct is not None:
            result = argos_translate_once(
                protected_text, "zh", "vi"
            )
        else:
            # Fallback: ZH -> EN -> VI.
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

    # Không phóng đại quá mức.
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

    # ch_sim = tiếng Trung giản thể
    # en = tiếng Anh
    # vi = tiếng Việt
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

        # Header thường có nhiều cụm ngắn.
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

        # Fallback bộ phận.
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
            # Có thể bật dịch ghi chú nếu muốn.
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
        "title_src": documents[0].get(
            "title_src", ""
        ),
        "title_tgt": documents[0].get(
            "title_tgt", ""
        ),
        "date_str": documents[0].get(
            "date_str", ""
        ),
        "rows": []
    }

    next_stt = 1

    for doc in documents:
        for row in doc.get("rows", []):
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
def build_excel_from_json(data, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_name = "Microsoft YaHei"

    orange_fill = PatternFill(
        fill_type="solid",
        fgColor="ED7D00"
    )

    thin_side = Side(
        style="thin",
        color="000000"
    )

    border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )

    title_src = data.get("title_src", "")
    title_tgt = data.get("title_tgt", "")
    date_str = data.get("date_str", "")
    rows = data.get("rows", [])

    title_parts = []

    if date_str:
        title_parts.append(date_str)

    if title_src:
        title_parts.append(title_src)

    full_title_src = " ".join(title_parts).strip()

    if title_tgt:
        full_title = (
            f"{full_title_src}\n"
            f"{title_tgt}"
        ).strip()
    else:
        full_title = full_title_src

    ws.merge_cells("A1:F1")

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

    ws.row_dimensions[1].height = 48

    if mode == "Trung ➔ Việt":
        headers = [
            ("STT", "STT"),
            ("部门", "Bộ phận"),
            ("开几台机", "Số máy mở"),
            ("正式工", "Chính thức"),
            ("临时工", "Thời vụ"),
            ("备注", "Ghi chú")
        ]
    else:
        headers = [
            ("STT", "STT"),
            ("Bộ phận", "部门"),
            ("Số máy mở", "开几台机"),
            ("Chính thức", "正式工"),
            ("Thời vụ", "临时工"),
            ("Ghi chú", "备注")
        ]

    for col_idx, (top_h, bottom_h) in enumerate(
        headers,
        start=1
    ):
        cell = ws.cell(
            row=2,
            column=col_idx
        )

        cell.value = (
            f"{top_h}\n{bottom_h}"
            if top_h != bottom_h
            else top_h
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

    ws.row_dimensions[2].height = 38

    current_row = 3

    for row in rows:
        ws.cell(
            row=current_row,
            column=1,
            value=row.get("stt", "")
        )

        dept_src = row.get("dept_src", "")
        dept_tgt = row.get("dept_tgt", "")

        dept_value = dept_src

        if dept_tgt:
            dept_value = (
                f"{dept_src}\n{dept_tgt}"
            ).strip()

        ws.cell(
            row=current_row,
            column=2,
            value=dept_value
        )

        ws.cell(
            row=current_row,
            column=3,
            value=row.get("machines", "")
        )

        ws.cell(
            row=current_row,
            column=4,
            value=row.get("formal", "")
        )

        ws.cell(
            row=current_row,
            column=5,
            value=row.get("temp", "")
        )

        # Ghi chú giữ nguyên + bản dịch nếu có.
        remark_src = row.get("remark", "")
        remark_tgt = row.get("remark_tgt", "")

        if remark_src and remark_tgt:
            remark_value = (
                f"{remark_src}\n{remark_tgt}"
            )
        else:
            remark_value = remark_src or remark_tgt

        ws.cell(
            row=current_row,
            column=6,
            value=remark_value
        )

        for col_idx in range(1, 7):
            cell = ws.cell(
                row=current_row,
                column=col_idx
            )

            cell.font = Font(
                name=font_name,
                size=10
            )

            cell.alignment = Alignment(
                horizontal=(
                    "center"
                    if col_idx in [1, 3, 4, 5]
                    else "left"
                ),
                vertical="center",
                wrap_text=True
            )

            cell.border = border

        ws.row_dimensions[current_row].height = 38

        current_row += 1

    widths = {
        "A": 8,
        "B": 32,
        "C": 14,
        "D": 14,
        "E": 14,
        "F": 30
    }

    for col_letter, width in widths.items():
        ws.column_dimensions[
            col_letter
        ].width = width

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ============================================================
# GIAO DIỆN GLOSSARY / TRANSLATION MEMORY
# ============================================================
def glossary_editor():
    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 Từ điển chuyên ngành")

    custom = get_custom_glossary()

    with st.sidebar.expander(
        "➕ Thêm thuật ngữ",
        expanded=False
    ):
        src = st.text_input(
            "Từ/cụm nguồn",
            key="glossary_src"
        )

        tgt = st.text_input(
            "Bản dịch",
            key="glossary_tgt"
        )

        if st.button(
            "Lưu thuật ngữ",
            key="save_glossary"
        ):
            if src.strip() and tgt.strip():
                custom[src.strip()] = {
                    "translation": tgt.strip(),
                    "type": "custom",
                    "priority": 200
                }

                save_json_file(
                    GLOSSARY_FILE,
                    custom
                )

                st.success(
                    "Đã lưu từ điển."
                )

                st.rerun()

    with st.sidebar.expander(
        "📖 Từ điển hiện tại",
        expanded=False
    ):
        if custom:
            for key, value in custom.items():
                if isinstance(value, dict):
                    st.write(
                        f"**{key}** → "
                        f"{value.get('translation', '')}"
                    )
                else:
                    st.write(
                        f"**{key}** → {value}"
                    )
        else:
            st.caption(
                "Chưa có thuật ngữ tùy chỉnh."
            )

        if custom:
            if st.button(
                "🗑 Xóa toàn bộ từ điển tùy chỉnh",
                key="clear_glossary"
            ):
                save_json_file(
                    GLOSSARY_FILE,
                    {}
                )

                st.rerun()


def memory_editor():
    st.sidebar.markdown("---")

    with st.sidebar.expander(
        "🧠 Translation Memory",
        expanded=False
    ):
        memory = get_translation_memory()

        st.write(
            f"Số mục đã nhớ: **{len(memory)}**"
        )

        if memory:
            if st.button(
                "🗑 Xóa Translation Memory",
                key="clear_tm"
            ):
                save_json_file(
                    TM_FILE,
                    {}
                )

                st.rerun()


# ============================================================
# GIAO DIỆN CHÍNH
# ============================================================
st.title(
    "🌐 Dịch & Xuất Bảng Chấm Công Song Ngữ"
)

st.caption(
    "Local OCR + Glossary + Translation Memory + Argos "
    "| Không cần API AI"
)

mode = st.sidebar.selectbox(
    "Chọn chiều dịch",
    [
        "Trung ➔ Việt",
        "Việt ➔ Trung"
    ]
)

multi_pass = st.sidebar.checkbox(
    "🔎 OCR nhiều lượt để tăng độ chính xác",
    value=True
)

show_ocr = st.sidebar.checkbox(
    "Hiển thị dữ liệu OCR thô",
    value=False
)

glossary_editor()
memory_editor()

uploaded_file = st.file_uploader(
    "Tải lên bảng chấm công",
    type=[
        "png",
        "jpg",
        "jpeg",
        "pdf"
    ]
)


# ============================================================
# XỬ LÝ FILE
# ============================================================
if uploaded_file is not None:
    try:
        with st.spinner(
            "Đang kiểm tra / chuẩn bị model dịch local..."
        ):
            model_errors = (
                initialize_translation_models()
            )

        if model_errors:
            st.warning(
                "Một số model chưa sẵn sàng:\n\n"
                + "\n".join(model_errors)
            )

        images = []

        if uploaded_file.type == "application/pdf":
            with st.spinner(
                "Đang chuyển PDF thành ảnh..."
            ):
                images = pdf_to_images(
                    uploaded_file.getvalue()
                )
        else:
            image = Image.open(
                uploaded_file
            ).convert("RGB")

            images = [image]

        st.info(
            f"Đã nhận **{len(images)} trang/ảnh**."
        )

        parsed_docs = []

        for page_index, img in enumerate(images, start=1):
            with st.spinner(
                f"Đang OCR + phân tích trang {page_index}/{len(images)}..."
            ):
                parsed = parse_attendance_image(
                    img,
                    mode,
                    multi_pass=multi_pass
                )

                parsed_docs.append(parsed)

        final_data = merge_parsed_documents(
            parsed_docs
        )

        st.success(
            "Đã xử lý xong."
        )

        if show_ocr:
            st.subheader(
                "🔍 Dữ liệu sau OCR / phân tích"
            )

            st.json(final_data)

        st.subheader(
            "📋 Xem trước kết quả"
        )

        preview_rows = []

        for row in final_data["rows"]:
            preview_rows.append({
                "STT": row.get("stt", ""),
                "Bộ phận nguồn": row.get(
                    "dept_src", ""
                ),
                "Bộ phận dịch": row.get(
                    "dept_tgt", ""
                ),
                "Số máy": row.get(
                    "machines", ""
                ),
                "Chính thức": row.get(
                    "formal", ""
                ),
                "Thời vụ": row.get(
                    "temp", ""
                ),
                "Ghi chú": row.get(
                    "remark", ""
                ),
                "Ghi chú dịch": row.get(
                    "remark_tgt", ""
                ),
            })

        if preview_rows:
            st.dataframe(
                preview_rows,
                use_container_width=True
            )
        else:
            st.warning(
                "Không phát hiện được dòng dữ liệu."
            )

        excel_data = build_excel_from_json(
            final_data,
            mode
        )

        st.download_button(
            label=(
                "📥 Tải xuống Excel song ngữ"
            ),
            data=excel_data,
            file_name=(
                "bang_cham_cong_song_ngu.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )

    except Exception as e:
        st.error(
            f"Đã xảy ra lỗi khi xử lý tệp: {e}"
        )

        with st.expander(
            "Chi tiết lỗi"
        ):
            st.exception(e)
'''
path = "/mnt/data/main_local_translation.py"
Path(path).write_text(code, encoding="utf-8")
print(path)
