import io
import re
import json
import hashlib
import urllib.request
from pathlib import Path
from collections import defaultdict

import streamlit as st
import openpyxl
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

# ============================================================
# LOCAL TRANSLATION + OCR
# Không dùng AI/API/quota
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
# CẤU HÌNH STREAMLIT
# ============================================================
st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Thông Minh",
    page_icon="🌐",
    layout="wide",
)

st.title("🌐 Dịch & Xuất Bảng Chấm Công Song Ngữ")
st.caption(
    "Local 100% | Không dùng AI/API | OCR đa tầng | Từ điển chuyên ngành | "
    "Translation Memory | Tự nhận diện cột"
)


# ============================================================
# THƯ MỤC LOCAL
# ============================================================
BASE_DIR = Path(".")
MODEL_DIR = BASE_DIR / ".argos_models"
MEMORY_FILE = BASE_DIR / "translation_memory.json"
MODEL_DIR.mkdir(exist_ok=True)


# ============================================================
# TỪ ĐIỂN CHUYÊN NGÀNH
# ============================================================
GLOSSARY_ZH_VI = {
    # Hệ thống chung
    "正式工": "Công nhân chính thức",
    "临时工": "Công nhân thời vụ",
    "部门": "Bộ phận",
    "车间": "Xưởng",
    "班组": "Tổ sản xuất",
    "生产线": "Chuyền sản xuất",
    "开几台机": "Số máy đang chạy",
    "开机": "Chạy máy",
    "机台": "Máy",
    "机器": "Máy",
    "备注": "Ghi chú",
    "总计": "Tổng cộng",
    "合计": "Tổng cộng",
    "姓名": "Họ tên",
    "班次": "Ca làm việc",
    "早班": "Ca sáng",
    "白班": "Ca ngày",
    "晚班": "Ca đêm",
    "夜班": "Ca đêm",
    "休息": "Nghỉ",
    "出勤": "Đi làm",
    "缺勤": "Vắng mặt",
    "迟到": "Đi muộn",
    "早退": "Về sớm",
    "加班": "Tăng ca",
    "请假": "Xin nghỉ",
    "病假": "Nghỉ bệnh",
    "事假": "Nghỉ việc riêng",

    # Sản xuất / nhà máy
    "裁断": "Cắt",
    "裁断组": "Tổ cắt",
    "针车": "May",
    "针车组": "Tổ may",
    "成型": "Thành hình",
    "成型组": "Tổ thành hình",
    "包装": "Đóng gói",
    "包装组": "Tổ đóng gói",
    "品检": "Kiểm hàng",
    "品质": "Chất lượng",
    "仓库": "Kho",
    "仓管": "Quản lý kho",
    "维修": "Bảo trì",
    "机修": "Bảo trì máy",
    "行政": "Hành chính",
    "人事": "Nhân sự",
    "财务": "Tài vụ",
    "采购": "Mua hàng",
    "生管": "Quản lý sản xuất",
    "生产": "Sản xuất",
}


GLOSSARY_VI_ZH = {v: k for k, v in GLOSSARY_ZH_VI.items()}


# ============================================================
# CỤM TỪ OCR THƯỜNG NHẦM
# ============================================================
OCR_NORMALIZATION_ZH = {
    "正式エ": "正式工",
    "正式T": "正式工",
    "正式丁": "正式工",
    "临时エ": "临时工",
    "臨時工": "临时工",
    "部門": "部门",
    "備注": "备注",
    "總計": "总计",
    "合計": "合计",
    "開機": "开机",
    "機台": "机台",
    "車間": "车间",
    "針車": "针车",
    "裁斷": "裁断",
    "成型組": "成型组",
    "裁斷組": "裁断组",
}


# ============================================================
# ARGOS MODELS
# ============================================================
ARGOS_PACKAGES = {
    "zh_en": "translate-zh_en-1_9.argosmodel",
    "en_zh": "translate-en_zh-1_9.argosmodel",
    "vi_en": "translate-vi_en-1_9.argosmodel",
    "en_vi": "translate-en_vi-1_9.argosmodel",
}

ARGOS_BASE_URL = "https://data.argosopentech.com/argospm/v1/"


# ============================================================
# CACHE / SESSION
# ============================================================
if "translation_stats" not in st.session_state:
    st.session_state.translation_stats = {
        "glossary": 0,
        "memory": 0,
        "argos": 0,
        "unchanged": 0,
    }


# ============================================================
# TEXT UTILITIES
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


def normalize_ocr_text(text):
    text = normalize_text(text)
    if not text:
        return ""

    for src, dst in sorted(
        OCR_NORMALIZATION_ZH.items(),
        key=lambda x: len(x[0]),
        reverse=True,
    ):
        text = text.replace(src, dst)

    text = re.sub(r"(?<=正)\s+(?=式)", "", text)
    text = re.sub(r"(?<=临)\s+(?=时)", "", text)
    text = re.sub(r"(?<=开)\s+(?=机)", "", text)
    text = re.sub(r"(?<=部)\s+(?=门)", "", text)

    return text.strip()


def is_number_like(text):
    text = normalize_text(text)
    if not text:
        return False

    if re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", text):
        return True

    if re.fullmatch(r"[\d\s.,:/\\\-+%]+", text):
        return True

    if re.fullmatch(r"[\d\s.,:/\\\-+%]+(?:人|台|个|名|小时|天)", text):
        return True

    return False


def should_translate(text):
    text = normalize_text(text)
    if not text:
        return False

    if is_number_like(text):
        return False

    if re.fullmatch(r"[A-Za-z0-9_\-/.:]+", text):
        if any(c.isalpha() for c in text) and any(c.isdigit() for c in text):
            return False

    return True


# ============================================================
# GLOSSARY
# ============================================================
def apply_glossary_exact(text, mode):
    text = normalize_text(text)
    if not text:
        return None

    glossary = GLOSSARY_ZH_VI if mode == "Trung ➔ Việt" else GLOSSARY_VI_ZH
    return glossary.get(text)


def apply_glossary_replace(text, mode):
    text = normalize_text(text)
    if not text:
        return text

    glossary = GLOSSARY_ZH_VI if mode == "Trung ➔ Việt" else GLOSSARY_VI_ZH

    result = text
    for src in sorted(glossary.keys(), key=len, reverse=True):
        result = result.replace(src, glossary[src])

    return result


def glossary_for_context(context, mode):
    if mode == "Trung ➔ Việt":
        base = GLOSSARY_ZH_VI
    else:
        base = GLOSSARY_VI_ZH
    return base


def apply_context_glossary(text, mode, context=None):
    text = normalize_text(text)
    if not text:
        return text

    glossary = glossary_for_context(context, mode)

    for src in sorted(glossary.keys(), key=len, reverse=True):
        text = text.replace(src, glossary[src])

    return text


# ============================================================
# TRANSLATION MEMORY
# ============================================================
def load_translation_memory():
    if not MEMORY_FILE.exists():
        return {
            "zh_vi": {},
            "vi_zh": {},
        }

    try:
        data = json.loads(MEMORY_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Translation memory không hợp lệ.")

        data.setdefault("zh_vi", {})
        data.setdefault("vi_zh", {})
        return data
    except Exception:
        return {
            "zh_vi": {},
            "vi_zh": {},
        }


def save_translation_memory(memory):
    tmp = MEMORY_FILE.with_suffix(".tmp")
    tmp.write_text(
        json.dumps(memory, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    tmp.replace(MEMORY_FILE)


def memory_key(mode):
    return "zh_vi" if mode == "Trung ➔ Việt" else "vi_zh"


def memory_lookup(text, mode):
    text = normalize_text(text)
    if not text:
        return None

    memory = load_translation_memory()
    return memory.get(memory_key(mode), {}).get(text)


def memory_save(source, target, mode):
    source = normalize_text(source)
    target = normalize_text(target)

    if not source or not target:
        return

    if source == target:
        return

    memory = load_translation_memory()
    key = memory_key(mode)

    if source not in memory[key]:
        memory[key][source] = target
        try:
            save_translation_memory(memory)
        except Exception:
            pass


# ============================================================
# ARGOS
# ============================================================
@st.cache_resource
def get_argos_languages():
    try:
        return argostranslate.translate.get_installed_languages()
    except Exception:
        return []


def clear_argos_language_cache():
    try:
        get_argos_languages.clear()
    except Exception:
        pass


def find_argos_language(code):
    for lang in get_argos_languages():
        if lang.code == code:
            return lang
    return None


def find_argos_translation(from_code, to_code):
    installed_languages = get_argos_languages()

    from_lang = None
    to_lang = None

    for lang in installed_languages:
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
    MODEL_DIR.mkdir(exist_ok=True)
    target = MODEL_DIR / package_name

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
        raise RuntimeError(f"Không thể tải model Argos: {e}")


def install_argos_package(from_code, to_code):
    package_key = f"{from_code}_{to_code}"

    if package_key not in ARGOS_PACKAGES:
        raise RuntimeError(
            f"Chưa cấu hình model Argos {from_code} -> {to_code}"
        )

    package_path = download_argos_package(ARGOS_PACKAGES[package_key])

    try:
        argostranslate.package.install_from_path(str(package_path))
    except Exception as e:
        message = str(e).lower()
        if "already installed" not in message:
            raise

    clear_argos_language_cache()


@st.cache_resource
def initialize_translation_models():
    required_pairs = [
        ("zh", "en"),
        ("en", "zh"),
        ("vi", "en"),
        ("en", "vi"),
    ]

    for from_code, to_code in required_pairs:
        if find_argos_translation(from_code, to_code) is None:
            install_argos_package(from_code, to_code)


def translate_direct(text, from_code, to_code):
    text = normalize_text(text)

    if not text or from_code == to_code:
        return text

    translation = find_argos_translation(from_code, to_code)

    if translation is None:
        install_argos_package(from_code, to_code)
        translation = find_argos_translation(from_code, to_code)

    if translation is None:
        return text

    try:
        result = translation.translate(text)
        result = normalize_text(result)
        return result if result else text
    except Exception:
        return text


# ============================================================
# POST PROCESS TRANSLATION
# ============================================================
def postprocess_translation(source, translated, mode):
    source = normalize_text(source)
    translated = normalize_text(translated)

    if not translated:
        return source

    if is_number_like(source):
        return source

    if mode == "Trung ➔ Việt":
        glossary = GLOSSARY_ZH_VI
    else:
        glossary = GLOSSARY_VI_ZH

    for src, dst in sorted(glossary.items(), key=lambda x: len(x[0]), reverse=True):
        if src in translated:
            translated = translated.replace(src, dst)

    return translated.strip()


# ============================================================
# SMART TRANSLATOR
# ============================================================
def translate_text(text, mode, context=None):
    text = normalize_ocr_text(text)

    if not text:
        return ""

    if not should_translate(text):
        st.session_state.translation_stats["unchanged"] += 1
        return text

    exact = apply_glossary_exact(text, mode)

    if exact is not None:
        st.session_state.translation_stats["glossary"] += 1
        return exact

    memory_result = memory_lookup(text, mode)

    if memory_result:
        st.session_state.translation_stats["memory"] += 1
        return memory_result

    context_text = apply_context_glossary(text, mode, context)

    if context_text != text:
        if mode == "Trung ➔ Việt":
            has_zh = bool(re.search(r"[\u3400-\u9fff]", context_text))
        else:
            has_zh = False

        if not has_zh:
            result = normalize_text(context_text)
            memory_save(text, result, mode)
            st.session_state.translation_stats["glossary"] += 1
            return result

    if mode == "Trung ➔ Việt":
        direct = find_argos_translation("zh", "vi")

        if direct is not None:
            try:
                result = normalize_text(direct.translate(context_text))
                if result:
                    result = postprocess_translation(text, result, mode)
                    memory_save(text, result, mode)
                    st.session_state.translation_stats["argos"] += 1
                    return result
            except Exception:
                pass

        english = translate_direct(context_text, "zh", "en")

        if not english or english == context_text:
            return text

        vietnamese = translate_direct(english, "en", "vi")
        result = normalize_text(vietnamese)

    else:
        direct = find_argos_translation("vi", "zh")

        if direct is not None:
            try:
                result = normalize_text(direct.translate(context_text))
                if result:
                    result = postprocess_translation(text, result, mode)
                    memory_save(text, result, mode)
                    st.session_state.translation_stats["argos"] += 1
                    return result
            except Exception:
                pass

        english = translate_direct(context_text, "vi", "en")

        if not english or english == context_text:
            return text

        chinese = translate_direct(english, "en", "zh")
        result = normalize_text(chinese)

    result = postprocess_translation(text, result, mode)

    if result:
        memory_save(text, result, mode)
        st.session_state.translation_stats["argos"] += 1
        return result

    st.session_state.translation_stats["unchanged"] += 1
    return text


# ============================================================
# OCR PREPROCESSING
# ============================================================
def resize_for_ocr(image):
    if image.mode != "RGB":
        image = image.convert("RGB")

    width, height = image.size

    if width < 1800:
        scale = 2.0
    elif width < 3000:
        scale = 1.5
    else:
        scale = 1.0

    if scale != 1.0:
        image = image.resize(
            (int(width * scale), int(height * scale)),
            Image.Resampling.LANCZOS,
        )

    return image


def create_ocr_variants(image):
    base = resize_for_ocr(image)
    variants = []

    variants.append(("original", base))

    enhanced = ImageEnhance.Contrast(base).enhance(1.45)
    enhanced = ImageEnhance.Sharpness(enhanced).enhance(1.6)
    variants.append(("enhanced", enhanced))

    gray = ImageOps.grayscale(base)
    gray = ImageOps.autocontrast(gray)
    variants.append(("gray", gray))

    sharp = gray.filter(ImageFilter.SHARPEN)
    sharp = ImageEnhance.Contrast(sharp).enhance(1.35)
    variants.append(("sharp_gray", sharp))

    arr = np.array(gray)
    threshold = 175
    binary = np.where(arr > threshold, 255, 0).astype(np.uint8)
    binary_img = Image.fromarray(binary)
    variants.append(("threshold", binary_img))

    return variants


@st.cache_resource
def get_ocr():
    if easyocr is None:
        raise RuntimeError(
            "Chưa cài EasyOCR. Hãy cài easyocr và torch trong requirements.txt."
        )

    return easyocr.Reader(
        ["ch_sim", "en"],
        gpu=False,
        verbose=False,
    )


# ============================================================
# OCR GEOMETRY
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
            max(ys),
        )
    except Exception:
        return None


def box_iou(box_a, box_b):
    if not box_a or not box_b:
        return 0.0

    ax1, ay1, ax2, ay2 = box_a
    bx1, by1, bx2, by2 = box_b

    ix1 = max(ax1, bx1)
    iy1 = max(ay1, by1)
    ix2 = min(ax2, bx2)
    iy2 = min(ay2, by2)

    iw = max(0.0, ix2 - ix1)
    ih = max(0.0, iy2 - iy1)

    inter = iw * ih

    area_a = max(0.0, ax2 - ax1) * max(0.0, ay2 - ay1)
    area_b = max(0.0, bx2 - bx1) * max(0.0, by2 - by1)

    union = area_a + area_b - inter

    if union <= 0:
        return 0.0

    return inter / union


def deduplicate_ocr_results(items):
    prepared = []

    for item in items:
        geo = get_box_geometry(item.get("box"))

        if not geo:
            continue

        text = normalize_ocr_text(item.get("text", ""))

        if not text:
            continue

        prepared.append(
            {
                **item,
                "text": text,
                "geo": geo,
                "score": float(item.get("score", 0.0)),
            }
        )

    prepared.sort(key=lambda x: x["score"], reverse=True)
    kept = []

    for item in prepared:
        duplicate = False

        for old in kept:
            iou = box_iou(item["geo"], old["geo"])

            if iou >= 0.45:
                duplicate = True
                break

            ax1, ay1, ax2, ay2 = item["geo"]
            bx1, by1, bx2, by2 = old["geo"]

            acx = (ax1 + ax2) / 2
            acy = (ay1 + ay2) / 2

            bcx = (bx1 + bx2) / 2
            bcy = (by1 + by2) / 2

            aw = max(ax2 - ax1, 1)
            ah = max(ay2 - ay1, 1)

            if (
                abs(acx - bcx) <= aw * 0.35
                and abs(acy - bcy) <= ah * 0.5
                and item["text"] == old["text"]
            ):
                duplicate = True
                break

        if not duplicate:
            kept.append(item)

    kept.sort(
        key=lambda x: (
            (x["geo"][1] + x["geo"][3]) / 2,
            x["geo"][0],
        )
    )

    for item in kept:
        item.pop("geo", None)

    return kept


def ocr_image(image):
    reader = get_ocr()
    variants = create_ocr_variants(image)
    all_results = []

    for variant_name, variant in variants:
        try:
            results = reader.readtext(
                np.array(variant),
                detail=1,
                paragraph=False,
                text_threshold=0.45,
                low_text=0.25,
                link_threshold=0.25,
                mag_ratio=1.0,
            )
        except Exception:
            continue

        for bbox, text, prob in results:
            text = normalize_ocr_text(text)

            if not text:
                continue

            if float(prob) < 0.25:
                continue

            all_results.append(
                {
                    "text": text,
                    "score": float(prob),
                    "box": bbox,
                    "variant": variant_name,
                }
            )

    return deduplicate_ocr_results(all_results)


# ============================================================
# OCR LINE GROUPING
# ============================================================
def group_ocr_lines(items):
    prepared = []

    for item in items:
        geo = get_box_geometry(item.get("box"))

        if not geo:
            continue

        x1, y1, x2, y2 = geo

        prepared.append(
            {
                **item,
                "x1": x1,
                "y1": y1,
                "x2": x2,
                "y2": y2,
                "cx": (x1 + x2) / 2,
                "cy": (y1 + y2) / 2,
                "height": max(y2 - y1, 1),
            }
        )

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
        normalize_ocr_text(item.get("text", ""))
        for item in line
        if normalize_ocr_text(item.get("text", ""))
    )


# ============================================================
# DATE EXTRACTION
# ============================================================
def extract_date(text):
    text = normalize_text(text)

    patterns = [
        r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?",
        r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})",
        r"(\d{4})年(\d{1,2})月(\d{1,2})日",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text):
            groups = match.groups()

            try:
                if len(groups[0]) == 4:
                    year = int(groups[0])
                    month = int(groups[1])
                    day = int(groups[2])
                else:
                    day = int(groups[0])
                    month = int(groups[1])
                    year = int(groups[2])

                if 1 <= month <= 12 and 1 <= day <= 31:
                    return f"{year:04d}-{month:02d}-{day:02d}"
            except Exception:
                pass

    return ""


# ============================================================
# NUMBER CLEANING
# ============================================================
def clean_number(text):
    if text is None:
        return ""

    text = str(text).strip()
    text = text.replace("，", ".").replace(",", ".")
    text = text.replace("O", "0").replace("o", "0")

    match = re.search(r"-?\d+(?:\.\d+)?", text)

    if not match:
        return text

    try:
        val = float(match.group(0))
        if val.is_integer():
            return int(val)
        return val
    except Exception:
        return text


# ============================================================
# HEADER DETECTION
# ============================================================
HEADER_KEYWORDS = [
    "部门",
    "开几台",
    "开机",
    "机台",
    "正式",
    "临时",
    "备注",
    "姓名",
    "班次",
    "bộ phận",
    "số máy",
    "máy",
    "chính thức",
    "thời vụ",
    "ghi chú",
    "stt",
    "tên",
    "họ tên",
]


def detect_header_line(lines):
    best_idx = None
    best_score = 0.0

    for idx, line in enumerate(lines):
        text = line_to_text(line).lower()

        if not text:
            continue

        score = 0.0

        for kw in HEADER_KEYWORDS:
            if kw.lower() in text:
                score += max(1.0, len(kw) / 4)

        if len(line) >= 3:
            score += 1.5

        if score > best_score:
            best_score = score
            best_idx = idx

    return best_idx


# ============================================================
# COLUMN CLASSIFICATION
# ============================================================
def classify_column_dynamic(col_name_lower):
    name = normalize_text(col_name_lower).lower()

    if any(
        k in name
        for k in [
            "备注",
            "ghi chú",
            "chú thích",
            "remark",
            "note",
        ]
    ):
        return "remark"

    if any(
        k in name
        for k in [
            "正式工",
            "正式",
            "chính thức",
            "cố định",
            "formal",
        ]
    ):
        return "formal"

    if any(
        k in name
        for k in [
            "临时工",
            "临时",
            "thời vụ",
            "lao động phụ",
            "temp",
        ]
    ):
        return "temp"

    if any(
        k in name
        for k in [
            "开几台机",
            "开机",
            "机台",
            "机器",
            "số máy",
            "số máy mở",
            "máy",
            "machine",
        ]
    ):
        return "machines"

    if any(
        k in name
        for k in [
            "部门",
            "bộ phận",
            "phòng ban",
            "phòng",
            "xưởng",
            "车间",
            "department",
        ]
    ):
        return "dept_src"

    if any(
        k in name
        for k in [
            "tên bộ phận",
            "tên phòng",
            "tên xưởng",
            "department name",
        ]
    ):
        return "dept_src"

    return "unknown"


def infer_column_boundaries(columns):
    if not columns:
        return []

    columns = sorted(columns, key=lambda x: x["x"])
    boundaries = []

    for i, col in enumerate(columns):
        if i == 0:
            left = float("-inf")
        else:
            left = (columns[i - 1]["x"] + col["x"]) / 2

        if i == len(columns) - 1:
            right = float("inf")
        else:
            right = (col["x"] + columns[i + 1]["x"]) / 2

        boundaries.append(
            {
                **col,
                "left": left,
                "right": right,
            }
        )

    return boundaries


def find_column_by_x(boundaries, x):
    if not boundaries:
        return None

    for col in boundaries:
        if col["left"] <= x < col["right"]:
            return col

    return min(
        boundaries,
        key=lambda c: abs(c["x"] - x),
    )


# ============================================================
# STT
# ============================================================
def extract_stt(text):
    text = normalize_text(text)

    patterns = [
        r"^\s*(\d+)[.)]?\s*$",
        r"^\s*(\d+)",
    ]

    for pattern in patterns:
        match = re.match(pattern, text)

        if match:
            try:
                return int(match.group(1))
            except Exception:
                pass

    return None


# ============================================================
# PARSE ATTENDANCE ROWS
# ============================================================
def parse_attendance_rows(lines, header_index):
    if not lines:
        return []

    if header_index is None:
        header_index = 0

    if header_index >= len(lines):
        header_index = 0

    header_line = lines[header_index]
    columns = []

    for item in header_line:
        name = normalize_ocr_text(item.get("text", "")).lower()

        if not name:
            continue

        columns.append(
            {
                "name": name,
                "x": item["cx"],
                "type": classify_column_dynamic(name),
            }
        )

    columns.sort(key=lambda x: x["x"])
    boundaries = infer_column_boundaries(columns)
    rows = []

    for line in lines[header_index + 1:]:
        if not line:
            continue

        line_text = line_to_text(line)

        if not line_text:
            continue

        lower_line = line_text.lower()

        if any(
            kw in lower_line
            for kw in [
                "一共",
                "总计",
                "合计",
                "tổng cộng",
                "total",
            ]
        ):
            continue

        if detect_header_line([line]) is not None:
            header_score = sum(
                1
                for kw in HEADER_KEYWORDS
                if kw.lower() in lower_line
            )

            if header_score >= 2:
                continue

        stt = None
        if line:
            stt = extract_stt(line[0].get("text", ""))

        if stt is None:
            stt = extract_stt(line_text)

        if stt is None:
            stt = len(rows) + 1

        row = {
            "stt": stt,
            "dept_src": "",
            "dept_tgt": "",
            "machines": "",
            "formal": "",
            "temp": "",
            "remark": "",
        }

        unknown_items = []

        for item in line:
            text = normalize_ocr_text(item.get("text", ""))

            if not text:
                continue

            col = find_column_by_x(boundaries, item["cx"])

            if col is None:
                unknown_items.append(text)
                continue

            col_type = col["type"]

            if col_type == "dept_src":
                row["dept_src"] = (
                    row["dept_src"] + " " + text
                ).strip()
            elif col_type == "machines":
                row["machines"] = (
                    row["machines"] + " " + text
                ).strip()
            elif col_type == "formal":
                row["formal"] = (
                    row["formal"] + " " + text
                ).strip()
            elif col_type == "temp":
                row["temp"] = (
                    row["temp"] + " " + text
                ).strip()
            elif col_type == "remark":
                row["remark"] = (
                    row["remark"] + " " + text
                ).strip()
            else:
                unknown_items.append(text)

        if not row["dept_src"] and unknown_items:
            row["dept_src"] = " ".join(unknown_items)
        elif unknown_items and not row["remark"]:
            row["remark"] = " ".join(unknown_items)

        rows.append(row)

    return rows


# ============================================================
# EXCEL GENERATOR & DOWNLOAD
# ============================================================
def create_excel_file(parsed_data, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng Chấm Công Song Ngữ"

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = [
        "STT", 
        "Bộ phận (Gốc)", 
        "Bộ phận (Dịch)", 
        "Số máy", 
        "Chính thức", 
        "Thời vụ", 
        "Ghi chú"
    ]
    
    ws.append(headers)

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = thin_border

    for row in parsed_data:
        dept_tgt = row["dept_tgt"]
        if not dept_tgt and row["dept_src"]:
            dept_tgt = translate_text(row["dept_src"], mode)

        row_data = [
            row["stt"],
            row["dept_src"],
            dept_tgt,
            clean_number(row["machines"]),
            clean_number(row["formal"]),
            clean_number(row["temp"]),
            row["remark"]
        ]
        ws.append(row_data)

    for row_idx in range(2, len(parsed_data) + 2):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [1, 4, 5, 6]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max(max_length + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
