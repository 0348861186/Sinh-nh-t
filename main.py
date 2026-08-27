import io
import re
import time
import math
import urllib.request
from pathlib import Path

import streamlit as st
import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# THƯ VIỆN DỊCH LOCAL - KHÔNG DÙNG GEMINI / AI API
# ============================================================

import argostranslate.package
import argostranslate.translate

# ============================================================
# OCR
# ============================================================

from PIL import Image, ImageEnhance, ImageFilter

try:
    from paddleocr import PaddleOCR
except Exception:
    PaddleOCR = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None


# ============================================================
# CẤU HÌNH STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Hệ Thống Dịch Bảng Chấm Công Hai Chiều",
    page_icon="🌐",
    layout="wide"
)

st.title("🌐 Dịch & Xuất Bảng Chấm Công Song Ngữ")
st.caption(
    "Không dùng Gemini/API quota | Hỗ trợ Trung ➔ Việt và Việt ➔ Trung | "
    "Excel, Ảnh, PDF | OCR local + dịch local"
)


# ============================================================
# CẤU HÌNH MODEL ARGOS
# ============================================================

ARGOS_PACKAGES = {
    "zh_en": "translate-zh_en-1_9.argosmodel",
    "en_zh": "translate-en_zh-1_9.argosmodel",
    "vi_en": "translate-vi_en-1_9.argosmodel",
    "en_vi": "translate-en_vi-1_9.argosmodel",
}

ARGOS_BASE_URL = "https://data.argosopentech.com/argospm/v1/"


# ============================================================
# KIỂM TRA CHUỖI
# ============================================================

def has_chinese(text):
    if not text:
        return False

    return bool(
        re.search(
            r'[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]',
            str(text)
        )
    )


def has_vietnamese(text):
    if not isinstance(text, str):
        return False

    vietnamese_pattern = (
        r'[àáảãạâầấẩẫậăằắẳẵặ'
        r'èéẻẽẹêềếểễệ'
        r'ìíỉĩị'
        r'òóỏõọôồốổỗộơờớởỡợ'
        r'ùúủũụưừứửữự'
        r'ỳýỷỹỵđĐ]'
    )

    return bool(
        re.search(
            vietnamese_pattern,
            text,
            re.IGNORECASE
        )
    )


# ============================================================
# CHUẨN HÓA TEXT
# ============================================================

def normalize_text(text):
    if text is None:
        return ""

    text = str(text)

    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    # Gom nhiều khoảng trắng nhưng không phá xuống dòng
    text = re.sub(r'[ \t]+', ' ', text)

    return text.strip()


# ============================================================
# ARGOS - TÌM MODEL ĐÃ CÀI
# ============================================================

@st.cache_resource
def get_argos_languages():
    """
    Lấy danh sách ngôn ngữ/model Argos đang có trên máy.
    """
    try:
        return argostranslate.translate.get_installed_languages()
    except Exception:
        return []


def find_argos_translation(from_code, to_code):
    """
    Tìm translation trực tiếp.
    """

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
        translation = from_lang.get_translation(to_lang)
        return translation
    except Exception:
        return None


# ============================================================
# TẢI MODEL ARGOS TỰ ĐỘNG
# ============================================================

def download_argos_package(package_name):
    """
    Tải model Argos nếu chưa có.
    """

    cache_dir = Path(".argos_models")
    cache_dir.mkdir(exist_ok=True)

    target = cache_dir / package_name

    if target.exists() and target.stat().st_size > 0:
        return target

    url = ARGOS_BASE_URL + package_name

    try:
        st.info(
            f"⬇️ Đang tải model dịch local: {package_name}"
        )

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
            f"Không thể tải model Argos:\n{url}\n\nLỗi: {e}"
        )


def install_argos_package(from_code, to_code):
    """
    Cài model Argos cho một cặp ngôn ngữ.
    """

    package_key = f"{from_code}_{to_code}"

    if package_key not in ARGOS_PACKAGES:
        raise RuntimeError(
            f"Chưa cấu hình model Argos cho {from_code} -> {to_code}"
        )

    package_name = ARGOS_PACKAGES[package_key]

    package_path = download_argos_package(package_name)

    try:
        argostranslate.package.install_from_path(
            str(package_path)
        )
    except Exception as e:

        # Có thể model đã được cài từ trước
        if "already installed" not in str(e).lower():
            raise


# ============================================================
# KHỞI TẠO CÁC MODEL DỊCH
# ============================================================

@st.cache_resource
def initialize_translation_models():

    required_pairs = [
        ("zh", "en"),
        ("en", "zh"),
        ("vi", "en"),
        ("en", "vi"),
    ]

    status = []

    for from_code, to_code in required_pairs:

        translation = find_argos_translation(
            from_code,
            to_code
        )

        if translation is None:

            install_argos_package(
                from_code,
                to_code
            )

            # Xóa cache resource ngôn ngữ để đọc lại model mới
            try:
                get_argos_languages.clear()
            except Exception:
                pass

            translation = find_argos_translation(
                from_code,
                to_code
            )

        if translation is None:
            raise RuntimeError(
                f"Không khởi tạo được model "
                f"{from_code} -> {to_code}"
            )

        status.append(
            f"{from_code} → {to_code}"
        )

    return status


# ============================================================
# DỊCH LOCAL
# ============================================================

def translate_direct(text, from_code, to_code):

    text = normalize_text(text)

    if not text:
        return ""

    if from_code == to_code:
        return text

    translation = find_argos_translation(
        from_code,
        to_code
    )

    if translation is None:

        install_argos_package(
            from_code,
            to_code
        )

        try:
            get_argos_languages.clear()
        except Exception:
            pass

        translation = find_argos_translation(
            from_code,
            to_code
        )

    if translation is None:
        raise RuntimeError(
            f"Không tìm thấy bộ dịch {from_code} → {to_code}"
        )

    try:
        result = translation.translate(text)

        if result:
            return result.strip()

        return text

    except Exception:
        return text


def translate_text(text, mode):

    text = normalize_text(text)

    if not text:
        return ""

    # Trung -> Việt
    if mode == "Trung ➔ Việt":

        # Argos không cần gọi API.
        # Nếu có cặp trực tiếp thì dùng trực tiếp.
        direct = find_argos_translation("zh", "vi")

        if direct is not None:
            return direct.translate(text).strip()

        # Pivot: Trung -> Anh -> Việt
        english = translate_direct(
            text,
            "zh",
            "en"
        )

        vietnamese = translate_direct(
            english,
            "en",
            "vi"
        )

        return vietnamese.strip()

    # Việt -> Trung
    else:

        direct = find_argos_translation("vi", "zh")

        if direct is not None:
            return direct.translate(text).strip()

        # Pivot: Việt -> Anh -> Trung
        english = translate_direct(
            text,
            "vi",
            "en"
        )

        chinese = translate_direct(
            english,
            "en",
            "zh"
        )

        return chinese.strip()


# ============================================================
# DỊCH NHIỀU TEXT
# ============================================================

def translate_texts(texts, mode):

    result = {}

    total = len(texts)

    progress = st.progress(0)

    for index, text in enumerate(texts, start=1):

        try:
            result[text] = translate_text(
                text,
                mode
            )
        except Exception as e:
            result[text] = ""

        progress.progress(
            min(index / max(total, 1), 1.0)
        )

    progress.empty()

    return result


# ============================================================
# TIỀN XỬ LÝ ẢNH
# ============================================================

def preprocess_image(image):

    if image.mode != "RGB":
        image = image.convert("RGB")

    # Phóng to ảnh nếu ảnh nhỏ
    width, height = image.size

    scale = 1

    if width < 1600:
        scale = 2

    if scale > 1:
        image = image.resize(
            (
                width * scale,
                height * scale
            ),
            Image.Resampling.LANCZOS
        )

    # Tăng contrast nhẹ
    image = ImageEnhance.Contrast(
        image
    ).enhance(1.25)

    # Tăng sharpness
    image = ImageEnhance.Sharpness(
        image
    ).enhance(1.25)

    return image


# ============================================================
# KHỞI TẠO PADDLE OCR
# ============================================================

@st.cache_resource
def get_ocr():

    if PaddleOCR is None:
        raise RuntimeError(
            "Chưa cài PaddleOCR/PaddlePaddle."
        )

    """
    PP-OCRv5 multilingual.
    Dùng lang=vi để ưu tiên tiếng Việt.
    PaddleOCR vẫn nhận được nhiều ký tự Latin.
    """

    try:

        ocr = PaddleOCR(
            lang="vi",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=True
        )

        return ocr

    except Exception:

        # Fallback cho một số phiên bản PaddleOCR
        ocr = PaddleOCR(
            lang="vi"
        )

        return ocr


# ============================================================
# CHUYỂN KẾT QUẢ PADDLE OCR THÀNH DANH SÁCH
# ============================================================

def parse_paddle_result(result):

    items = []

    try:

        for res in result:

            data = None

            # PaddleOCR 3.x
            if hasattr(res, "json"):

                try:
                    obj = res.json()

                    if isinstance(obj, str):
                        import json
                        obj = json.loads(obj)

                    data = obj.get("res", obj)

                except Exception:
                    data = None

            # Trường hợp object có thuộc tính res
            if data is None and hasattr(res, "res"):

                try:
                    data = res.res
                except Exception:
                    data = None

            if not data:
                continue

            texts = data.get("rec_texts", [])
            scores = data.get("rec_scores", [])
            polys = data.get("rec_polys", [])

            for i, text in enumerate(texts):

                text = normalize_text(text)

                if not text:
                    continue

                score = 1.0

                if i < len(scores):
                    try:
                        score = float(scores[i])
                    except Exception:
                        pass

                box = None

                if i < len(polys):
                    try:
                        box = polys[i].tolist()
                    except Exception:
                        box = polys[i]

                items.append(
                    {
                        "text": text,
                        "score": score,
                        "box": box
                    }
                )

    except Exception:
        pass

    return items


# ============================================================
# OCR MỘT ẢNH
# ============================================================

def ocr_image(image):

    ocr = get_ocr()

    processed = preprocess_image(image)

    temp_path = Path(".ocr_temp.png")

    processed.save(
        temp_path,
        format="PNG"
    )

    try:

        result = ocr.predict(
            str(temp_path)
        )

        items = parse_paddle_result(
            result
        )

        return items

    finally:

        try:
            temp_path.unlink()
        except Exception:
            pass


# ============================================================
# TÍNH TỌA ĐỘ BOX
# ============================================================

def get_box_geometry(box):

    if not box:
        return None

    try:

        xs = [
            float(point[0])
            for point in box
        ]

        ys = [
            float(point[1])
            for point in box
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
# GOM OCR THÀNH DÒNG
# ============================================================

def group_ocr_lines(items):

    prepared = []

    for item in items:

        geometry = get_box_geometry(
            item.get("box")
        )

        if geometry is None:
            continue

        x1, y1, x2, y2 = geometry

        prepared.append(
            {
                **item,
                "x1": x1,
                "y1": y1,
                "x2": x2,
                "y2": y2,
                "cx": (x1 + x2) / 2,
                "cy": (y1 + y2) / 2,
                "height": max(y2 - y1, 1)
            }
        )

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
                x["cy"] for x in line
            ) / len(line)

            avg_h = sum(
                x["height"] for x in line
            ) / len(line)

            tolerance = max(
                12,
                avg_h * 0.65
            )

            if abs(item["cy"] - avg_y) <= tolerance:

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


# ============================================================
# GỘP TEXT TRONG DÒNG
# ============================================================

def line_to_text(line):

    texts = []

    for item in line:

        text = item["text"].strip()

        if text:
            texts.append(text)

    return " ".join(texts).strip()


# ============================================================
# TÌM NGÀY
# ============================================================

def extract_date(text):

    patterns = [

        # YYYY-MM-DD
        r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})',

        # DD-MM-YYYY
        r'(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})',

        # YYYY年MM月DD日
        r'(\d{4})年(\d{1,2})月(\d{1,2})日',

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text
        )

        if match:

            groups = match.groups()

            try:

                if len(groups) == 3:

                    # Nếu phần đầu là năm
                    if len(groups[0]) == 4:

                        year = int(groups[0])
                        month = int(groups[1])
                        day = int(groups[2])

                    else:

                        day = int(groups[0])
                        month = int(groups[1])
                        year = int(groups[2])

                    return (
                        f"{year:04d}-{month:02d}-{day:02d}"
                    )

            except Exception:
                pass

    return ""


# ============================================================
# NHẬN DIỆN SỐ
# ============================================================

def clean_number(text):

    if text is None:
        return ""

    text = str(text).strip()

    text = text.replace(",", ".")
    text = text.replace("，", ".")
    text = text.replace("。", ".")

    # Chỉ còn số, dấu chấm, dấu âm
    match = re.search(
        r'-?\d+(?:\.\d+)?',
        text
    )

    if not match:
        return text

    number = match.group(0)

    try:

        value = float(number)

        if value.is_integer():
            return int(value)

        return value

    except Exception:

        return number


# ============================================================
# XÁC ĐỊNH DÒNG HEADER
# ============================================================

def detect_header_line(lines):

    keywords = [
        "部门",
        "开几台机",
        "正式工",
        "临时工",
        "备注",
        "bộ phận",
        "số máy",
        "chính thức",
        "thời vụ",
        "ghi chú",
        "stt"
    ]

    best_index = None
    best_score = 0

    for idx, line in enumerate(lines):

        text = line_to_text(
            line
        ).lower()

        score = 0

        for keyword in keywords:

            if keyword.lower() in text:
                score += 1

        if score > best_score:

            best_score = score
            best_index = idx

    return best_index


# ============================================================
# XÁC ĐỊNH CỘT DỰA TRÊN HEADER
# ============================================================

def detect_column_centers(header_line):

    columns = []

    for item in header_line:

        text = item["text"].lower()

        columns.append(
            {
                "name": text,
                "x": item["cx"]
            }
        )

    return columns


def classify_column(x, columns):

    if not columns:
        return None

    nearest = min(
        columns,
        key=lambda c: abs(c["x"] - x)
    )

    return nearest["name"]


# ============================================================
# TẠO ROW TỪ DÒNG OCR
# ============================================================

def parse_attendance_rows(
    lines,
    header_index
):

    if header_index is None:
        return []

    header_line = lines[header_index]

    columns = detect_column_centers(
        header_line
    )

    rows = []

    for line in lines[header_index + 1:]:

        if not line:
            continue

        line_text = line_to_text(
            line
        )

        if not line_text:
            continue

        # Bỏ dòng tổng cộng
        if any(
            keyword in line_text.lower()
            for keyword in [
                "一共",
                "总计",
                "合计",
                "tổng cộng",
                "total"
            ]
        ):
            continue

        # Lấy STT
        stt = ""

        first_text = line[0]["text"]

        stt_match = re.match(
            r'^\s*(\d+)[\.\、\)]?\s*$',
            first_text
        )

        if stt_match:
            stt = int(
                stt_match.group(1)
            )

        # Nếu không thấy STT thì thử toàn dòng
        if not stt:

            match = re.match(
                r'^\s*(\d+)',
                line_text
            )

            if match:
                stt = int(
                    match.group(1)
                )

        if not stt:
            # Có thể đã đến phần khác của tài liệu
            continue

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

            x = item["cx"]
            text = item["text"].strip()

            if not text:
                continue

            column_name = classify_column(
                x,
                columns
            )

            if column_name is None:
                continue

            lower = column_name.lower()

            if (
                "部门" in lower
                or "bộ phận" in lower
            ):

                if row["dept_src"]:
                    row["dept_src"] += " " + text
                else:
                    row["dept_src"] = text

            elif (
                "开几台" in lower
                or "số máy" in lower
            ):

                row["machines"] = clean_number(
                    text
                )

            elif (
                "正式工" in lower
                or "chính thức" in lower
            ):

                row["formal"] = clean_number(
                    text
                )

            elif (
                "临时工" in lower
                or "thời vụ" in lower
            ):

                row["temp"] = clean_number(
                    text
                )

            elif (
                "备注" in lower
                or "ghi chú" in lower
            ):

                if row["remark"]:
                    row["remark"] += " " + text
                else:
                    row["remark"] = text

        # Nếu OCR không xác định được cột,
        # dùng fallback theo vị trí text.
        if not row["dept_src"]:

            candidate_texts = []

            for item in line:

                txt = item["text"].strip()

                if not txt:
                    continue

                if re.fullmatch(
                    r'\d+(?:\.\d+)?',
                    txt
                ):
                    continue

                candidate_texts.append(txt)

            if candidate_texts:

                row["dept_src"] = (
                    candidate_texts[0]
                )

        rows.append(row)

    return rows


# ============================================================
# DỊCH DEPARTMENT
# ============================================================

def translate_attendance_rows(
    rows,
    mode
):

    for row in rows:

        source_department = str(
            row.get("dept_src", "")
        ).strip()

        if not source_department:
            continue

        translated = translate_text(
            source_department,
            mode
        )

        row["dept_tgt"] = translated

    return rows


# ============================================================
# PHÂN TÍCH ẢNH
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
            "PaddleOCR không nhận diện được chữ trong ảnh."
        )

    lines = group_ocr_lines(
        items
    )

    if not lines:
        raise RuntimeError(
            "Không tạo được dòng OCR từ ảnh."
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

    # Lấy title phía trên header
    title_lines = []

    if header_index is not None:

        for line in lines[:header_index]:

            txt = line_to_text(line)

            if txt:
                title_lines.append(txt)

    title_src = (
        " ".join(title_lines).strip()
        if title_lines
        else ""
    )

    rows = parse_attendance_rows(
        lines,
        header_index
    )

    # Dịch title
    title_tgt = ""

    if title_src:

        try:
            title_tgt = translate_text(
                title_src,
                mode
            )
        except Exception:
            title_tgt = ""

    # Dịch department
    rows = translate_attendance_rows(
        rows,
        mode
    )

    return {
        "title_src": title_src,
        "title_tgt": title_tgt,
        "date_str": date_str,
        "rows": rows
    }


# ============================================================
# PDF -> DANH SÁCH ẢNH
# ============================================================

def pdf_to_images(pdf_bytes):

    if fitz is None:

        raise RuntimeError(
            "Chưa cài PyMuPDF. "
            "Hãy cài pymupdf."
        )

    images = []

    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:

        for page_index in range(
            len(document)
        ):

            page = document[
                page_index
            ]

            # Zoom 2.5x để OCR rõ hơn
            matrix = fitz.Matrix(
                2.5,
                2.5
            )

            pix = page.get_pixmap(
                matrix=matrix,
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

        document.close()

    return images


# ============================================================
# GỘP KẾT QUẢ NHIỀU TRANG PDF
# ============================================================

def merge_parsed_documents(
    documents,
    mode
):

    if not documents:
        return {
            "title_src": "",
            "title_tgt": "",
            "date_str": "",
            "rows": []
        }

    first = documents[0]

    merged = {
        "title_src": first.get(
            "title_src",
            ""
        ),
        "title_tgt": first.get(
            "title_tgt",
            ""
        ),
        "date_str": first.get(
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
# BUILD EXCEL TỪ JSON
# GIỮ NGUYÊN LOGIC CODE GỐC
# ============================================================

def build_excel_from_json(
    data,
    mode
):

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

    t_src = data.get(
        "title_src",
        ""
    )

    t_tgt = data.get(
        "title_tgt",
        ""
    )

    dt_str = data.get(
        "date_str",
        ""
    )

    rows = data.get(
        "rows",
        []
    )

    if mode == "Trung ➔ Việt":

        top_title = t_src
        bot_title = t_tgt

    else:

        top_title = t_tgt
        bot_title = t_src

    full_title = (
        f"{dt_str} "
        f"{top_title}\n"
        f"{bot_title} ngày {dt_str}"
    ).strip()

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

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

    ws.row_dimensions[1].height = 42

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

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

    for col_idx, (
        top_h,
        bot_h
    ) in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            row=2,
            column=col_idx
        )

        if top_h != bot_h:

            cell.value = (
                f"{top_h}\n{bot_h}"
            )

        else:

            cell.value = top_h

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

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    current_row = 3

    total_workers = 0

    for row in rows:

        stt = row.get(
            "stt",
            ""
        )

        d_src = (
            str(
                row.get(
                    "dept_src",
                    ""
                )
            )
            if row.get("dept_src")
            else ""
        )

        d_tgt = (
            str(
                row.get(
                    "dept_tgt",
                    ""
                )
            )
            if row.get("dept_tgt")
            else ""
        )

        mac = row.get(
            "machines",
            ""
        ) or ""

        fml = row.get(
            "formal",
            ""
        ) or ""

        tmp = row.get(
            "temp",
            ""
        ) or ""

        rmk = (
            str(
                row.get(
                    "remark",
                    ""
                )
            )
            if row.get("remark")
            else ""
        )

        try:

            if fml:
                total_workers += float(fml)

            if tmp:
                total_workers += float(tmp)

        except (
            ValueError,
            TypeError
        ):

            pass

        ws.cell(
            row=current_row,
            column=1,
            value=stt
        )

        ws.cell(
            row=current_row,
            column=2,
            value=f"{d_src}\n{d_tgt}".strip()
        )

        ws.cell(
            row=current_row,
            column=3,
            value=mac
        )

        ws.cell(
            row=current_row,
            column=4,
            value=fml
        )

        ws.cell(
            row=current_row,
            column=5,
            value=tmp
        )

        ws.cell(
            row=current_row,
            column=6,
            value=rmk
        )

        for col in range(1, 7):

            c = ws.cell(
                row=current_row,
                column=col
            )

            c.font = Font(
                name=font_name,
                size=10
            )

            c.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            c.border = border

        ws.row_dimensions[
            current_row
        ].height = 32

        current_row += 1

    # --------------------------------------------------------
    # TOTAL
    # --------------------------------------------------------

    total_row = current_row

    ws.merge_cells(
        start_row=total_row,
        start_column=1,
        end_row=total_row,
        end_column=2
    )

    if mode == "Trung ➔ Việt":

        total_label = (
            "一共\nTổng cộng"
        )

    else:

        total_label = (
            "Tổng cộng\n一共"
        )

    ws.cell(
        row=total_row,
        column=1,
        value=total_label
    )

    ws.merge_cells(
        start_row=total_row,
        start_column=3,
        end_row=total_row,
        end_column=5
    )

    total_value = (
        int(total_workers)
        if (
            isinstance(
                total_workers,
                float
            )
            and total_workers.is_integer()
        )
        else total_workers
    )

    ws.cell(
        row=total_row,
        column=3,
        value=total_value
    )

    for col in range(1, 7):

        c = ws.cell(
            row=total_row,
            column=col
        )

        c.font = Font(
            name=font_name,
            size=11,
            bold=True
        )

        c.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        c.border = border

    ws.row_dimensions[
        total_row
    ].height = 36

    # --------------------------------------------------------
    # COLUMN WIDTH
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 17
    ws.column_dimensions["E"].width = 17
    ws.column_dimensions["F"].width = 18

    out = io.BytesIO()

    wb.save(out)

    out.seek(0)

    return out


# ============================================================
# KIỂM TRA FILE EXCEL
# ============================================================

def get_excel_texts(
    wb,
    translation_mode
):

    texts_to_translate = set()

    for sheet in wb.worksheets:

        for row in sheet.iter_rows():

            for cell in row:

                if not cell.value:
                    continue

                if not isinstance(
                    cell.value,
                    str
                ):
                    continue

                val = cell.value.strip()

                if not val:
                    continue

                # Bỏ công thức
                if val.startswith("="):
                    continue

                if translation_mode == "Trung ➔ Việt":

                    if has_chinese(val):
                        texts_to_translate.add(
                            val
                        )

                else:

                    if (
                        has_vietnamese(val)
                        or not has_chinese(val)
                    ):

                        if (
                            len(val) > 1
                            and not val.isnumeric()
                        ):

                            texts_to_translate.add(
                                val
                            )

    return list(
        texts_to_translate
    )


# ============================================================
# DỊCH EXCEL
# ============================================================

def translate_excel(
    file_bytes,
    translation_mode
):

    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes)
    )

    texts = get_excel_texts(
        wb,
        translation_mode
    )

    if not texts:
        return None, 0

    translation_dict = translate_texts(
        texts,
        translation_mode
    )

    # --------------------------------------------------------
    # CHÈN BẢN DỊCH
    # GIỮ NGUYÊN LOGIC GỐC
    # --------------------------------------------------------

    for sheet in wb.worksheets:

        for row in sheet.iter_rows():

            for cell in row:

                if not cell.value:
                    continue

                if not isinstance(
                    cell.value,
                    str
                ):
                    continue

                orig = cell.value.strip()

                trans = translation_dict.get(
                    orig,
                    ""
                )

                if not trans:
                    continue

                cell.value = (
                    f"{orig}\n{trans}"
                )

                curr_align = cell.alignment

                cell.alignment = Alignment(
                    horizontal=(
                        curr_align.horizontal
                        or "center"
                    ),
                    vertical=(
                        curr_align.vertical
                        or "center"
                    ),
                    wrap_text=True
                )

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output, len(texts)


# ============================================================
# GIAO DIỆN
# ============================================================

col1, col2 = st.columns(
    [1.2, 2]
)

with col1:

    translation_mode = st.radio(
        "Chế độ dịch:",
        options=[
            "Trung ➔ Việt",
            "Việt ➔ Trung"
        ],
        horizontal=True
    )

with col2:

    uploaded_file = st.file_uploader(
        "Tải lên Ảnh, PDF hoặc File Excel:",
        type=[
            "png",
            "jpg",
            "jpeg",
            "pdf",
            "xlsx"
        ]
    )


# ============================================================
# TRẠNG THÁI HỆ THỐNG
# ============================================================

with st.expander(
    "⚙️ Kiểm tra bộ dịch local",
    expanded=False
):

    st.write(
        "Ứng dụng không sử dụng Gemini API. "
        "Model dịch được chạy local bằng Argos Translate."
    )

    if st.button(
        "🔧 Kiểm tra / tải model dịch"
    ):

        try:

            with st.spinner(
                "Đang kiểm tra model dịch..."
            ):

                status = (
                    initialize_translation_models()
                )

            st.success(
                "Đã sẵn sàng: "
                + ", ".join(status)
            )

        except Exception as e:

            st.error(
                f"Không khởi tạo được model: {e}"
            )


# ============================================================
# XỬ LÝ FILE
# ============================================================

if uploaded_file is not None:

    extension = (
        uploaded_file.name
        .lower()
        .split(".")[-1]
    )

    is_excel = (
        extension == "xlsx"
    )

    if is_excel:

        button_label = (
            f"🚀 Dịch ({translation_mode}) "
            "& Bảo Toàn Format Excel"
        )

    else:

        button_label = (
            f"🚀 OCR Ảnh/PDF & Dịch "
            f"({translation_mode})"
        )

    if st.button(
        button_label,
        use_container_width=True
    ):

        try:

            # =================================================
            # KHỞI ĐỘNG MODEL
            # =================================================

            with st.spinner(
                "🔧 Đang kiểm tra bộ dịch local..."
            ):

                initialize_translation_models()

            # =================================================
            # TRƯỜNG HỢP 1: EXCEL
            # =================================================

            if is_excel:

                with st.spinner(
                    "1️⃣ Đang đọc các ô cần dịch..."
                ):

                    file_bytes = (
                        uploaded_file.read()
                    )

                with st.spinner(
                    f"2️⃣ Đang dịch "
                    f"bằng model local "
                    f"({translation_mode})..."
                ):

                    output, count = translate_excel(
                        file_bytes,
                        translation_mode
                    )

                if output is None:

                    st.warning(
                        "Không tìm thấy nội dung "
                        "văn bản phù hợp với chế độ "
                        "dịch đã chọn!"
                    )

                else:

                    st.success(
                        f"✅ Đã dịch thành công "
                        f"{count} nội dung văn bản "
                        f"bằng thư viện local."
                    )

                    st.download_button(
                        label=(
                            "⬇️ Tải File Excel Song Ngữ (.xlsx)"
                        ),
                        data=output.getvalue(),
                        file_name=(
                            f"Translated_"
                            f"{uploaded_file.name}"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        ),
                        use_container_width=True
                    )

            # =================================================
            # TRƯỜNG HỢP 2: ẢNH
            # =================================================

            elif extension in [
                "png",
                "jpg",
                "jpeg"
            ]:

                with st.spinner(
                    "1️⃣ PaddleOCR đang nhận diện ảnh..."
                ):

                    file_bytes = (
                        uploaded_file.read()
                    )

                    image = Image.open(
                        io.BytesIO(file_bytes)
                    )

                    parsed_data = (
                        parse_attendance_image(
                            image,
                            translation_mode
                        )
                    )

                with st.spinner(
                    "2️⃣ Đang tạo Excel..."
                ):

                    excel_bytes = (
                        build_excel_from_json(
                            parsed_data,
                            translation_mode
                        )
                    )

                st.success(
                    "✅ Đã OCR, dịch và tạo Excel thành công!"
                )

                st.download_button(
                    label="⬇️ Tải File Excel (.xlsx)",
                    data=excel_bytes.getvalue(),
                    file_name=(
                        f"Bang_cham_cong_"
                        f"{parsed_data.get('date_str', 'export')}.xlsx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

                # Hiển thị preview OCR
                with st.expander(
                    "🔎 Xem kết quả OCR"
                ):

                    st.json(
                        parsed_data
                    )

            # =================================================
            # TRƯỜNG HỢP 3: PDF
            # =================================================

            elif extension == "pdf":

                if fitz is None:

                    st.error(
                        "Thiếu PyMuPDF. "
                        "Hãy cài pymupdf."
                    )

                else:

                    with st.spinner(
                        "1️⃣ Đang chuyển PDF thành ảnh..."
                    ):

                        pdf_bytes = (
                            uploaded_file.read()
                        )

                        pages = pdf_to_images(
                            pdf_bytes
                        )

                    st.info(
                        f"PDF có {len(pages)} trang."
                    )

                    documents = []

                    progress = st.progress(
                        0
                    )

                    for index, page_image in enumerate(
                        pages,
                        start=1
                    ):

                        st.write(
                            f"📄 Đang OCR trang "
                            f"{index}/{len(pages)}..."
                        )

                        try:

                            document = (
                                parse_attendance_image(
                                    page_image,
                                    translation_mode
                                )
                            )

                            documents.append(
                                document
                            )

                        except Exception as page_error:

                            st.warning(
                                f"⚠️ Trang {index} "
                                f"không xử lý được: "
                                f"{page_error}"
                            )

                        progress.progress(
                            index / max(
                                len(pages),
                                1
                            )
                        )

                    progress.empty()

                    if not documents:

                        st.error(
                            "Không OCR được dữ liệu nào từ PDF."
                        )

                    else:

                        with st.spinner(
                            "2️⃣ Đang ghép dữ liệu các trang..."
                        ):

                            parsed_data = (
                                merge_parsed_documents(
                                    documents,
                                    translation_mode
                                )
                            )

                        with st.spinner(
                            "3️⃣ Đang tạo Excel..."
                        ):

                            excel_bytes = (
                                build_excel_from_json(
                                    parsed_data,
                                    translation_mode
                                )
                            )

                        st.success(
                            "✅ Đã OCR, dịch và chuyển PDF "
                            "sang Excel thành công!"
                        )

                        st.download_button(
                            label=(
                                "⬇️ Tải File Excel (.xlsx)"
                            ),
                            data=(
                                excel_bytes.getvalue()
                            ),
                            file_name=(
                                f"Bang_cham_cong_"
                                f"{parsed_data.get('date_str', 'export')}.xlsx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.spreadsheetml.sheet"
                            ),
                            use_container_width=True
                        )

                        with st.expander(
                            "🔎 Xem dữ liệu OCR"
                        ):

                            st.json(
                                parsed_data
                            )

        except Exception as e:

            st.error(
                "❌ Xảy ra lỗi trong quá trình xử lý:"
            )

            st.exception(e)
