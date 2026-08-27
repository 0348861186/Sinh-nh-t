import io
import re
import time
import math
import urllib.request
from pathlib import Path

import streamlit as st
import openpyxl
import numpy as np

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# THƯ VIỆN DỊCH LOCAL - KHÔNG DÙNG GEMINI / AI API
# ============================================================

import argostranslate.package
import argostranslate.translate

# ============================================================
# OCR (SỬ DỤNG EASYOCR THAY CHO PADDLEOCR)
# ============================================================

from PIL import Image, ImageEnhance, ImageFilter

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
    text = re.sub(r'[ \t]+', ' ', text)

    return text.strip()


# ============================================================
# ARGOS - TÌM MODEL ĐÃ CÀI
# ============================================================

@st.cache_resource
def get_argos_languages():
    try:
        return argostranslate.translate.get_installed_languages()
    except Exception:
        return []


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


# ============================================================
# TẢI MODEL ARGOS TỰ ĐỘNG
# ============================================================

def download_argos_package(package_name):
    cache_dir = Path(".argos_models")
    cache_dir.mkdir(exist_ok=True)
    target = cache_dir / package_name

    if target.exists() and target.stat().st_size > 0:
        return target

    url = ARGOS_BASE_URL + package_name

    try:
        st.info(f"⬇️ Đang tải model dịch local: {package_name}")
        urllib.request.urlretrieve(url, str(target))
        return target
    except Exception as e:
        if target.exists():
            try:
                target.unlink()
            except Exception:
                pass
        raise RuntimeError(f"Không thể tải model Argos:\n{url}\n\nLỗi: {e}")


def install_argos_package(from_code, to_code):
    package_key = f"{from_code}_{to_code}"
    if package_key not in ARGOS_PACKAGES:
        raise RuntimeError(f"Chưa cấu hình model Argos cho {from_code} -> {to_code}")

    package_name = ARGOS_PACKAGES[package_key]
    package_path = download_argos_package(package_name)

    try:
        argostranslate.package.install_from_path(str(package_path))
    except Exception as e:
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
        translation = find_argos_translation(from_code, to_code)
        if translation is None:
            install_argos_package(from_code, to_code)
            try:
                get_argos_languages.clear()
            except Exception:
                pass
            translation = find_argos_translation(from_code, to_code)

        if translation is None:
            raise RuntimeError(f"Không khởi tạo được model {from_code} -> {to_code}")
        status.append(f"{from_code} → {to_code}")

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

    translation = find_argos_translation(from_code, to_code)
    if translation is None:
        install_argos_package(from_code, to_code)
        try:
            get_argos_languages.clear()
        except Exception:
            pass
        translation = find_argos_translation(from_code, to_code)

    if translation is None:
        raise RuntimeError(f"Không tìm thấy bộ dịch {from_code} → {to_code}")

    try:
        result = translation.translate(text)
        return result.strip() if result else text
    except Exception:
        return text


def translate_text(text, mode):
    text = normalize_text(text)
    if not text:
        return ""

    if mode == "Trung ➔ Việt":
        direct = find_argos_translation("zh", "vi")
        if direct is not None:
            return direct.translate(text).strip()
        english = translate_direct(text, "zh", "en")
        vietnamese = translate_direct(english, "en", "vi")
        return vietnamese.strip()
    else:
        direct = find_argos_translation("vi", "zh")
        if direct is not None:
            return direct.translate(text).strip()
        english = translate_direct(text, "vi", "en")
        chinese = translate_direct(english, "en", "zh")
        return chinese.strip()


# ============================================================
# TIỀN XỬ LÝ ẢNH & EASYOCR
# ============================================================

def preprocess_image(image):
    if image.mode != "RGB":
        image = image.convert("RGB")
    width, height = image.size
    scale = 1 if width >= 1600 else 2
    if scale > 1:
        image = image.resize((width * scale, height * scale), Image.Resampling.LANCZOS)
    image = ImageEnhance.Contrast(image).enhance(1.25)
    image = ImageEnhance.Sharpness(image).enhance(1.25)
    return image


@st.cache_resource
def get_ocr():
    if easyocr is None:
        raise RuntimeError("Chưa cài đặt thư viện EasyOCR.")
    # EasyOCR yêu cầu kết hợp ch_sim với en để tránh lỗi tương thích
    return easyocr.Reader(['ch_sim', 'en'], gpu=False)


def ocr_image(image):
    reader = get_ocr()
    processed = preprocess_image(image)
    img_np = np.array(processed)
    
    results = reader.readtext(img_np)
    items = []
    for (bbox, text, prob) in results:
        text = normalize_text(text)
        if text:
            items.append({
                "text": text,
                "score": float(prob),
                "box": bbox
            })
    return items


def get_box_geometry(box):
    if not box:
        return None
    try:
        xs = [float(point[0]) for point in box]
        ys = [float(point[1]) for point in box]
        return (min(xs), min(ys), max(xs), max(ys))
    except Exception:
        return None


def group_ocr_lines(items):
    prepared = []
    for item in items:
        geometry = get_box_geometry(item.get("box"))
        if geometry is None:
            continue
        x1, y1, x2, y2 = geometry
        prepared.append({
            **item,
            "x1": x1, "y1": y1, "x2": x2, "y2": y2,
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
            tolerance = max(12, avg_h * 0.65)
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
    return " ".join(item["text"].strip() for item in line if item["text"].strip())


def extract_date(text):
    patterns = [
        r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})',
        r'(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})',
        r'(\d{4})年(\d{1,2})月(\d{1,2})日',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            groups = match.groups()
            try:
                if len(groups) == 3:
                    if len(groups[0]) == 4:
                        return f"{int(groups[0]):04d}-{int(groups[1]):02d}-{int(groups[2]):02d}"
                    else:
                        return f"{int(groups[2]):04d}-{int(groups[1]):02d}-{int(groups[0]):02d}"
            except Exception:
                pass
    return ""


def clean_number(text):
    if text is None:
        return ""
    text = str(text).strip().replace(",", ".").replace("，", ".").replace("。", ".")
    match = re.search(r'-?\d+(?:\.\d+)?', text)
    if not match:
        return text
    try:
        val = float(match.group(0))
        return int(val) if val.is_integer() else val
    except Exception:
        return text


def detect_header_line(lines):
    keywords = ["部门", "开几台机", "正式工", "临时工", "备注", "bộ phận", "số máy", "chính thức", "thời vụ", "ghi chú", "stt"]
    best_index = None
    best_score = 0
    for idx, line in enumerate(lines):
        text = line_to_text(line).lower()
        score = sum(1 for kw in keywords if kw in text)
        if score > best_score:
            best_score = score
            best_index = idx
    return best_index


def detect_column_centers(header_line):
    return [{"name": item["text"].lower(), "x": item["cx"]} for item in header_line]


def classify_column(x, columns):
    if not columns:
        return None
    nearest = min(columns, key=lambda c: abs(c["x"] - x))
    return nearest["name"]


def parse_attendance_rows(lines, header_index):
    if header_index is None:
        return []

    columns = detect_column_centers(lines[header_index])
    rows = []

    for line in lines[header_index + 1:]:
        if not line:
            continue
        line_text = line_to_text(line)
        if not line_text:
            continue
        if any(kw in line_text.lower() for kw in ["一共", "总计", "合计", "tổng cộng", "total"]):
            continue

        stt = ""
        stt_match = re.match(r'^\s*(\d+)[\.\\)]?\s*$', line[0]["text"])
        if stt_match:
            stt = int(stt_match.group(1))
        else:
            match = re.match(r'^\s*(\d+)', line_text)
            if match:
                stt = int(match.group(1))

        if not stt:
            continue

        row = {"stt": stt, "dept_src": "", "dept_tgt": "", "machines": "", "formal": "", "temp": "", "remark": ""}

        for item in line:
            column_name = classify_column(item["cx"], columns)
            if column_name is None:
                continue
            text = item["text"].strip()
            if not text:
                continue
            lower = column_name.lower()

            if "部门" in lower or "bộ phận" in lower:
                row["dept_src"] = (row["dept_src"] + " " + text).strip()
            elif "开几台" in lower or "số máy" in lower:
                row["machines"] = clean_number(text)
            elif "正式工" in lower or "chính thức" in lower:
                row["formal"] = clean_number(text)
            elif "临时工" in lower or "thời vụ" in lower:
                row["temp"] = clean_number(text)
            elif "备注" in lower or "ghi chú" in lower:
                row["remark"] = (row["remark"] + " " + text).strip()

        if not row["dept_src"]:
            for item in line:
                txt = item["text"].strip()
                if txt and not re.fullmatch(r'\d+(?:\.\d+)?', txt):
                    row["dept_src"] = txt
                    break

        rows.append(row)

    return rows


def parse_attendance_image(image, mode):
    items = ocr_image(image)
    if not items:
        raise RuntimeError("EasyOCR không nhận diện được chữ trong ảnh.")

    lines = group_ocr_lines(items)
    if not lines:
        raise RuntimeError("Không tạo được dòng OCR từ ảnh.")

    header_index = detect_header_line(lines)
    all_text = "\n".join(line_to_text(line) for line in lines)
    date_str = extract_date(all_text)

    title_lines = [line_to_text(line) for line in lines[:header_index]] if header_index is not None else []
    title_src = " ".join([t for t in title_lines if t]).strip()

    rows = parse_attendance_rows(lines, header_index)
    title_tgt = translate_text(title_src, mode) if title_src else ""

    for row in rows:
        if row.get("dept_src"):
            row["dept_tgt"] = translate_text(row["dept_src"], mode)

    return {"title_src": title_src, "title_tgt": title_tgt, "date_str": date_str, "rows": rows}


def pdf_to_images(pdf_bytes):
    if fitz is None:
        raise RuntimeError("Chưa cài PyMuPDF.")
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    try:
        for page in document:
            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
            images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
    finally:
        document.close()
    return images


def merge_parsed_documents(documents, mode):
    if not documents:
        return {"title_src": "", "title_tgt": "", "date_str": "", "rows": []}
    first = documents[0]
    merged = {"title_src": first.get("title_src", ""), "title_tgt": first.get("title_tgt", ""), "date_str": first.get("date_str", ""), "rows": []}
    next_stt = 1
    for doc in documents:
        for row in doc.get("rows", []):
            new_row = dict(row)
            new_row["stt"] = next_stt
            merged["rows"].append(new_row)
            next_stt += 1
    return merged


def build_excel_from_json(data, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    font_name = "Microsoft YaHei"

    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin_side = Side(style="thin", color="000000")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    t_src = data.get("title_src", "")
    t_tgt = data.get("title_tgt", "")
    dt_str = data.get("date_str", "")
    rows = data.get("rows", [])

    top_title = t_src if mode == "Trung ➔ Việt" else t_tgt
    bot_title = t_tgt if mode == "Trung ➔ Việt" else t_src
    full_title = f"{dt_str} {top_title}\n{bot_title} ngày {dt_str}".strip()

    ws.merge_cells("A1:F1")
    ws["A1"] = full_title
    ws["A1"].font = Font(name=font_name, size=13, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    headers = [
        ("STT", "STT"),
        ("部门" if mode == "Trung ➔ Việt" else "Bộ phận", "Bộ phận" if mode == "Trung ➔ Việt" else "部门"),
        ("开几台机" if mode == "Trung ➔ Việt" else "Số máy mở", "Số máy mở" if mode == "Trung ➔ Việt" else "开几台机"),
        ("正式工" if mode == "Trung ➔ Việt" else "Chính thức", "Chính thức" if mode == "Trung ➔ Việt" else "正式工"),
        ("临时工" if mode == "Trung ➔ Việt" else "Thời vụ", "Thời vụ" if mode == "Trung ➔ Việt" else "临时工"),
        ("备注" if mode == "Trung ➔ Việt" else "Ghi chú", "Ghi chú" if mode == "Trung ➔ Việt" else "备注")
    ]

    for col_idx, (top_h, bot_h) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = f"{top_h}\n{bot_h}" if top_h != bot_h else top_h
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border

    ws.row_dimensions[2].height = 38

    current_row = 3
    for row in rows:
        stt = row.get("stt", "")
        d_src = str(row.get("dept_src", "")) if row.get("dept_src") else ""
        d_tgt = str(row.get("dept_tgt", "")) if row.get("dept_tgt") else ""
        mac = row.get("machines", "") or ""
        fml = row.get("formal", "") or ""
        tmp = row.get("temp", "") or ""
        rmk = str(row.get("remark", "")) if row.get("remark") else ""

        ws.cell(row=current_row, column=1, value=stt)
        ws.cell(row=current_row, column=2, value=f"{d_src}\n{d_tgt}".strip())
        ws.cell(row=current_row, column=3, value=mac)
        ws.cell(row=current_row, column=4, value=fml)
        ws.cell(row=current_row, column=5, value=tmp)
        ws.cell(row=current_row, column=6, value=rmk)

        for col_idx in range(1, 7):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(
                horizontal="center" if col_idx in [1, 3, 4, 5] else "left",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        ws.row_dimensions[current_row].height = 32
        current_row += 1

    column_widths = {'A': 8, 'B': 30, 'C': 12, 'D': 12, 'E': 12, 'F': 20}
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# GIAO DIỆN CHÍNH STREAMLIT
# ============================================================

mode = st.sidebar.selectbox("Chọn chế độ dịch", ["Trung ➔ Việt", "Việt ➔ Trung"])

uploaded_file = st.file_uploader(
    "Tải lên tệp ảnh (PNG, JPG) hoặc PDF bảng chấm công",
    type=["png", "jpg", "jpeg", "pdf"]
)

if uploaded_file is not None:
    try:
        with st.spinner("Đang khởi tạo hệ thống dịch..."):
            initialize_translation_models()

        images = []
        if uploaded_file.type == "application/pdf":
            with st.spinner("Đang chuyển đổi PDF sang ảnh..."):
                images = pdf_to_images(uploaded_file.getvalue())
        else:
            images = [Image.open(uploaded_file)]

        parsed_docs = []
        for img in images:
            with st.spinner("Đang chạy OCR và trích xuất dữ liệu..."):
                parsed_docs.append(parse_attendance_image(img, mode))

        final_data = merge_parsed_documents(parsed_docs, mode)

        st.success("Trích xuất và dịch thành công!")
        
        st.subheader("Bản xem trước dữ liệu")
        st.json(final_data)

        excel_data = build_excel_from_json(final_data, mode)
        st.download_button(
            label="📥 Tải xuống file Excel song ngữ",
            data=excel_data,
            file_name="bang_cham_cong_song_ngu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Đã xảy ra lỗi: {e}")
