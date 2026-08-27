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
# THƯ VIỆN DỊCH LOCAL & OCR
# ============================================================

import argostranslate.package
import argostranslate.translate

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
    page_title="Hệ Thống Dịch Bảng Chấm Công Thông Minh",
    page_icon="🌐",
    layout="wide"
)

st.title("🌐 Dịch & Xuất Bảng Chấm Công Song Ngữ (Tùy Biến Toàn Diện - Bản Nâng Cấp)")
st.caption(
    "Tối ưu hóa OCR & Dịch thuật chuyên sâu nhà máy | Local 100% không mất phí | "
    "Tự động ánh xạ cột & Từ điển mở rộng"
)


# ============================================================
# BỘ TỪ ĐIỂN CHUYÊN NGÀNH NHÀ MÁY (MỞ RỘNG)
# ============================================================

GLOSSARY_ZH_VI = {
    # Từ khóa hệ thống chung & Nhân sự
    "正式工": "Công nhân chính thức",
    "临时工": "Công nhân thời vụ",
    "实习生": "Thực tập sinh",
    "离职": "Đã nghỉ việc",
    "请假": "Xin nghỉ phép",
    "旷工": "Nghỉ không phép",
    "部门": "Bộ phận",
    "车间": "Xưởng",
    "姓名": "Họ tên",
    "班次": "Ca làm việc",
    "早班": "Ca sáng",
    "晚班": "Ca đêm",
    "休息": "Nghỉ ngơi",
    "开几台机": "Số máy chạy",
    "备注": "Ghi chú",
    "总计": "Tổng cộng",
    "合计": "Tổng cộng",
    
    # [TÙY CHỈNH]: Bổ sung các tổ/xưởng/máy móc riêng của nhà máy tại đây
    "裁断组": "Tổ cắt",
    "针车组": "Tổ may",
    "包装组": "Tổ đóng gói",
    "品管部": "Phòng QA/QC",
}

def apply_glossary(text, mode):
    if not text:
        return text
    text_cleaned = text.strip()
    
    if mode == "Trung ➔ Việt":
        if text_cleaned in GLOSSARY_ZH_VI:
            return GLOSSARY_ZH_VI[text_cleaned]
    else:
        vi_zh = {v: k for k, v in GLOSSARY_ZH_VI.items()}
        if text_cleaned in vi_zh:
            return vi_zh[text_cleaned]
            
    return None


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


def normalize_text(text):
    if text is None:
        return ""
    text = str(text).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()


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


def download_argos_package(package_name):
    cache_dir = Path(".argos_models")
    cache_dir.mkdir(exist_ok=True)
    target = cache_dir / package_name
    if target.exists() and target.stat().st_size > 0:
        return target
    url = ARGOS_BASE_URL + package_name
    try:
        urllib.request.urlretrieve(url, str(target))
        return target
    except Exception as e:
        if target.exists():
            try: target.unlink()
            except Exception: pass
        raise RuntimeError(f"Không thể tải model Argos: {e}")


def install_argos_package(from_code, to_code):
    package_key = f"{from_code}_{to_code}"
    if package_key not in ARGOS_PACKAGES:
        raise RuntimeError(f"Chưa cấu hình model {from_code} -> {to_code}")
    package_path = download_argos_package(ARGOS_PACKAGES[package_key])
    try:
        argostranslate.package.install_from_path(str(package_path))
    except Exception as e:
        if "already installed" not in str(e).lower():
            raise


@st.cache_resource
def initialize_translation_models():
    required_pairs = [("zh", "en"), ("en", "zh"), ("vi", "en"), ("en", "vi")]
    for from_code, to_code in required_pairs:
        if find_argos_translation(from_code, to_code) is None:
            install_argos_package(from_code, to_code)
            try: get_argos_languages.clear()
            except Exception: pass


def translate_direct(text, from_code, to_code):
    text = normalize_text(text)
    if not text or from_code == to_code: return text
    translation = find_argos_translation(from_code, to_code)
    if translation is None:
        install_argos_package(from_code, to_code)
        try: get_argos_languages.clear()
        except Exception: pass
        translation = find_argos_translation(from_code, to_code)
    if translation is None: return text
    try:
        result = translation.translate(text)
        return result.strip() if result else text
    except Exception:
        return text


def translate_text(text, mode):
    text = normalize_text(text)
    if not text:
        return ""

    # Làm sạch nhẹ ký tự đặc biệt để tra từ điển hiệu quả hơn
    text_cleaned_lookup = re.sub(r'[^\w\s]', '', text).strip()

    # 1. Tra từ điển chuyên ngành trước
    glossary_result = apply_glossary(text_cleaned_lookup, mode)
    if glossary_result is not None:
        return glossary_result

    # 2. Dịch qua model local kết hợp
    if mode == "Trung ➔ Việt":
        direct = find_argos_translation("zh", "vi")
        if direct is not None:
            try:
                res = direct.translate(text).strip()
                if res: return res
            except Exception:
                pass
        english = translate_direct(text, "zh", "en")
        vietnamese = translate_direct(english, "en", "vi")
        return vietnamese.strip() if vietnamese else text
    else:
        direct = find_argos_translation("vi", "zh")
        if direct is not None:
            try:
                res = direct.translate(text).strip()
                if res: return res
            except Exception:
                pass
        english = translate_direct(text, "vi", "en")
        chinese = translate_direct(english, "en", "zh")
        return chinese.strip() if chinese else text


# ============================================================
# OCR & TIỀN XỬ LÝ ẢNH (NÂNG CẤP ĐỘ NÉT & ĐỘ TƯƠNG PHẢN)
# ============================================================

def preprocess_image(image):
    if image.mode != "RGB":
        image = image.convert("RGB")
    
    # Chuyển sang ảnh xám (Grayscale) giúp OCR đọc chữ tốt hơn trên nền giấy nhà máy
    image = image.convert("L")
    
    width, height = image.size
    scale = 2 if width < 2000 else 1  # Phóng lớn ảnh nếu độ phân giải thấp
    if scale > 1:
        image = image.resize((width * scale, height * scale), Image.Resampling.LANCZOS)
        
    # Tăng cường mạnh độ tương phản và độ sắc nét giúp loại bỏ nhiễu nền
    image = ImageEnhance.Contrast(image).enhance(1.8)
    image = ImageEnhance.Sharpness(image).enhance(2.0)
    
    return image.convert("RGB")


@st.cache_resource
def get_ocr():
    if easyocr is None:
        raise RuntimeError("Chưa cài đặt thư viện EasyOCR.")
    # Hỗ trợ cả giản thể, phồn thể và tiếng Anh để nhận diện toàn diện văn bản xưởng
    return easyocr.Reader(['ch_sim', 'ch_tra', 'en'], gpu=False)


def ocr_image(image):
    reader = get_ocr()
    processed = preprocess_image(image)
    results = reader.readtext(np.array(processed))
    items = []
    for (bbox, text, prob) in results:
        text = normalize_text(text)
        if text:
            items.append({"text": text, "score": float(prob), "box": bbox})
    return items


def get_box_geometry(box):
    if not box: return None
    try:
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        return (min(xs), min(ys), max(xs), max(ys))
    except Exception:
        return None


def group_ocr_lines(items):
    prepared = []
    for item in items:
        geo = get_box_geometry(item.get("box"))
        if not geo: continue
        x1, y1, x2, y2 = geo
        prepared.append({**item, "x1": x1, "y1": y1, "x2": x2, "y2": y2, "cx": (x1+x2)/2, "cy": (y1+y2)/2, "height": max(y2-y1, 1)})

    prepared.sort(key=lambda x: (x["cy"], x["x1"]))
    lines = []
    for item in prepared:
        placed = False
        for line in lines:
            avg_y = sum(x["cy"] for x in line) / len(line)
            avg_h = sum(x["height"] for x in line) / len(line)
            if abs(item["cy"] - avg_y) <= max(12, avg_h * 0.7):
                line.append(item)
                placed = True
                break
        if not placed: lines.append([item])
    for line in lines: line.sort(key=lambda x: x["x1"])
    return lines


def line_to_text(line):
    return " ".join(item["text"].strip() for item in line if item["text"].strip())


def extract_date(text):
    patterns = [
        r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})',
        r'(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})',
        r'(\d{4})年(\d{1,2})月(\d{1,2})日',
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            g = m.groups()
            try:
                if len(g[0]) == 4: return f"{int(g[0]):04d}-{int(g[1]):02d}-{int(g[2]):02d}"
                else: return f"{int(g[2]):04d}-{int(g[1]):02d}-{int(g[0]):02d}"
            except Exception: pass
    return ""


def clean_number(text):
    if text is None: return ""
    text = str(text).strip().replace(",", ".").replace("，", ".")
    match = re.search(r'-?\d+(?:\.\d+)?', text)
    if not match: return text
    try:
        val = float(match.group(0))
        return int(val) if val.is_integer() else val
    except Exception:
        return text


# ============================================================
# HỆ THỐNG ĐỘNG - TỰ ĐỘNG NHẬN DIỆN CỘT THÔNG MINH
# ============================================================

def detect_header_line(lines):
    keywords = ["部门", "开几台", "正式", "临时", "备注", "bộ phận", "số máy", "chính thức", "thời vụ", "ghi chú", "stt", "tên", "họ tên"]
    best_idx, best_score = None, 0
    for idx, line in enumerate(lines):
        text = line_to_text(line).lower()
        score = sum(1 for kw in keywords if kw in text)
        if score > best_score:
            best_score = score
            best_idx = idx
    return best_idx


def classify_column_dynamic(col_name_lower):
    if any(k in col_name_lower for k in ["部门", "bộ phận", "phòng", "xưởng", "tên"]):
        return "dept_src"
    if any(k in col_name_lower for k in ["máy", "台", "开机", "số máy"]):
        return "machines"
    if any(k in col_name_lower for k in ["正式", "chính thức", "cố định"]):
        return "formal"
    if any(k in col_name_lower for k in ["临时", "thời vụ", "phụ"]):
        return "temp"
    if any(k in col_name_lower for k in ["备注", "ghi chú", "chú thích"]):
        return "remark"
    return "unknown"


def parse_attendance_rows(lines, header_index):
    if header_index is None:
        header_index = 0

    header_line = lines[header_index]
    columns = [{"name": item["text"].lower(), "x": item["cx"]} for item in header_line]
    
    rows = []
    for line in lines[header_index + 1:]:
        if not line: continue
        line_text = line_to_text(line)
        if not line_text: continue
        if any(kw in line_text.lower() for kw in ["一共", "总计", "合计", "tổng cộng", "total"]):
            continue

        stt = ""
        stt_match = re.match(r'^\s*(\d+)[\.\\)]?\s*$', line[0]["text"])
        if stt_match:
            stt = int(stt_match.group(1))
        else:
            match = re.match(r'^\s*(\d+)', line_text)
            if match: stt = int(match.group(1))
            else: stt = len(rows) + 1

        row = {"stt": stt, "dept_src": "", "dept_tgt": "", "machines": "", "formal": "", "temp": "", "remark": ""}

        for item in line:
            if not columns:
                continue
            nearest_col = min(columns, key=lambda c: abs(c["x"] - item["cx"]))
            col_type = classify_column_dynamic(nearest_col["name"])
            text = item["text"].strip()
            
            if col_type == "dept_src":
                row["dept_src"] = (row["dept_src"] + " " + text).strip()
            elif col_type == "machines":
                row["machines"] = clean_number(text)
            elif col_type == "formal":
                row["formal"] = clean_number(text)
            elif col_type == "temp":
                row["temp"] = clean_number(text)
            elif col_type == "remark":
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
        raise RuntimeError("Không đọc được chữ từ ảnh này.")
    lines = group_ocr_lines(items)
    header_index = detect_header_line(lines)
    all_text = "\n".join(line_to_text(line) for line in lines)
    date_str = extract_date(all_text)

    title_lines = [line_to_text(line) for line in lines[:header_index]] if header_index else []
    title_src = " ".join([t for t in title_lines if t]).strip()
    
    rows = parse_attendance_rows(lines, header_index)
    title_tgt = translate_text(title_src, mode) if title_src else ""

    for row in rows:
        if row.get("dept_src"):
            row["dept_tgt"] = translate_text(row["dept_src"], mode)

    return {"title_src": title_src, "title_tgt": title_tgt, "date_str": date_str, "rows": rows}


def pdf_to_images(pdf_bytes):
    if fitz is None: raise RuntimeError("Chưa cài PyMuPDF.")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    try:
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
            images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
    finally:
        doc.close()
    return images


def merge_parsed_documents(documents, mode):
    if not documents: return {"title_src": "", "title_tgt": "", "date_str": "", "rows": []}
    merged = {"title_src": documents[0].get("title_src", ""), "title_tgt": documents[0].get("title_tgt", ""), "date_str": documents[0].get("date_str", ""), "rows": []}
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

    t_src, t_tgt, dt_str, rows = data.get("title_src", ""), data.get("title_tgt", ""), data.get("date_str", ""), data.get("rows", [])
    full_title = f"{dt_str} {t_src}\n{t_tgt}".strip()

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
        ws.cell(row=current_row, column=1, value=row.get("stt", ""))
        ws.cell(row=current_row, column=2, value=f"{row.get('dept_src','')}\n{row.get('dept_tgt','')}".strip())
        ws.cell(row=current_row, column=3, value=row.get("machines", ""))
        ws.cell(row=current_row, column=4, value=row.get("formal", ""))
        ws.cell(row=current_row, column=5, value=row.get("temp", ""))
        ws.cell(row=current_row, column=6, value=row.get("remark", ""))

        for col_idx in range(1, 7):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center" if col_idx in [1, 3, 4, 5] else "left", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[current_row].height = 32
        current_row += 1

    for col_letter, width in {'A': 8, 'B': 30, 'C': 12, 'D': 12, 'E': 12, 'F': 20}.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# GIAO DIỆN CHÍNH
# ============================================================

mode = st.sidebar.selectbox("Chọn chiều dịch", ["Trung ➔ Việt", "Việt ➔ Trung"])

uploaded_file = st.file_uploader(
    "Tải lên tệp bảng chấm công bất kỳ (Ảnh PNG, JPG hoặc PDF)",
    type=["png", "jpg", "jpeg", "pdf"]
)

if uploaded_file is not None:
    try:
        with st.spinner("Đang chuẩn bị mô hình dịch local..."):
            initialize_translation_models()

        images = []
        if uploaded_file.type == "application/pdf":
            with st.spinner("Đang xử lý file PDF..."):
                images = pdf_to_images(uploaded_file.getvalue())
        else:
            images = [Image.open(uploaded_file)]

        parsed_docs = []
        for img in images:
            with st.spinner("Đang quét OCR thông minh & dịch linh hoạt..."):
                parsed_docs.append(parse_attendance_image(img, mode))

        final_data = merge_parsed_documents(parsed_docs, mode)

        st.success("Xử lý thành công mọi định dạng file!")
        st.subheader("Xem trước kết quả trích xuất")
        st.json(final_data)

        excel_data = build_excel_from_json(final_data, mode)
        st.download_button(
            label="📥 Tải xuống Excel Song Ngữ Chuẩn Xác",
            data=excel_data,
            file_name="bang_cham_cong_xu_ly_dong.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Đã xảy ra lỗi khi xử lý tệp: {e}")
