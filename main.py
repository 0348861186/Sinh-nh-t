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
    page_title="Hệ Thống Dịch Bảng Chấm Công Chuyên Sâu",
    page_icon="🌐",
    layout="wide"
)

st.title("🌐 Dịch & Xuất Bảng Chấm Công Song Ngữ (Tối Ưu Dịch Thuật)")
st.caption(
    "Bộ từ điển chuyên ngành nhà máy tích hợp | Khắc phục triệt để lỗi dịch sai nghĩa"
)


# ============================================================
# BỘ TỪ ĐIỂN CHUYÊN NGÀNH NHÀ MÁY (BẠN CÓ THỂ BỔ SUNG THÊM Ở ĐÂY)
# ============================================================

GLOSSARY_ZH_VI = {
    # Thuật ngữ chấm công & nhân sự nhà máy
    "正式工": "Công nhân chính thức",
    "临时工": "Công nhân thời vụ",
    "部门": "Bộ phận",
    "开几台机": "Số máy chạy",
    "备注": "Ghi chú",
    "总计": "Tổng cộng",
    "合计": "Tổng cộng",
    "姓名": "Họ tên",
    "班次": "Ca làm việc",
    "早班": "Ca sáng",
    "晚班": "Ca đêm",
    "休息": "Nghỉ ngơi",
    "请假": "Xin nghỉ",
    "旷工": "Nghỉ không phép",
    "加班": "Tăng ca",
    "迟到": "Đi muộn",
    "早退": "Về sớm",
    " and ": " và ",
    
    # Các xưởng / bộ phận thường gặp (Ví dụ mẫu, bạn có thể thay đổi/thêm theo xưởng của bạn)
    "车间": "Xưởng sản xuất",
    "1车间": "Xưởng 1",
    "2车间": "Xưởng 2",
    "品管部": "Phòng Quản lý chất lượng (QC)",
    "货仓": "Kho hàng",
    "包装部": "Bộ phận đóng gói",
    "裁断部": "Bộ phận cắt",
}

def smart_translate(text, mode):
    if not text:
        return ""
    text_cleaned = text.strip()
    
    # 1. Tra cứu chính xác trong từ điển nhà máy
    if mode == "Trung ➔ Việt":
        if text_cleaned in GLOSSARY_ZH_VI:
            return GLOSSARY_ZH_VI[text_cleaned]
        
        # Thử tìm kiếm từng phần nếu chuỗi dài chứa từ khóa trong từ điển
        translated = text_cleaned
        for zh, vi in sorted(GLOSSARY_ZH_VI.items(), key=lambda x: len(x[0]), reverse=True):
            if zh in translated:
                translated = translated.replace(zh, vi)
        if translated != text_cleaned:
            return translated
    else:
        vi_zh = {v: k for k, v in GLOSSARY_ZH_VI.items()}
        if text_cleaned in vi_zh:
            return vi_zh[text_cleaned]
            
        translated = text_cleaned
        for vi, zh in sorted(vi_zh.items(), key=lambda x: len(x[0]), reverse=True):
            if vi in translated:
                translated = translated.replace(vi, zh)
        if translated != text_cleaned:
            return translated

    # 2. Nếu không có trong từ điển, tiến hành dịch qua Argos Translate
    try:
        from_code, to_code = ("zh", "vi") if mode == "Trung ➔ Việt" else ("vi", "zh")
        direct = find_argos_translation(from_code, to_code)
        if direct is not None:
            res = direct.translate(text_cleaned).strip()
            if res:
                return res
        
        # Fallback qua trung gian tiếng Anh nếu thiếu cặp trực tiếp
        mid_code = "en"
        t1 = translate_direct(text_cleaned, from_code, mid_code)
        t2 = translate_direct(t1, mid_code, to_code)
        return t2.strip() if t2 else text_cleaned
    except Exception:
        return text_cleaned


# ============================================================
# CẤU HÌNH MODEL ARGOS & OCR
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
    return re.sub(r'[ \t]+', ' ', text).strip()

@st.cache_resource
def get_argos_languages():
    try:
        return argostranslate.translate.get_installed_languages()
    except Exception:
        return []

def find_argos_translation(from_code, to_code):
    for lang in get_argos_languages():
        if lang.code == from_code:
            for target in get_argos_languages():
                if target.code == to_code:
                    try:
                        return lang.get_translation(target)
                    except Exception:
                        pass
    return None

def download_argos_package(package_name):
    cache_dir = Path(".argos_models")
    cache_dir.mkdir(exist_ok=True)
    target = cache_dir / package_name
    if target.exists() and target.stat().st_size > 0:
        return target
    try:
        urllib.request.urlretrieve(ARGOS_BASE_URL + package_name, str(target))
        return target
    except Exception as e:
        if target.exists():
            try: target.unlink()
            except Exception: pass
        raise RuntimeError(f"Không thể tải model: {e}")

def install_argos_package(from_code, to_code):
    key = f"{from_code}_{to_code}"
    if key in ARGOS_PACKAGES:
        path = download_argos_package(ARGOS_PACKAGES[key])
        try:
            argostranslate.package.install_from_path(str(path))
        except Exception:
            pass

@st.cache_resource
def initialize_translation_models():
    for f, t in [("zh", "en"), ("en", "zh"), ("vi", "en"), ("en", "vi")]:
        if find_argos_translation(f, t) is None:
            install_argos_package(f, t)
            try: get_argos_languages.clear()
            except Exception: pass

def translate_direct(text, from_code, to_code):
    text = normalize_text(text)
    if not text or from_code == to_code:
        return text
    tr = find_argos_translation(from_code, to_code)
    if tr is None:
        install_argos_package(from_code, to_code)
        try: get_argos_languages.clear()
        except Exception: pass
        tr = find_argos_translation(from_code, to_code)
    if tr is None:
        return text
    try:
        res = tr.translate(text)
        return res.strip() if res else text
    except Exception:
        return text

def preprocess_image(image):
    if image.mode != "RGB":
        image = image.convert("RGB")
    w, h = image.size
    scale = 1 if w >= 1600 else 2
    if scale > 1:
        image = image.resize((w * scale, h * scale), Image.Resampling.LANCZOS)
    image = ImageEnhance.Contrast(image).enhance(1.3)
    image = ImageEnhance.Sharpness(image).enhance(1.3)
    return image

@st.cache_resource
def get_ocr():
    if easyocr is None:
        raise RuntimeError("Chưa cài đặt EasyOCR.")
    return easyocr.Reader(['ch_sim', 'en'], gpu=False)

def ocr_image(image):
    reader = get_ocr()
    results = reader.readtext(np.array(preprocess_image(image)))
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

def detect_header_line(lines):
    keywords = ["部门", "开几台", "正式", "临时", "备注", "bộ phận", "số máy", "chính thức", "thời vụ", "ghi chú", "stt", "tên"]
    best_idx, best_score = None, 0
    for idx, line in enumerate(lines):
        text = line_to_text(line).lower()
        score = sum(1 for kw in keywords if kw in text)
        if score > best_score:
            best_score = score
            best_idx = idx
    return best_idx

def classify_column_dynamic(name):
    if any(k in name for k in ["部门", "bộ phận", "phòng", "xưởng", "tên"]): return "dept_src"
    if any(k in name for k in ["máy", "台", "开机", "số máy"]): return "machines"
    if any(k in name for k in ["正式", "chính thức"]): return "formal"
    if any(k in name for k in ["临时", "thời vụ"]): return "temp"
    if any(k in name for k in ["备注", "ghi chú"]): return "remark"
    return "unknown"

def parse_attendance_rows(lines, header_index):
    if header_index is None: header_index = 0
    columns = [{"name": item["text"].lower(), "x": item["cx"]} for item in lines[header_index]]
    rows = []
    
    for line in lines[header_index + 1:]:
        if not line: continue
        line_text = line_to_text(line)
        if not line_text or any(kw in line_text.lower() for kw in ["一共", "总计", "合计", "tổng cộng"]):
            continue

        stt = len(rows) + 1
        m_stt = re.match(r'^\s*(\d+)[\.\\)]?\s*$', line[0]["text"])
        if m_stt: stt = int(m_stt.group(1))

        row = {"stt": stt, "dept_src": "", "dept_tgt": "", "machines": "", "formal": "", "temp": "", "remark": ""}
        for item in line:
            if not columns: continue
            col = min(columns, key=lambda c: abs(c["x"] - item["cx"]))
            c_type = classify_column_dynamic(col["name"])
            txt = item["text"].strip()
            
            if c_type == "dept_src": row["dept_src"] = (row["dept_src"] + " " + txt).strip()
            elif c_type == "machines": row["machines"] = clean_number(txt)
            elif c_type == "formal": row["formal"] = clean_number(txt)
            elif c_type == "temp": row["temp"] = clean_number(txt)
            elif c_type == "remark": row["remark"] = (row["remark"] + " " + txt).strip()

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
    if not items: raise RuntimeError("Không đọc được chữ từ ảnh.")
    lines = group_ocr_lines(items)
    h_idx = detect_header_line(lines)
    all_text = "\n".join(line_to_text(l) for l in lines)
    date_str = extract_date(all_text)

    title_lines = [line_to_text(l) for l in lines[:h_idx]] if h_idx else []
    title_src = " ".join([t for t in title_lines if t]).strip()
    
    rows = parse_attendance_rows(lines, h_idx)
    title_tgt = smart_translate(title_src, mode) if title_src else ""

    for row in rows:
        if row.get("dept_src"):
            row["dept_tgt"] = smart_translate(row["dept_src"], mode)

    return {"title_src": title_src, "title_tgt": title_tgt, "date_str": date_str, "rows": rows}

def pdf_to_images(pdf_bytes):
    if fitz is None: raise RuntimeError("Chưa cài PyMuPDF.")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    imgs = []
    try:
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
            imgs.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
    finally:
        doc.close()
    return imgs

def merge_parsed_documents(documents):
    if not documents: return {"title_src": "", "title_tgt": "", "date_str": "", "rows": []}
    merged = {"title_src": documents[0].get("title_src", ""), "title_tgt": documents[0].get("title_tgt", ""), "date_str": documents[0].get("date_str", ""), "rows": []}
    next_stt = 1
    for doc in documents:
        for row in doc.get("rows", []):
            r = dict(row)
            r["stt"] = next_stt
            merged["rows"].append(r)
            next_stt += 1
    return merged

def build_excel_from_json(data, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    font_name = "Microsoft YaHei"

    orange_fill = PatternFill(fill_type="solid", fgColor="ED7D00")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

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

    for idx, (th, bh) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=idx)
        cell.value = f"{th}\n{bh}" if th != bh else th
        cell.font = Font(name=font_name, size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = orange_fill
        cell.border = border
    ws.row_dimensions[2].height = 38

    curr = 3
    for row in rows:
        ws.cell(row=curr, column=1, value=row.get("stt", ""))
        ws.cell(row=curr, column=2, value=f"{row.get('dept_src','')}\n{row.get('dept_tgt','')}".strip())
        ws.cell(row=curr, column=3, value=row.get("machines", ""))
        ws.cell(row=curr, column=4, value=row.get("formal", ""))
        ws.cell(row=curr, column=5, value=row.get("temp", ""))
        ws.cell(row=curr, column=6, value=row.get("remark", ""))

        for c in range(1, 7):
            cell = ws.cell(row=curr, column=c)
            cell.font = Font(name=font_name, size=10)
            cell.alignment = Alignment(horizontal="center" if c in [1, 3, 4, 5] else "left", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[curr].height = 32
        curr += 1

    for col, w in {'A': 8, 'B': 30, 'C': 12, 'D': 12, 'E': 12, 'F': 20}.items():
        ws.column_dimensions[col].width = w

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# GIAO DIỆN
# ============================================================

mode = st.sidebar.selectbox("Chọn chiều dịch", ["Trung ➔ Việt", "Việt ➔ Trung"])
uploaded_file = st.file_uploader("Tải lên tệp bảng chấm công (Ảnh / PDF)", type=["png", "jpg", "jpeg", "pdf"])

if uploaded_file is not None:
    try:
        with st.spinner("Đang khởi tạo hệ thống dịch..."):
            initialize_translation_models()

        images = pdf_to_images(uploaded_file.getvalue()) if uploaded_file.type == "application/pdf" else [Image.open(uploaded_file)]
        
        docs = []
        for img in images:
            with st.spinner("Đang xử lý thông minh..."):
                docs.append(parse_attendance_image(img, mode))

        final_data = merge_parsed_documents(docs)
        st.success("Xử lý thành công!")
        st.json(final_data)

        st.download_button(
            label="📥 Tải xuống Excel Song Ngữ",
            data=build_excel_from_json(final_data, mode),
            file_name="bang_cham_cong_chuan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Lỗi: {e}")
