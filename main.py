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

ARGOS_PACKAGES = {
    "zh_en": "translate-zh_en-1_9.argosmodel",
    "en_zh": "translate-en_zh-1_9.argosmodel",
    "vi_en": "translate-vi_en-1_9.argosmodel",
    "en_vi": "translate-en_vi-1_9.argosmodel",
}

ARGOS_BASE_URL = "https://data.argosopentech.com/argospm/v1/"


# ============================================================
# TEXT UTILITIES & LOGIC HỆ THỐNG
# ============================================================
def normalize_text(text):
    if text is None:
        return ""
    text = str(text).replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalize_ocr_text(text):
    text = normalize_text(text)
    if not text:
        return ""
    for src, dst in sorted(OCR_NORMALIZATION_ZH.items(), key=lambda x: len(x[0]), reverse=True):
        text = text.replace(src, dst)
    return text.strip()


def is_number_like(text):
    text = normalize_text(text)
    if not text:
        return False
    return bool(re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", text) or re.fullmatch(r"[\d\s.,:/\\\-+%]+", text))


def should_translate(text):
    text = normalize_text(text)
    if not text or is_number_like(text):
        return False
    return True


def apply_glossary_exact(text, mode):
    text = normalize_text(text)
    if not text:
        return None
    glossary = GLOSSARY_ZH_VI if mode == "Trung ➔ Việt" else GLOSSARY_VI_ZH
    return glossary.get(text)


def load_translation_memory():
    if not MEMORY_FILE.exists():
        return {"zh_vi": {}, "vi_zh": {}}
    try:
        return json.loads(MEMORY_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"zh_vi": {}, "vi_zh": {}}


def save_translation_memory(memory):
    try:
        MEMORY_FILE.write_text(json.dumps(memory, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def memory_lookup(text, mode):
    memory = load_translation_memory()
    key = "zh_vi" if mode == "Trung ➔ Việt" else "vi_zh"
    return memory.get(key, {}).get(normalize_text(text))


def memory_save(source, target, mode):
    source, target = normalize_text(source), normalize_text(target)
    if not source or not target or source == target:
        return
    memory = load_translation_memory()
    key = "zh_vi" if mode == "Trung ➔ Việt" else "vi_zh"
    if source not in memory[key]:
        memory[key][source] = target
        save_translation_memory(memory)


def find_argos_translation(from_code, to_code):
    try:
        langs = argostranslate.translate.get_installed_languages()
        f_lang = next((l for l in langs if l.code == from_code), None)
        t_lang = next((l for l in langs if l.code == to_code), None)
        if f_lang and t_lang:
            return f_lang.get_translation(t_lang)
    except Exception:
        pass
    return None


def translate_text(text, mode, context=None):
    text = normalize_ocr_text(text)
    if not text or not should_translate(text):
        return text

    exact = apply_glossary_exact(text, mode)
    if exact is not None:
        return exact

    mem = memory_lookup(text, mode)
    if mem:
        return mem

    from_c, to_c = ("zh", "vi") if mode == "Trung ➔ Việt" else ("vi", "zh")
    trans = find_argos_translation(from_c, to_c)
    if trans:
        try:
            res = normalize_text(trans.translate(text))
            if res:
                memory_save(text, res, mode)
                return res
        except Exception:
            pass
    return text


def resize_for_ocr(image):
    if image.mode != "RGB":
        image = image.convert("RGB")
    w, h = image.size
    scale = 2.0 if w < 1800 else (1.5 if w < 3000 else 1.0)
    if scale != 1.0:
        image = image.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
    return image


def create_ocr_variants(image):
    base = resize_for_ocr(image)
    gray = ImageOps.grayscale(base)
    return [("original", base), ("gray", gray), ("sharp", gray.filter(ImageFilter.SHARPEN))]


@st.cache_resource
def get_ocr():
    if easyocr is None:
        raise RuntimeError("Chưa cài EasyOCR.")
    return easyocr.Reader(["ch_sim", "en"], gpu=False, verbose=False)


def get_box_geometry(box):
    try:
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        return min(xs), min(ys), max(xs), max(ys)
    except Exception:
        return None


def ocr_image(image):
    reader = get_ocr()
    all_results = []
    for _, variant in create_ocr_variants(image):
        try:
            results = reader.readtext(np.array(variant), detail=1, text_threshold=0.45)
            for bbox, text, prob in results:
                t = normalize_ocr_text(text)
                if t and float(prob) >= 0.25:
                    all_results.append({"text": t, "box": bbox})
        except Exception:
            continue
    return all_results


def group_ocr_lines(items):
    prepared = []
    for item in items:
        geo = get_box_geometry(item.get("box"))
        if geo:
            x1, y1, x2, y2 = geo
            prepared.append({"text": item["text"], "x1": x1, "y1": y1, "x2": x2, "y2": y2, "cx": (x1 + x2) / 2, "cy": (y1 + y2) / 2})
    prepared.sort(key=lambda x: (x["cy"], x["x1"]))
    lines, current_line = [], []
    last_cy = None
    for item in prepared:
        if last_cy is None or abs(item["cy"] - last_cy) <= 15:
            current_line.append(item)
        else:
            current_line.sort(key=lambda x: x["x1"])
            lines.append(current_line)
            current_line = [item]
        last_cy = item["cy"]
    if current_line:
        current_line.sort(key=lambda x: x["x1"])
        lines.append(current_line)
    return lines


def line_to_text(line):
    return " ".join(item["text"] for item in line)


def detect_header_line(lines):
    for idx, line in enumerate(lines):
        txt = line_to_text(line).lower()
        if any(kw in txt for kw in ["部门", "正式", "bộ phận", "chính thức", "số máy"]):
            return idx
    return 0


def classify_column_dynamic(name):
    name = name.lower()
    if any(k in name for k in ["备注", "ghi chú", "remark"]): return "remark"
    if any(k in name for k in ["正式", "chính thức", "formal"]): return "formal"
    if any(k in name for k in ["临时", "thời vụ", "temp"]): return "temp"
    if any(k in name for k in ["机", "máy", "machine"]): return "machines"
    if any(k in name for k in ["部门", "bộ phận", "department"]): return "dept_src"
    return "unknown"


def parse_attendance_rows(lines, header_index):
    if not lines:
        return []
    header_line = lines[header_index if header_index < len(lines) else 0]
    cols = [{"type": classify_column_dynamic(i["text"]), "x": i["cx"]} for i in header_line]
    
    rows = []
    for line in lines[header_index + 1:]:
        txt = line_to_text(line)
        if not txt or any(kw in txt for kw in ["总计", "合计", "tổng cộng"]):
            continue
        
        row = {"stt": len(rows) + 1, "dept_src": "", "dept_tgt": "", "machines": "", "formal": "", "temp": "", "remark": ""}
        for item in line:
            col = min(cols, key=lambda c: abs(c["x"] - item["cx"])) if cols else {"type": "dept_src"}
            ctype = col["type"]
            if ctype in row:
                row[ctype] = (row[ctype] + " " + item["text"]).strip()
            else:
                row["dept_src"] = (row["dept_src"] + " " + item["text"]).strip()
        rows.append(row)
    return rows


def create_excel_file(parsed_data, mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bảng Chấm Công"

    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    headers = ["STT", "Bộ phận (Gốc)", "Bộ phận (Dịch)", "Số máy", "Chính thức", "Thời vụ", "Ghi chú"]
    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in parsed_data:
        dept_tgt = row["dept_tgt"] or translate_text(row["dept_src"], mode)
        ws.append([row["stt"], row["dept_src"], dept_tgt, row["machines"], row["formal"], row["temp"], row["remark"]])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============================================================
# GIAO DIỆN CHÍNH (UI)
# ============================================================
uploaded_file = st.file_uploader("📂 Tải lên hình ảnh bảng chấm công", type=["png", "jpg", "jpeg"])

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption="Ảnh bảng chấm công gốc", use_container_width=True)

    mode = st.selectbox("🌐 Chọn chiều dịch:", ["Trung ➔ Việt", "Việt ➔ Trung"])

    if st.button("🚀 Bắt đầu OCR & Xử lý Dữ liệu", type="primary"):
        with st.spinner("Đang nhận diện văn bản và xử lý bảng..."):
            ocr_results = ocr_image(image)
            lines = group_ocr_lines(ocr_results)
            header_idx = detect_header_line(lines)
            parsed_rows = parse_attendance_rows(lines, header_idx)

            if parsed_rows:
                st.session_state["parsed_rows"] = parsed_rows
                st.session_state["mode"] = mode
                st.success(f"✅ Đã quét thành công {len(parsed_rows)} dòng dữ liệu!")
            else:
                st.warning("⚠️ Không tìm thấy bảng dữ liệu hợp lệ trong ảnh.")

    if "parsed_rows" in st.session_state and st.session_state["parsed_rows"]:
        st.subheader("📊 Kết quả trích xuất:")
        
        display_data = []
        for r in st.session_state["parsed_rows"]:
            t_dept = r["dept_tgt"] or translate_text(r["dept_src"], st.session_state["mode"])
            display_data.append({
                "STT": r["stt"],
                "Bộ phận gốc": r["dept_src"],
                "Bộ phận dịch": t_dept,
                "Số máy": r["machines"],
                "Chính thức": r["formal"],
                "Thời vụ": r["temp"],
                "Ghi chú": r["remark"]
            })
        
        st.dataframe(display_data, use_container_width=True)

        excel_data = create_excel_file(st.session_state["parsed_rows"], st.session_state["mode"])
        
        st.markdown("---")
        st.download_button(
            label="📥 Tải xuống File Excel Kết Quả",
            data=excel_data,
            file_name="bang_cham_cong_song_ngu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
