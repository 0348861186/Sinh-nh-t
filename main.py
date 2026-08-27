
import io
import re
import json
from pathlib import Path

import streamlit as st
import openpyxl
import numpy as np

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from PIL import Image, ImageOps

# ============================================================
# LOCAL OCR
# ============================================================
try:
    import easyocr
except Exception:
    easyocr = None

# ============================================================
# LOCAL TRANSLATION
# ============================================================
try:
    import argostranslate.translate
except Exception:
    argostranslate = None


# ============================================================
# STREAMLIT CONFIG
# ============================================================
st.set_page_config(
    page_title="Dịch Bảng Chấm Công Local",
    page_icon="🌐",
    layout="wide",
)

st.title("🌐 Dịch & Xuất Bảng Chấm Công Song Ngữ")

st.caption(
    "Local 100% | Không AI/API | EasyOCR + Argos Translate | "
    "Từ điển chuyên ngành | Translation Memory"
)


# ============================================================
# LOCAL FILES
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


GLOSSARY_VI_ZH = {
    value: key
    for key, value in GLOSSARY_ZH_VI.items()
}


# ============================================================
# OCR NORMALIZATION
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
# TEXT UTILITIES
# ============================================================
def normalize_text(text):
    if text is None:
        return ""

    text = str(text)

    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    text = text.replace("\u3000", " ")

    text = re.sub(r"[ \t]+", " ", text)

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

    return text.strip()


def is_number_like(text):
    text = normalize_text(text)

    if not text:
        return False

    if re.fullmatch(
        r"[-+]?\d+(?:[.,]\d+)?",
        text
    ):
        return True

    if re.fullmatch(
        r"[\d\s.,:/\\\-+%]+",
        text
    ):
        return True

    return False


def should_translate(text):
    text = normalize_text(text)

    if not text:
        return False

    if is_number_like(text):
        return False

    return True


# ============================================================
# TRANSLATION MEMORY
# ============================================================
@st.cache_data
def load_translation_memory_cached():
    if not MEMORY_FILE.exists():
        return {
            "zh_vi": {},
            "vi_zh": {},
        }

    try:
        data = json.loads(
            MEMORY_FILE.read_text(
                encoding="utf-8"
            )
        )

        if not isinstance(data, dict):
            raise ValueError

        data.setdefault("zh_vi", {})
        data.setdefault("vi_zh", {})

        return data

    except Exception:
        return {
            "zh_vi": {},
            "vi_zh": {},
        }


def save_translation_memory(memory):
    try:
        MEMORY_FILE.write_text(
            json.dumps(
                memory,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        # Xóa cache để lần chạy sau đọc lại file mới
        load_translation_memory_cached.clear()

    except Exception:
        pass


# ============================================================
# ARGOS TRANSLATOR CACHE
# ============================================================
@st.cache_resource
def get_argos_translator(from_code, to_code):
    if argostranslate is None:
        return None

    try:
        languages = (
            argostranslate.translate
            .get_installed_languages()
        )

        from_lang = next(
            (
                lang
                for lang in languages
                if lang.code == from_code
            ),
            None,
        )

        to_lang = next(
            (
                lang
                for lang in languages
                if lang.code == to_code
            ),
            None,
        )

        if not from_lang or not to_lang:
            return None

        return from_lang.get_translation(to_lang)

    except Exception:
        return None


# ============================================================
# TRANSLATION CACHE
# ============================================================
@st.cache_data
def translate_cached(text, mode):
    text = normalize_ocr_text(text)

    if not text:
        return ""

    if not should_translate(text):
        return text

    # --------------------------------------------------------
    # 1. GLOSSARY
    # --------------------------------------------------------
    if mode == "Trung ➔ Việt":
        exact = GLOSSARY_ZH_VI.get(text)
    else:
        exact = GLOSSARY_VI_ZH.get(text)

    if exact:
        return exact

    # --------------------------------------------------------
    # 2. MEMORY
    # --------------------------------------------------------
    memory = load_translation_memory_cached()

    memory_key = (
        "zh_vi"
        if mode == "Trung ➔ Việt"
        else "vi_zh"
    )

    remembered = memory.get(
        memory_key,
        {}
    ).get(text)

    if remembered:
        return remembered

    # --------------------------------------------------------
    # 3. ARGOS
    # --------------------------------------------------------
    if mode == "Trung ➔ Việt":
        from_code = "zh"
        to_code = "vi"
    else:
        from_code = "vi"
        to_code = "zh"

    translator = get_argos_translator(
        from_code,
        to_code,
    )

    if translator:
        try:
            result = translator.translate(text)

            result = normalize_text(result)

            if result:
                return result

        except Exception:
            pass

    # Không có model hoặc dịch thất bại
    return text


def translate_text(text, mode):
    """
    Hàm trung gian để giữ tương thích
    với logic code cũ.
    """
    return translate_cached(
        text,
        mode,
    )


# ============================================================
# OCR
# ============================================================
@st.cache_resource
def get_ocr():
    if easyocr is None:
        raise RuntimeError(
            "Chưa cài EasyOCR. "
            "Hãy cài easyocr."
        )

    return easyocr.Reader(
        ["ch_sim", "en"],
        gpu=False,
        verbose=False,
    )


# ============================================================
# TỐI ƯU ẢNH TRƯỚC OCR
# ============================================================
def prepare_image_for_ocr(image):
    if image.mode != "RGB":
        image = image.convert("RGB")

    width, height = image.size

    # --------------------------------------------------------
    # Giới hạn kích thước ảnh để giảm RAM
    # --------------------------------------------------------
    MAX_WIDTH = 2500
    MAX_HEIGHT = 2500

    scale = min(
        MAX_WIDTH / width,
        MAX_HEIGHT / height,
        1.0,
    )

    if scale < 1.0:
        new_size = (
            int(width * scale),
            int(height * scale),
        )

        image = image.resize(
            new_size,
            Image.Resampling.LANCZOS,
        )

    return image


# ============================================================
# OCR IMAGE
# ============================================================
def ocr_image(image):
    reader = get_ocr()

    image = prepare_image_for_ocr(image)

    try:
        image_array = np.array(image)

        results = reader.readtext(
            image_array,
            detail=1,

            # Giữ độ nhạy tương đối tốt
            text_threshold=0.45,

            # Không chạy nhiều biến thể ảnh
            paragraph=False,
        )

    except Exception as e:
        raise RuntimeError(
            f"OCR thất bại: {e}"
        )

    output = []

    for bbox, text, probability in results:

        text = normalize_ocr_text(text)

        if not text:
            continue

        try:
            probability = float(probability)
        except Exception:
            probability = 0.0

        if probability < 0.25:
            continue

        output.append(
            {
                "text": text,
                "box": bbox,
                "prob": probability,
            }
        )

    return output


# ============================================================
# OCR BOX GEOMETRY
# ============================================================
def get_box_geometry(box):
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
            max(ys),
        )

    except Exception:
        return None


# ============================================================
# GROUP OCR INTO LINES
# ============================================================
def group_ocr_lines(items):
    prepared = []

    for item in items:

        geometry = get_box_geometry(
            item.get("box")
        )

        if not geometry:
            continue

        x1, y1, x2, y2 = geometry

        prepared.append(
            {
                "text": item["text"],
                "x1": x1,
                "y1": y1,
                "x2": x2,
                "y2": y2,

                "cx": (x1 + x2) / 2,
                "cy": (y1 + y2) / 2,

                "height": max(
                    1,
                    y2 - y1,
                ),
            }
        )

    if not prepared:
        return []

    prepared.sort(
        key=lambda x: (
            x["cy"],
            x["x1"],
        )
    )

    lines = []
    current_line = []

    last_cy = None
    last_height = 20

    for item in prepared:

        if last_cy is None:

            current_line = [item]

        else:

            tolerance = max(
                12,
                min(
                    30,
                    last_height * 0.65,
                ),
            )

            if abs(
                item["cy"] - last_cy
            ) <= tolerance:

                current_line.append(item)

            else:

                current_line.sort(
                    key=lambda x: x["x1"]
                )

                lines.append(
                    current_line
                )

                current_line = [item]

        last_cy = item["cy"]
        last_height = item["height"]

    if current_line:

        current_line.sort(
            key=lambda x: x["x1"]
        )

        lines.append(
            current_line
        )

    return lines


def line_to_text(line):
    return " ".join(
        item["text"]
        for item in line
    )


# ============================================================
# HEADER DETECTION
# ============================================================
def detect_header_line(lines):

    if not lines:
        return 0

    keywords = [
        "部门",
        "正式",
        "临时",
        "机台",
        "开机",
        "备注",

        "bộ phận",
        "chính thức",
        "thời vụ",
        "số máy",
        "ghi chú",
    ]

    for index, line in enumerate(lines):

        text = line_to_text(
            line
        ).lower()

        if any(
            keyword in text
            for keyword in keywords
        ):
            return index

    return 0


# ============================================================
# COLUMN CLASSIFICATION
# ============================================================
def classify_column_dynamic(name):

    name = normalize_text(
        name
    ).lower()

    if any(
        keyword in name
        for keyword in [
            "备注",
            "ghi chú",
            "remark",
        ]
    ):
        return "remark"

    if any(
        keyword in name
        for keyword in [
            "正式",
            "chính thức",
            "formal",
        ]
    ):
        return "formal"

    if any(
        keyword in name
        for keyword in [
            "临时",
            "thời vụ",
            "temp",
        ]
    ):
        return "temp"

    if any(
        keyword in name
        for keyword in [
            "机",
            "机器",
            "机台",
            "máy",
            "machine",
        ]
    ):
        return "machines"

    if any(
        keyword in name
        for keyword in [
            "部门",
            "bộ phận",
            "department",
        ]
    ):
        return "dept_src"

    return "unknown"


# ============================================================
# PARSE ATTENDANCE ROWS
# ============================================================
def parse_attendance_rows(
    lines,
    header_index,
):

    if not lines:
        return []

    if header_index >= len(lines):
        header_index = 0

    header_line = lines[
        header_index
    ]

    # --------------------------------------------------------
    # Xác định vị trí cột
    # --------------------------------------------------------
    columns = []

    for item in header_line:

        column_type = (
            classify_column_dynamic(
                item["text"]
            )
        )

        columns.append(
            {
                "type": column_type,
                "x": item["cx"],
            }
        )

    rows = []

    # --------------------------------------------------------
    # Đọc từng dòng sau header
    # --------------------------------------------------------
    for line in lines[
        header_index + 1:
    ]:

        text = line_to_text(
            line
        )

        if not text:
            continue

        lower_text = text.lower()

        # Bỏ dòng tổng
        if any(
            keyword in lower_text
            for keyword in [
                "总计",
                "合计",
                "tổng cộng",
            ]
        ):
            continue

        row = {
            "stt": len(rows) + 1,

            "dept_src": "",
            "dept_tgt": "",

            "machines": "",
            "formal": "",
            "temp": "",
            "remark": "",
        }

        # ----------------------------------------------------
        # Gán từng OCR item vào cột gần nhất
        # ----------------------------------------------------
        for item in line:

            if columns:

                nearest_column = min(
                    columns,
                    key=lambda column:
                    abs(
                        column["x"]
                        - item["cx"]
                    ),
                )

                column_type = (
                    nearest_column["type"]
                )

            else:

                column_type = "dept_src"

            if column_type in row:

                current = row[
                    column_type
                ]

                if current:
                    row[column_type] = (
                        current
                        + " "
                        + item["text"]
                    ).strip()
                else:
                    row[column_type] = (
                        item["text"]
                    )

            else:

                if row["dept_src"]:
                    row["dept_src"] = (
                        row["dept_src"]
                        + " "
                        + item["text"]
                    ).strip()
                else:
                    row["dept_src"] = (
                        item["text"]
                    )

        # ----------------------------------------------------
        # Chỉ thêm dòng có dữ liệu
        # ----------------------------------------------------
        if any(
            row[key]
            for key in [
                "dept_src",
                "machines",
                "formal",
                "temp",
                "remark",
            ]
        ):
            rows.append(row)

    return rows


# ============================================================
# DỊCH TOÀN BỘ DỮ LIỆU
# ============================================================
def translate_rows(
    rows,
    mode,
):

    translated_rows = []

    for row in rows:

        new_row = dict(row)

        source = normalize_ocr_text(
            row.get(
                "dept_src",
                "",
            )
        )

        target = row.get(
            "dept_tgt",
            "",
        )

        if source and not target:

            target = translate_text(
                source,
                mode,
            )

        new_row[
            "dept_src"
        ] = source

        new_row[
            "dept_tgt"
        ] = target

        translated_rows.append(
            new_row
        )

    return translated_rows


# ============================================================
# EXCEL
# ============================================================
def create_excel_file(
    parsed_data,
    mode,
):

    wb = Workbook()

    ws = wb.active

    ws.title = "Bảng Chấm Công"

    # --------------------------------------------------------
    # Style
    # --------------------------------------------------------
    header_fill = PatternFill(
        start_color="4F81BD",
        end_color="4F81BD",
        fill_type="solid",
    )

    header_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="FFFFFF",
    )

    header_alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    headers = [
        "STT",
        "Bộ phận (Gốc)",
        "Bộ phận (Dịch)",
        "Số máy",
        "Chính thức",
        "Thời vụ",
        "Ghi chú",
    ]

    ws.append(headers)

    for column_index in range(
        1,
        len(headers) + 1,
    ):

        cell = ws.cell(
            row=1,
            column=column_index,
        )

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = (
            header_alignment
        )

    # --------------------------------------------------------
    # Data
    # --------------------------------------------------------
    for row in parsed_data:

        source = normalize_text(
            row.get(
                "dept_src",
                "",
            )
        )

        target = normalize_text(
            row.get(
                "dept_tgt",
                "",
            )
        )

        if source and not target:

            target = translate_text(
                source,
                mode,
            )

        ws.append(
            [
                row.get(
                    "stt",
                    "",
                ),

                source,

                target,

                row.get(
                    "machines",
                    "",
                ),

                row.get(
                    "formal",
                    "",
                ),

                row.get(
                    "temp",
                    "",
                ),

                row.get(
                    "remark",
                    "",
                ),
            ]
        )

    # --------------------------------------------------------
    # Width
    # --------------------------------------------------------
    widths = {
        "A": 8,
        "B": 24,
        "C": 28,
        "D": 12,
        "E": 15,
        "F": 15,
        "G": 25,
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    # --------------------------------------------------------
    # Freeze header
    # --------------------------------------------------------
    ws.freeze_panes = "A2"

    # --------------------------------------------------------
    # Output
    # --------------------------------------------------------
    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ============================================================
# XÓA SESSION KHI LOAD FILE MỚI
# ============================================================
def clear_previous_result():

    for key in [
        "parsed_rows",
        "translated_rows",
    ]:

        if key in st.session_state:
            del st.session_state[key]


# ============================================================
# UI
# ============================================================
uploaded_file = st.file_uploader(
    "📂 Tải lên hình ảnh bảng chấm công",
    type=[
        "png",
        "jpg",
        "jpeg",
        "webp",
        "bmp",
    ],
)


if uploaded_file is not None:

    # --------------------------------------------------------
    # Đọc ảnh
    # --------------------------------------------------------
    try:

        image = Image.open(
            uploaded_file
        )

        image.load()

    except Exception as e:

        st.error(
            f"❌ Không thể đọc ảnh: {e}"
        )

        st.stop()

    # --------------------------------------------------------
    # Hiển thị ảnh
    # --------------------------------------------------------
    st.image(
        image,
        caption="Ảnh bảng chấm công gốc",
        use_container_width=True,
    )

    # --------------------------------------------------------
    # Thông tin ảnh
    # --------------------------------------------------------
    width, height = image.size

    st.caption(
        f"📐 Kích thước ảnh: "
        f"{width} × {height}px"
    )

    # --------------------------------------------------------
    # Translation mode
    # --------------------------------------------------------
    mode = st.selectbox(
        "🌐 Chọn chiều dịch:",
        [
            "Trung ➔ Việt",
            "Việt ➔ Trung",
        ],
    )

    # --------------------------------------------------------
    # Process
    # --------------------------------------------------------
    if st.button(
        "🚀 Bắt đầu OCR & Xử lý",
        type="primary",
        use_container_width=True,
    ):

        clear_previous_result()

        # ====================================================
        # OCR
        # ====================================================
        with st.spinner(
            "🔎 Đang nhận diện văn bản..."
        ):

            try:

                ocr_results = ocr_image(
                    image
                )

            except Exception as e:

                st.error(
                    f"❌ Lỗi OCR: {e}"
                )

                st.stop()

        if not ocr_results:

            st.warning(
                "⚠️ OCR không nhận diện được "
                "văn bản trong ảnh."
            )

            st.stop()

        st.success(
            f"✅ OCR nhận diện "
            f"{len(ocr_results)} vùng văn bản."
        )

        # ====================================================
        # GROUP
        # ====================================================
        with st.spinner(
            "📐 Đang phân tích cấu trúc bảng..."
        ):

            lines = group_ocr_lines(
                ocr_results
            )

            header_index = (
                detect_header_line(
                    lines
                )
            )

            parsed_rows = (
                parse_attendance_rows(
                    lines,
                    header_index,
                )
            )

        if not parsed_rows:

            st.warning(
                "⚠️ Không tìm thấy "
                "dòng dữ liệu hợp lệ."
            )

            st.stop()

        # ====================================================
        # TRANSLATION
        # ====================================================
        with st.spinner(
            "🌐 Đang dịch bằng bộ dịch Local..."
        ):

            translated_rows = (
                translate_rows(
                    parsed_rows,
                    mode,
                )
            )

        st.session_state[
            "parsed_rows"
        ] = parsed_rows

        st.session_state[
            "translated_rows"
        ] = translated_rows

        st.session_state[
            "mode"
        ] = mode

        st.success(
            f"✅ Hoàn tất: "
            f"{len(translated_rows)} dòng dữ liệu."
        )


# ============================================================
# DISPLAY RESULT
# ============================================================
if (
    "translated_rows"
    in st.session_state
    and st.session_state[
        "translated_rows"
    ]
):

    rows = st.session_state[
        "translated_rows"
    ]

    mode = st.session_state[
        "mode"
    ]

    st.subheader(
        "📊 Kết quả trích xuất"
    )

    display_data = []

    for row in rows:

        display_data.append(
            {
                "STT": row.get(
                    "stt",
                    "",
                ),

                "Bộ phận gốc": row.get(
                    "dept_src",
                    "",
                ),

                "Bộ phận dịch": row.get(
                    "dept_tgt",
                    "",
                ),

                "Số máy": row.get(
                    "machines",
                    "",
                ),

                "Chính thức": row.get(
                    "formal",
                    "",
                ),

                "Thời vụ": row.get(
                    "temp",
                    "",
                ),

                "Ghi chú": row.get(
                    "remark",
                    "",
                ),
            }
        )

    st.dataframe(
        display_data,
        use_container_width=True,
        hide_index=True,
    )

    # ========================================================
    # EXCEL
    # ========================================================
    excel_data = create_excel_file(
        rows,
        mode,
    )

    st.markdown("---")

    st.download_button(
        label="📥 Tải xuống File Excel Kết Quả",
        data=excel_data,
        file_name=(
            "bang_cham_cong_song_ngu.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )


# ============================================================
# LOCAL STATUS
# ============================================================
with st.expander(
    "⚙️ Kiểm tra hệ thống Local"
):

    if easyocr is None:
        st.error(
            "❌ EasyOCR chưa được cài."
        )
    else:
        st.success(
            "✅ EasyOCR đã sẵn sàng."
        )

    if argostranslate is None:
        st.warning(
            "⚠️ Argos Translate chưa được cài. "
            "Từ điển chuyên ngành vẫn hoạt động."
        )
    else:
        st.success(
            "✅ Argos Translate đã sẵn sàng."
        )

    st.info(
        "💡 Hệ thống ưu tiên theo thứ tự: "
        "Từ điển chuyên ngành → Translation Memory "
        "→ Argos Translate → giữ nguyên văn bản."
    )
