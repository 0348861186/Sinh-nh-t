import os, io, requests
import pandas as pd
from datetime import datetime

# Github Actions sẽ tự động bơm 2 biến này vào
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
DB_FILE = "current_hr_data.xlsx" # Đường dẫn file Excel trên Github

def normalize(text): return str(text).strip().lower().replace("_", " ").replace("-", " ")

def find_col(cols, keywords):
    norm_cols = {c: normalize(c) for c in cols}
    for c, n in norm_cols.items():
        if n in keywords: return c
    for c, n in norm_cols.items():
        for k in keywords:
            if k in n: return c
    return None

def send_msg(text):
    requests.post(f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage", data={"chat_id": TELEGRAM_CHAT_ID, "text": text})

def send_file(df, name, caption):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w: df.to_excel(w, index=False)
    out.seek(0)
    requests.post(f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument", data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption}, files={"document": (name, out)})

def main():
    if not os.path.exists(DB_FILE):
        print("Không tìm thấy file Excel!")
        return

    df = pd.read_excel(DB_FILE)
    cols = df.columns
    
    # Tìm cột tự động (Fallback AI)
    c_ma = find_col(cols, ["mã nv", "ma nv", "employee id"])
    c_ten = find_col(cols, ["họ và tên", "ho va ten", "họ tên"])
    c_ns = find_col(cols, ["ngày sinh", "ngay sinh", "ns"])
    c_bp = find_col(cols, ["bộ phận", "bo phan", "phòng ban"])
    c_hd = find_col(cols, ["ngày hết hạn hợp đồng", "ngay het han hop dong"])

    if not all([c_ma, c_ten, c_ns, c_bp, c_hd]):
        print("File excel sai chuẩn cột.")
        return

    # Ép kiểu dữ liệu
    df[c_ns] = pd.to_datetime(df[c_ns], errors="coerce", dayfirst=True)
    df[c_hd] = pd.to_datetime(df[c_hd], errors="coerce", dayfirst=True)

    now = datetime.now() # Github Actions chạy giờ UTC, nên bạn có thể set cron bù giờ
    month, year = now.month, now.year

    bday = df[df[c_ns].dt.month == month].copy()
    contract = df[(df[c_hd].dt.month == month) & (df[c_hd].dt.year == year)].copy()

    msg = f"📊 BÁO CÁO NHÂN SỰ {month:02d}/{year}\n🎂 Sinh nhật: {len(bday)}\n📄 HĐ đến hạn: {len(contract)}\n⏰ Gửi tự động từ Github Actions"
    send_msg(msg)
    if not bday.empty: send_file(bday, f"sinh_nhat_T{month}.xlsx", f"🎂 Sinh nhật T{month}")
    if not contract.empty: send_file(contract, f"hop_dong_T{month}.xlsx", f"📄 Hợp đồng T{month}")
    print("XONG!")

if __name__ == "__main__":
    main()
