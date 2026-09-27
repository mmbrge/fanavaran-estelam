# -*- coding: utf-8 -*-
"""
calibrate_grid.py
==================
کالیبراسیون گرید نتایج جستجو: مرزهای ستون «شماره الحاقیه» + ارتفاع هر ردیف رو
ثبت می‌کنه، تا grid_reader.py بتونه بعداً با OCR هر ردیف رو بخونه و آخرین ردیفی
که مقدارش 0 هست رو پیدا کنه.

قبل از اجرا:
    - نتیجه یه جستجوی نمونه (با حداقل ۲-۳ ردیف) توی فناوران باز باشه.

اجرا:
    python calibrate_grid.py

مراحل: هاور روی هرکدوم و Enter بزن:
    1) گوشه بالا-چپ سلول «شماره الحاقیه» توی اولین ردیف نتایج
    2) گوشه پایین-راست همون سلول (توی همون ردیف اول)
    3) همون گوشه بالا-چپ ولی توی ردیف دوم (برای محاسبه ارتفاع ردیف)
    4) نقطه‌ای که باید براش کلیک کنیم تا کل اون ردیف انتخاب بشه (مثلاً وسط ردیف)
"""

import ctypes
import json
import os
import sys

# --- رفع ناهماهنگی DPI/مقیاس صفحه‌نمایش (باید دقیقاً مثل run_no_claim_flow.py باشه) ---
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

try:
    import pyautogui
except ImportError:
    print("pyautogui نصب نیست: pip install pyautogui")
    sys.exit(1)

try:
    import win32gui
    import win32process
except ImportError:
    print("pywin32 نصب نیست: pip install pywin32")
    sys.exit(1)

try:
    import psutil
except ImportError:
    psutil = None

COORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coords.json")
TARGET_EXE = "Bime.exe"


def find_bime_window_rect():
    if psutil is None:
        print("psutil نصب نیست: pip install psutil")
        sys.exit(1)
    target_pids = {p.info["pid"] for p in psutil.process_iter(["pid", "name"])
                   if TARGET_EXE.lower() in (p.info.get("name") or "").lower()}
    if not target_pids:
        print(f"{TARGET_EXE} در حال اجرا نیست.")
        sys.exit(1)
    found = []

    def cb(hwnd, _):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid in target_pids and win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            if title.strip() and (rect[2] - rect[0]) > 200 and (rect[3] - rect[1]) > 200:
                found.append((hwnd, rect))
        return True

    win32gui.EnumWindows(cb, None)
    if not found:
        print("پنجره اصلی Bime.exe پیدا نشد.")
        sys.exit(1)
    found.sort(key=lambda f: (f[1][2] - f[1][0]) * (f[1][3] - f[1][1]), reverse=True)
    return found[0][1]


def load_coords():
    if os.path.exists(COORDS_FILE):
        with open(COORDS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"window_size": None, "points": {}, "grid": {}}


def save_coords(data):
    with open(COORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nذخیره شد در: {COORDS_FILE}")


def prompt_point(label, left, top):
    input(f"  → ماوس رو روی «{label}» ببر و Enter بزن...")
    x, y = pyautogui.position()
    rel = (x - left, y - top)
    print(f"    ثبت شد: ({rel[0]}, {rel[1]})")
    return rel


def main():
    rect = find_bime_window_rect()
    left, top, right, bottom = rect
    print(f"اندازه پنجره: {right - left} x {bottom - top}\n")
    print("نتیجه یه جستجوی نمونه با حداقل ۲-۳ ردیف باید الان روی صفحه باز باشه.\n")

    col_tl = prompt_point("گوشه بالا-چپ سلول «شماره الحاقیه» در ردیف اول", left, top)
    col_br = prompt_point("گوشه پایین-راست همون سلول (ردیف اول)", left, top)
    row2_tl = prompt_point("همون گوشه بالا-چپ ولی توی ردیف دوم", left, top)
    row_click = prompt_point("نقطه‌ای که باید کلیک بشه تا کل ردیف انتخاب بشه", left, top)
    grid_bottom = prompt_point("پایین‌ترین نقطه‌ی واقعی گرید (لبه‌ی زیر آخرین ردیف قابل‌مشاهده، قبل از اسکرول‌بار/فوتر)", left, top)

    row_height = row2_tl[1] - col_tl[1]
    if row_height <= 0:
        print("⚠ ارتفاع ردیف منفی یا صفر شد — احتمالاً ردیف دوم رو اشتباه ثبت کردی.")

    data = load_coords()
    data["window_size"] = [right - left, bottom - top]
    data["grid"] = {
        "col_left": col_tl[0],
        "col_top": col_tl[1],
        "col_right": col_br[0],
        "col_bottom": col_br[1],
        "row_height": row_height,
        "row_click_x": row_click[0],
        "row_click_y_offset": row_click[1] - col_tl[1],  # نسبت به بالای ردیف اول
        "grid_bottom_y": grid_bottom[1],
    }
    save_coords(data)
    print("\nکالیبراسیون گرید کامل شد.")


if __name__ == "__main__":
    main()
