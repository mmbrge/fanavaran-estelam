# -*- coding: utf-8 -*-
"""
flow_runner.py
==============
موتور اجرای «فلو»: یه فایل JSON که توش مرحله‌به‌مرحله می‌گیم چه کلیدی بزنه، کجا
کلیک کنه، چی تایپ کنه. اینطوری هر وقت خواستیم توالی کارها رو عوض کنیم، فقط فایل
JSON رو ویرایش می‌کنیم — نیازی به تغییر کد پایتون نیست.

انواع مرحله (step type):
    - "key":      {"type": "key", "keys": "f5"}                 -> زدن یه کلید یا ترکیب (enter, tab, ctrl+end, ...)
    - "click":    {"type": "click", "point": "search_field"}    -> کلیک روی نقطه‌ای که با calibrate_coords.py ثبت شده
    - "type":     {"type": "type", "text": "{insurer_code}"}    -> تایپ متن؛ {insurer_code} و {insurer_name} از ردیف اکسل جایگزین می‌شن
    - "sleep":    {"type": "sleep", "seconds": 1.5}              -> مکث
    - "note":     {"type": "note", "text": "..."}                -> فقط پیام توی کنسول (برای مراحلی که هنوز کامل نشده)

نحوه اجرا:
    python flow_runner.py --flow flows/no_claim_flow.json --excel "ورودی.xlsx" --insurer-col "بیمه گذار"

پیش‌فرض بعد از هر ردیف مکث می‌کنه و منتظر تأیید کاربر (Enter) می‌مونه تا بره سراغ
ردیف بعدی — یعنی تأیید نهایی همیشه دست خودته. اگه خواستی بدون مکث اجرا بشه (فقط
بعد از اطمینان کامل از درستی فلو):
    python flow_runner.py --flow flows/no_claim_flow.json --excel "ورودی.xlsx" --auto
"""

import argparse
import ctypes
import json
import os
import sys
import time

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

from excel_utils import get_yellow_rows, write_status
from grid_reader import find_last_zero_row

COORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coords.json")
TARGET_EXE = "Bime.exe"

pyautogui.PAUSE = 0.15  # مکث کوتاه پیش‌فرض بین هر دستور pyautogui (پایداری بیشتر)


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
    return found[0][1]  # (left, top, right, bottom)


def load_coords():
    if not os.path.exists(COORDS_FILE):
        print(f"فایل {COORDS_FILE} پیدا نشد. اول با calibrate_coords.py نقاط رو ثبت کن.")
        sys.exit(1)
    with open(COORDS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def load_flow(flow_path):
    with open(flow_path, "r", encoding="utf-8") as f:
        return json.load(f)


def substitute(text, context):
    """جایگزینی {insurer_code} و بقیه فیلدهای context توی متن."""
    try:
        return text.format(**context)
    except KeyError as e:
        print(f"⚠ متغیر {e} توی context تعریف نشده — متن خام استفاده می‌شه: {text!r}")
        return text


def run_flow(flow, coords, context, window_rect):
    """یه فلو رو (لیست steps) برای یه ردیف اجرا می‌کنه."""
    left, top, right, bottom = window_rect
    points = coords.get("points", {})

    for i, step in enumerate(flow.get("steps", [])):
        stype = step.get("type")

        if stype == "note":
            print(f"  📝 {step.get('text', '')}")

        elif stype == "sleep":
            time.sleep(step.get("seconds", 1))

        elif stype == "key":
            keys = step.get("keys", "")
            print(f"  ⌨️  کلید: {keys}")
            if "+" in keys:
                pyautogui.hotkey(*keys.split("+"))
            else:
                pyautogui.press(keys)

        elif stype == "click":
            label = step.get("point")
            point = points.get(label)
            if point is None:
                print(f"  ❌ نقطه‌ی '{label}' توی coords.json پیدا نشد — این مرحله رد شد.")
                continue
            abs_x, abs_y = left + point["x"], top + point["y"]
            print(f"  🖱️  کلیک روی '{label}' -> ({abs_x}, {abs_y})")
            pyautogui.click(abs_x, abs_y)

        elif stype == "type":
            text = substitute(step.get("text", ""), context)
            print(f"  ⌨️  تایپ: {text!r}")
            pyautogui.typewrite(text, interval=0.02)

        elif stype == "select_last_zero_row":
            grid_config = coords.get("grid")
            if not grid_config:
                print("  ❌ گرید کالیبره نشده (coords.json فاقد بخش 'grid' است) — از calibrate_grid.py استفاده کن.")
                continue
            print("  🔍 در حال خوندن ردیف‌های گرید با OCR...")
            result = find_last_zero_row(window_rect, grid_config)
            if result is None:
                print("  ❌ هیچ ردیفی با مقدار 0 پیدا نشد — این ردیف رد شد.")
                context["_row_select_failed"] = True
            else:
                row_idx, click_x, click_y = result
                print(f"  ✅ ردیف {row_idx} انتخاب شد -> کلیک ({click_x}, {click_y})")
                pyautogui.click(click_x, click_y)

        else:
            print(f"  ⚠ نوع مرحله ناشناخته: {stype} — رد شد.")


def main():
    parser = argparse.ArgumentParser(description="اجرای فلوی اتوماسیون فناوران")
    parser.add_argument("--flow", required=True, help="مسیر فایل JSON فلو")
    parser.add_argument("--excel", required=True, help="مسیر فایل اکسل ورودی")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--insurer-col", default="بیمه گذار")
    parser.add_argument("--header-row", type=int, default=1)
    parser.add_argument("--auto", action="store_true",
                         help="بدون مکث/تأیید بین ردیف‌ها اجرا کن (فقط بعد از اطمینان کامل!)")
    args = parser.parse_args()

    flow = load_flow(args.flow)
    coords = load_coords()

    saved_size = coords.get("window_size")
    window_rect = find_bime_window_rect()
    win_w, win_h = window_rect[2] - window_rect[0], window_rect[3] - window_rect[1]
    if saved_size and (saved_size[0] != win_w or saved_size[1] != win_h):
        print(f"⚠ هشدار: اندازه پنجره با زمان کالیبراسیون فرق داره ({saved_size} != [{win_w}, {win_h}]).")
        if input("ادامه بدم؟ (y/n): ").strip().lower() != "y":
            return

    rows = get_yellow_rows(args.excel, args.sheet, args.insurer_col, args.header_row)
    print(f"\n{len(rows)} ردیف زرد پیدا شد.\n")

    for idx, row in enumerate(rows, 1):
        print(f"\n{'=' * 50}")
        print(f"ردیف {idx}/{len(rows)}  (اکسل ردیف {row.row_index}):  نام={row.name}  کد={row.code}")
        print("=" * 50)

        if not row.code:
            print("  ❌ کدی استخراج نشد — این ردیف رد شد.")
            continue

        context = {"insurer_code": row.code, "insurer_name": row.name or ""}
        window_rect = find_bime_window_rect()  # هر بار تازه بگیر (شاید جابه‌جا شده باشه)
        run_flow(flow, coords, context, window_rect)

        if not args.auto:
            resp = input("\n➡️  Enter برای ادامه به ردیف بعد، یا 'q' برای توقف: ").strip().lower()
            if resp == "q":
                print("متوقف شد.")
                break


if __name__ == "__main__":
    main()
