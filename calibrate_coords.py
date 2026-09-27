# -*- coding: utf-8 -*-
"""
calibrate_coords.py
====================
چون UI Automation نمی‌تونه دکمه/فیلدهای فناوران رو ببینه (گرید و کنترل‌ها با روش
سفارشی رندر می‌شن)، مختصات هر عنصر رو یه‌بار به‌صورت دستی ثبت می‌کنیم و بعداً
اتوماسیون بر همون مبنا کلیک/تایپ می‌کنه.

مختصات نسبت به گوشه بالا-چپ پنجره ذخیره می‌شه (نه مطلق روی صفحه)، پس اگه پنجره
جابه‌جا بشه (ولی اندازه‌ش عوض نشه) بازم درست کار می‌کنه. **پنجره باید هر بار با
همون اندازه (ترجیحاً maximize) باز باشه.**

نحوه استفاده:
    1) فناوران رو باز کن و برو به همون صفحه‌ای که می‌خوای عناصرش رو کالیبره کنی
       (مثلاً فرم صدور بیمه‌نامه بدنه).
    2) اجرا کن:
         python calibrate_coords.py --labels "search_field,filter_btn,estelam_btn,print_btn"
       (اسم لیبل‌ها دلخواهه؛ همینایی که خودت می‌خوای بعداً توی کد صداشون کنی)
    3) برای هر لیبل، ماوس رو دقیقاً روی وسط اون عنصر ببر و توی همین کنسول Enter بزن.
    4) در پایان، مختصات توی coords.json ذخیره می‌شه (به همراه اندازه پنجره در لحظه
       ثبت، برای هشدار اگه بعداً اندازه فرق کرد).

برای تست یه مختصات ثبت‌شده (بدون کلیک واقعی، فقط جابه‌جایی موس):
    python calibrate_coords.py --test search_field

برای کلیک واقعی روی یه مختصات ثبت‌شده (احتیاط کن!):
    python calibrate_coords.py --click search_field
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
    import win32gui
    import win32process
except ImportError:
    print("pywin32 نصب نیست. pip install pywin32")
    sys.exit(1)

try:
    import psutil
except ImportError:
    psutil = None

try:
    import pyautogui
except ImportError:
    pyautogui = None

COORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coords.json")
TARGET_EXE = "Bime.exe"


def find_bime_window_rect():
    """پنجره اصلی Bime.exe رو پیدا می‌کنه و مستطیلش (left, top, right, bottom) رو برمی‌گردونه."""
    if psutil is None:
        print("psutil نصب نیست: pip install psutil")
        sys.exit(1)

    target_pids = set()
    for proc in psutil.process_iter(["pid", "name"]):
        if TARGET_EXE.lower() in (proc.info.get("name") or "").lower():
            target_pids.add(proc.info["pid"])

    if not target_pids:
        print(f"پروسه‌ای با نام {TARGET_EXE} پیدا نشد. مطمئن شو فناوران بازه.")
        sys.exit(1)

    found = []

    def enum_cb(hwnd, _):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid in target_pids and win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            # پنجره اصلی معمولاً بزرگترین پنجرهٔ visible با عنوان غیرخالیه
            if title.strip() and (rect[2] - rect[0]) > 200 and (rect[3] - rect[1]) > 200:
                found.append((hwnd, title, rect))
        return True

    win32gui.EnumWindows(enum_cb, None)

    if not found:
        print("پنجره‌ی اصلی و قابل‌مشاهده‌ای برای Bime.exe پیدا نشد.")
        sys.exit(1)

    # بزرگترین پنجره رو به‌عنوان پنجره اصلی در نظر می‌گیریم
    found.sort(key=lambda f: (f[2][2] - f[2][0]) * (f[2][3] - f[2][1]), reverse=True)
    hwnd, title, rect = found[0]
    print(f"پنجره اصلی پیدا شد: {title!r}  (hwnd={hwnd})  rect={rect}")
    return rect


def load_coords():
    if os.path.exists(COORDS_FILE):
        with open(COORDS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"window_size": None, "points": {}}


def save_coords(data):
    with open(COORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\nذخیره شد در: {COORDS_FILE}")


def interactive_calibrate():
    """
    حالت تعاملی: تا خودت نگی تمومه، ادامه می‌ده. هر بار اسم نقطه رو خودت تایپ
    می‌کنی، بعد هاور می‌کنی و Enter می‌زنی. برای پایان دادن، به‌جای اسم، 'q' بزن.
    نقاطی که از قبل توی coords.json بودن حفظ می‌شن (اضافه می‌شه، پاک نمی‌شه).
    """
    if pyautogui is None:
        print("pyautogui نصب نیست. pip install pyautogui")
        sys.exit(1)

    rect = find_bime_window_rect()
    left, top, right, bottom = rect
    win_w, win_h = right - left, bottom - top

    data = load_coords()
    data["window_size"] = [win_w, win_h]

    print(f"\nاندازه پنجره: {win_w} x {win_h}")
    print("حالت تعاملی: هر بار اسم نقطه رو بده، هاور کن، Enter بزن. برای پایان 'q' بزن.\n")
    print("نکته: برای اینکه موس جابه‌جا نشه وقتی می‌خوای بین فناوران و ترمینال سوییچ کنی،")
    print("از Alt+Tab (نه کلیک) استفاده کن.\n")

    count = 0
    while True:
        label = input("اسم نقطه بعدی (یا 'q' برای پایان): ").strip()
        if label.lower() == "q":
            break
        if not label:
            continue

        input(f"  → ماوس رو روی «{label}» ببر (با Alt+Tab برگرد اینجا) و Enter بزن...")
        x, y = pyautogui.position()
        rel_x, rel_y = x - left, y - top
        data["points"][label] = {"x": rel_x, "y": rel_y}
        print(f"    ✅ ثبت شد: {label} = ({rel_x}, {rel_y})\n")
        count += 1
        save_coords(data)  # هر بار فوری ذخیره می‌شه تا چیزی از دست نره

    print(f"\n{count} نقطه جدید ثبت شد. مجموع نقاط: {len(data['points'])}")
    print("لیست همه نقاط:")
    for k, v in data["points"].items():
        print(f"  {k}: ({v['x']}, {v['y']})")


def calibrate(labels):
    if pyautogui is None:
        print("pyautogui نصب نیست. pip install pyautogui")
        sys.exit(1)

    rect = find_bime_window_rect()
    left, top, right, bottom = rect
    win_w, win_h = right - left, bottom - top

    data = load_coords()
    data["window_size"] = [win_w, win_h]

    print(f"\nاندازه پنجره: {win_w} x {win_h}")
    print("برای هر لیبل، ماوس رو روی وسط عنصر ببر و Enter بزن (Ctrl+C برای لغو)\n")

    for label in labels:
        input(f"  → ماوس رو روی «{label}» ببر و Enter بزن...")
        x, y = pyautogui.position()
        rel_x, rel_y = x - left, y - top
        data["points"][label] = {"x": rel_x, "y": rel_y}
        print(f"    ثبت شد: {label} = ({rel_x}, {rel_y}) نسبت به گوشه پنجره\n")

    save_coords(data)


def test_point(label, click=False):
    if pyautogui is None:
        print("pyautogui نصب نیست. pip install pyautogui")
        sys.exit(1)

    data = load_coords()
    point = data.get("points", {}).get(label)
    if point is None:
        print(f"لیبل '{label}' توی coords.json پیدا نشد. لیبل‌های موجود: {list(data.get('points', {}).keys())}")
        return

    rect = find_bime_window_rect()
    left, top, right, bottom = rect
    win_w, win_h = right - left, bottom - top

    saved_size = data.get("window_size")
    if saved_size and (saved_size[0] != win_w or saved_size[1] != win_h):
        print(
            f"⚠ هشدار: اندازه پنجره فرق کرده! زمان ثبت: {saved_size}، الان: [{win_w}, {win_h}]. "
            "مختصات ممکنه دیگه درست نباشه — بهتره دوباره کالیبره کنی."
        )

    abs_x, abs_y = left + point["x"], top + point["y"]
    print(f"جابه‌جایی موس به ({abs_x}, {abs_y})  [نسبی: ({point['x']}, {point['y']})]")
    pyautogui.moveTo(abs_x, abs_y, duration=0.3)

    if click:
        time.sleep(0.2)
        pyautogui.click()
        print("کلیک شد.")


def main():
    parser = argparse.ArgumentParser(description="کالیبراسیون مختصات عناصر فناوران")
    parser.add_argument("--labels", type=str, default=None,
                         help="لیست لیبل‌های موردنظر با کاما جدا شده، مثلاً 'search_field,estelam_btn'")
    parser.add_argument("--interactive", action="store_true",
                         help="حالت تعاملی: تا خودت 'q' نزنی ادامه می‌ده، اسم هر نقطه رو خودت موقع اجرا می‌دی")
    parser.add_argument("--test", type=str, default=None, help="فقط ماوس رو به این لیبل ببر (بدون کلیک)")
    parser.add_argument("--click", type=str, default=None, help="ماوس رو به این لیبل ببر و کلیک کن (احتیاط!)")
    args = parser.parse_args()

    if args.interactive:
        interactive_calibrate()
    elif args.labels:
        labels = [l.strip() for l in args.labels.split(",") if l.strip()]
        calibrate(labels)
    elif args.test:
        test_point(args.test, click=False)
    elif args.click:
        test_point(args.click, click=True)
    else:
        print("یکی از --labels یا --interactive یا --test یا --click رو بده.")


if __name__ == "__main__":
    main()
