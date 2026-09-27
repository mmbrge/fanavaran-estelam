# -*- coding: utf-8 -*-
"""
capture_context_menu.py
========================
چون منوی راست‌کلیک با Alt+Tab بسته می‌شه، نمی‌شه با calibrate_coords.py عادی
مختصات گزینه‌ی «جستجو» توی اون منو رو ثبت کرد. این اسکریپت به‌جاش:
    1) خودش روی نقطه‌ی '6' (سلول ستون بیمه‌گذار) کلیک راست می‌کنه
    2) کمی صبر می‌کنه تا منو کامل باز بشه
    3) از کل پنجره‌ی Bime اسکرین‌شات می‌گیره و ذخیره می‌کنه

چون اسکرین‌شات فقط از محدوده‌ی پنجره گرفته می‌شه (نه کل صفحه)، هر مختصاتی که
توی خود عکس (مثلاً با Paint) اندازه بگیری، دقیقاً همون x,y نسبیه که باید توی
coords.json زیر "context_search_item" بذاری — نیازی به تبدیل نیست.

اجرا:
    python capture_context_menu.py

بعدش:
    1) فایل menu_screenshot.png رو با Paint باز کن
    2) ماوس رو روی وسط گزینه‌ی «جستجو» ببر — گوشه پایین-چپ Paint مختصات x,y رو نشون می‌ده
    3) اون عدد رو بده تا خودم توی coords.json اضافه‌ش کنم، یا دستی این‌طوری اضافه کن:
       "context_search_item": {"x": <عدد x>, "y": <عدد y>}
"""

import json
import os
import sys
import time
import ctypes

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
    import win32con
except ImportError:
    print("pywin32 نصب نیست: pip install pywin32")
    sys.exit(1)

try:
    import psutil
except ImportError:
    psutil = None

COORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coords.json")
TARGET_EXE = "Bime.exe"
OUTPUT_IMAGE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "menu_screenshot.png")


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
    return found[0]


def main():
    if not os.path.exists(COORDS_FILE):
        print(f"coords.json پیدا نشد: {COORDS_FILE}")
        sys.exit(1)
    with open(COORDS_FILE, "r", encoding="utf-8") as f:
        coords = json.load(f)

    point6 = coords.get("points", {}).get("6")
    if point6 is None:
        print("نقطه '6' توی coords.json ثبت نشده — اول اون رو کالیبره کن.")
        sys.exit(1)

    hwnd, rect = find_bime_window_rect()
    left, top, right, bottom = rect

    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pyautogui.keyDown("alt")
        win32gui.SetForegroundWindow(hwnd)
        pyautogui.keyUp("alt")
    time.sleep(0.5)

    x, y = left + point6["x"], top + point6["y"]
    print(f"کلیک چپ روی نقطه '6' -> ({x},{y}) (انتخاب سلول)")
    pyautogui.click(x, y)
    time.sleep(0.5)

    print(f"کلیک راست روی همون نقطه -> ({x},{y})")
    pyautogui.click(x, y, button="right")
    time.sleep(1.0)  # صبر تا منو کامل رندر بشه

    screenshot = pyautogui.screenshot(region=(left, top, right - left, bottom - top))
    screenshot.save(OUTPUT_IMAGE)
    print(f"\n✅ اسکرین‌شات ذخیره شد: {OUTPUT_IMAGE}")
    print("این عکس رو با Paint باز کن، روی وسط گزینه‌ی «جستجو» موس رو ببر و مختصات")
    print("پایین-چپ Paint رو بخون — همون عدد مستقیم قابل استفاده‌ست (نیازی به تبدیل نیست).")

    # صفحه رو با Esc ببند تا منو باز نمونه
    pyautogui.press("esc")


if __name__ == "__main__":
    main()
