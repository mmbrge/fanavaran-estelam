# -*- coding: utf-8 -*-
"""
raw_enum_windows.py
====================
ابزار دیباگ سطح پایین: مستقیم با win32gui همه‌ی پنجره‌های سیستم (چه top-level چه
child، چه مخفی چه نمایان) رو لیست می‌کنه و اونایی که متعلق به یه PID یا exe خاص
هستن رو نشون می‌ده. این برای وقتیه که pywinauto (سطح بالاتر) چیزی رو گزارش نمی‌ده
ولی می‌دونیم پنجره واقعاً روی صفحه هست.

اجرا:
    python raw_enum_windows.py --pid 8664
    python raw_enum_windows.py --exe Bime.exe
"""

import argparse
import sys

try:
    import win32gui
    import win32process
    import win32api
    import win32con
except ImportError:
    print("pywin32 نصب نیست یا کامل نیست. pip install pywin32")
    sys.exit(1)

try:
    import psutil
except ImportError:
    psutil = None


def get_pids_by_exe(exe_name: str):
    pids = []
    if psutil is None:
        print("psutil نصب نیست: pip install psutil")
        return pids
    for proc in psutil.process_iter(["pid", "name"]):
        name = proc.info.get("name") or ""
        if exe_name.lower() in name.lower():
            pids.append(proc.info["pid"])
    return pids


def describe_window(hwnd):
    try:
        title = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        visible = win32gui.IsWindowVisible(hwnd)
        enabled = win32gui.IsWindowEnabled(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        parent = win32gui.GetParent(hwnd)
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        is_child = bool(style & win32con.WS_CHILD)
        return {
            "hwnd": hwnd, "title": title, "class": cls, "visible": visible,
            "enabled": enabled, "pid": pid, "rect": rect, "parent": parent,
            "is_child": is_child,
        }
    except Exception as e:
        return {"hwnd": hwnd, "error": str(e)}


def enum_all_windows_for_pids(target_pids):
    """
    همه‌ی پنجره‌های سیستم (top-level) رو می‌گرده، و برای هرکدوم که pidش توی
    target_pids هست، خودش + همه‌ی فرزندانش (بازگشتی) رو چاپ می‌کنه.
    """
    matches = []

    def enum_top_level(hwnd, _):
        info = describe_window(hwnd)
        if info.get("pid") in target_pids:
            matches.append(info)
        return True

    win32gui.EnumWindows(enum_top_level, None)

    print(f"\n=== {len(matches)} پنجره top-level برای pidهای {target_pids} پیدا شد ===\n")
    for info in matches:
        if "error" in info:
            print(f"  خطا: {info['error']}")
            continue
        print(
            f"  hwnd={info['hwnd']}  عنوان={info['title']!r}  کلاس={info['class']!r}  "
            f"visible={info['visible']}  enabled={info['enabled']}  child={info['is_child']}  "
            f"rect={info['rect']}"
        )
        # فرزندان این پنجره رو هم چاپ کن (تا عمق ۲) چون ممکنه پنجره واقعی، child باشه
        children = []
        try:
            win32gui.EnumChildWindows(
                info["hwnd"], lambda h, _: children.append(describe_window(h)) or True, None
            )
        except Exception:
            pass
        for c in children[:30]:  # سقف ۳۰ تا برای خوانایی
            if "error" in c:
                continue
            title_or_empty = c["title"] or ""
            print(f"      └─ child hwnd={c['hwnd']}  عنوان={title_or_empty!r}  کلاس={c['class']!r}  visible={c['visible']}")
        if len(children) > 30:
            print(f"      └─ (+{len(children) - 30} فرزند دیگه، برای خوانایی نشون داده نشد)")

    if not matches:
        print("هیچ پنجره top-level‌ای برای این pidها پیدا نشد.")
        print("این یعنی یا pid اشتباهه، یا پنجره واقعاً به این پروسه‌ها تعلق نداره.")


def main():
    parser = argparse.ArgumentParser(description="لیست خام همه‌ی پنجره‌های یک پروسه (bypass pywinauto)")
    parser.add_argument("--pid", type=int, action="append", help="pid هدف (می‌تونی چندبار تکرار کنی)")
    parser.add_argument("--exe", type=str, default=None, help="نام exe برای پیدا کردن همه pidهای مرتبط")
    args = parser.parse_args()

    target_pids = set(args.pid or [])
    if args.exe:
        found = get_pids_by_exe(args.exe)
        print(f"pidهای پیدا شده برای '{args.exe}': {found}")
        target_pids.update(found)

    if not target_pids:
        print("یکی از --pid یا --exe رو بده.")
        return

    enum_all_windows_for_pids(target_pids)


if __name__ == "__main__":
    main()
