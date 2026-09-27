# -*- coding: utf-8 -*-
"""
inspect_window.py
==================
قدم اول پروژه: پیدا کردن پنجره فناوران و بیرون کشیدن ساختار کنترل‌هاش.

هدف این اسکریپت:
    1) لیست همه پنجره‌های باز رو نشون بده تا عنوان دقیق پنجره فناوران رو پیدا کنیم.
    2) وقتی عنوان رو دادی، کل درخت کنترل‌های اون پنجره (دکمه‌ها، فیلدها، گرید و ...)
       رو با شناسه‌هاشون (automation_id / control_type / نام) چاپ کنه.
    3) این خروجی رو بعداً برای نوشتن اسکریپت‌های واقعی (کلیک روی دکمه، پر کردن فیلد،
       انتخاب ردیف گرید) استفاده می‌کنیم.

نحوه اجرا (روی همون ویندوزی که فناوران نصبه):
    pip install -r requirements.txt
    python inspect_window.py --list          # لیست همه پنجره‌های باز
    python inspect_window.py --title "بخشی از عنوان پنجره فناوران"

اگه فناوران یه برنامه قدیمی (Delphi/VB6/win32 خام) باشه، ممکنه با backend='uia'
خیلی از کنترل‌ها اسم/آی‌دی نداشته باشن. در اون صورت دوباره با backend='win32'
اجرا کن (فلگ --backend win32) و نتیجه رو مقایسه کن.
"""

import argparse
import os
import sys

try:
    from pywinauto import Desktop, Application
except ImportError:
    print("pywinauto نصب نیست. اول اجرا کن: pip install -r requirements.txt")
    sys.exit(1)

try:
    import win32api
    import win32con
    import win32process
    _HAS_WIN32 = True
except ImportError:
    _HAS_WIN32 = False


def get_process_exe(pid: int):
    """اسم فایل exe پروسه‌ای که پنجره بهش تعلق داره رو برمی‌گردونه (برای شناسایی پایدار)."""
    if not _HAS_WIN32 or not pid:
        return None
    try:
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid
        )
        path = win32process.GetModuleFileNameEx(handle, 0)
        return path
    except Exception:
        return None


def _walk_descendants(wrapper, max_depth: int, out):
    """
    fallback: درخت کنترل‌ها رو دستی پیمایش و چاپ می‌کنه. روی هر wrapper کار می‌کنه
    (برخلاف print_control_identifiers که فقط روی WindowSpecification هست).
    برای هر کنترل: نوع، نام، automation_id، کلاس و مستطیل (مختصات) رو می‌نویسه.
    """
    def line(ctrl, depth):
        try:
            info = ctrl.element_info
            ctype = getattr(info, "control_type", "") or ""
            name = getattr(info, "name", "") or ""
            auto_id = getattr(info, "automation_id", "") or ""
            cls = getattr(info, "class_name", "") or ""
            rect = ctrl.rectangle()
            return (f"{'  ' * depth}- [{ctype}] name={name!r} auto_id={auto_id!r} "
                    f"class={cls!r} rect=({rect.left},{rect.top},{rect.right},{rect.bottom})")
        except Exception as e:
            return f"{'  ' * depth}- (خطا در خواندن کنترل: {e})"

    def rec(ctrl, depth):
        print(line(ctrl, depth), file=out)
        if depth >= max_depth:
            return
        try:
            children = ctrl.children()
        except Exception:
            return
        for ch in children:
            rec(ch, depth + 1)

    rec(wrapper, 0)


def dump_tree(app, wrapper, max_depth: int, out_file: str = None):
    """
    درخت کنترل‌های یه پنجره رو چاپ می‌کنه (یا توی فایل UTF-8 می‌نویسه).
    نکته: app.windows()/desktop.windows() «wrapper» برمی‌گردونن که متد
    print_control_identifiers ندارن؛ باید از app.window(handle=...) یه
    WindowSpecification ساخت. اگه اون هم خطا داد، پیمایش دستی انجام می‌شه.
    """
    import contextlib
    import io
    buf = io.StringIO()
    try:
        spec = app.window(handle=wrapper.handle)
        if out_file:
            # پارامتر filename خود pywinauto با کدگذاری پیش‌فرض ویندوز (cp1252) می‌نویسه
            # و روی متن فارسی خطای 'charmap' می‌ده؛ پس خروجی رو از stdout می‌گیریم و
            # خودمون با UTF-8 ذخیره می‌کنیم.
            with contextlib.redirect_stdout(buf):
                spec.print_control_identifiers(depth=max_depth)
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(buf.getvalue())
            print(f"✅ درخت کنترل‌ها ذخیره شد: {out_file}")
            return
        spec.print_control_identifiers(depth=max_depth)
        return
    except Exception as e:
        print(f"(print_control_identifiers جواب نداد: {e} — پیمایش دستی...)")

    _walk_descendants(wrapper, max_depth, buf)
    text = buf.getvalue()
    if out_file:
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"✅ درخت کنترل‌ها ذخیره شد: {out_file}")
    else:
        print(text)


def list_windows(backend: str = "uia"):
    """لیست عنوان + نام exe همه پنجره‌های سطح بالای باز روی دسکتاپ."""
    desktop = Desktop(backend=backend)
    windows = desktop.windows()
    print(f"\n=== پنجره‌های باز (backend={backend}) — {len(windows)} عدد ===\n")
    for i, w in enumerate(windows):
        try:
            title = w.window_text()
            cls = w.friendly_class_name()
            pid = w.process_id()
            exe = get_process_exe(pid)
            exe_short = exe.split("\\")[-1] if exe else "؟"
            if title.strip():
                print(f"[{i}] عنوان: {title!r}  |  کلاس: {cls}  |  exe: {exe_short}  |  pid: {pid}")
        except Exception as e:
            print(f"[{i}] (خطا در خواندن پنجره: {e})")


def _safe_visible(w) -> bool:
    try:
        return w.is_visible()
    except Exception:
        return False


def find_all_pids_by_exe(exe_name: str):
    """همه pidهایی که اسم exe‌شون شامل exe_name هست رو برمی‌گردونه (اپ‌های Electron/Chromium
    مثل Bime معمولاً چند پروسه با اسم یکسان دارن: یکی اصلی با پنجره، بقیه renderer/GPU/helper
    بدون پنجره)."""
    pids = []
    try:
        import psutil
        for proc in psutil.process_iter(["pid", "name", "exe"]):
            name = (proc.info.get("name") or "")
            if exe_name.lower() in name.lower():
                pids.append(proc.info["pid"])
    except ImportError:
        print("psutil نصب نیست: pip install psutil")
    return pids


def connect_by_process(exe_name: str, backend: str = "uia", max_depth: int = 6, out_file: str = None,
                       all_windows: bool = False):
    """
    همه پروسه‌های exe_name رو پیدا می‌کنه (چون اپ‌های Electron/Chromium چندین پروسه با اسم
    یکسان دارن) و برای هرکدوم که پنجره داره، لیست پنجره‌ها + درخت کنترل رو چاپ می‌کنه.
    """
    pids = find_all_pids_by_exe(exe_name)
    if not pids:
        print(f"هیچ پروسه‌ای با نام شامل '{exe_name}' در حال اجرا نیست.")
        return

    print(f"\n{len(pids)} پروسه با نام '{exe_name}' پیدا شد: {pids}\n")

    found_any_window = False
    for pid in pids:
        try:
            app = Application(backend=backend).connect(process=pid, timeout=3)
            wins = app.windows()
        except Exception as e:
            print(f"[pid={pid}] وصل نشد یا خطا: {e}")
            continue

        if not wins:
            print(f"[pid={pid}] ۰ پنجره — احتمالاً renderer/GPU/helper process")
            continue

        found_any_window = True
        print(f"\n=== [pid={pid}] {len(wins)} پنجره پیدا شد ===\n")
        for i, w in enumerate(wins):
            try:
                print(f"  [{i}] عنوان: {w.window_text()!r}  |  کلاس: {w.friendly_class_name()}  |  "
                      f"visible: {w.is_visible()}  |  hwnd: {w.handle}")
            except Exception as e:
                print(f"  [{i}] (خطا: {e})")

        if all_windows:
            # همه‌ی پنجره‌های visible این پروسه (مثلاً پنجره‌ی PDF، پرینت، Save) جدا جدا
            targets = [w for w in wins if _safe_visible(w)] or wins
            for i, w in enumerate(targets):
                win_out = None
                if out_file:
                    base, ext = os.path.splitext(out_file)
                    win_out = f"{base}_{i}{ext or '.txt'}"
                print(f"\n--- [{i}] درخت کنترل‌های پنجره '{w.window_text()}' (hwnd={w.handle}) ---\n")
                try:
                    dump_tree(app, w, max_depth, win_out)
                except Exception as e:
                    print(f"خطا در چاپ درخت کنترل: {e}")
            continue

        main_win = None
        for w in wins:
            try:
                if w.is_visible() and w.window_text().strip():
                    main_win = w
                    break
            except Exception:
                continue
        main_win = main_win or wins[0]

        print(f"\n--- درخت کنترل‌های پنجره '{main_win.window_text()}' (pid={pid}) ---\n")
        try:
            dump_tree(app, main_win, max_depth, out_file)
        except Exception as e:
            print(f"خطا در چاپ درخت کنترل: {e}")

    if not found_any_window:
        print("\nهیچ‌کدوم از پروسه‌های Bime.exe پنجره‌ی قابل‌مشاهده‌ای نداشتن.")
        print("مطمئن شو Bime واقعاً باز و minimize نشده (نه توی tray).")


def dump_by_hwnd(hwnd: int, backend: str = "uia", max_depth: int = 8, out_file: str = None):
    """
    مستقیم با یه hwnd خام (که مثلاً از raw_enum_windows.py پیدا کردیم) به پنجره وصل
    می‌شه و درخت کنترل‌هاش رو چاپ می‌کنه. مطمئن‌ترین روش وقتی از قبل hwnd رو داریم.
    """
    try:
        app = Application(backend=backend).connect(handle=hwnd, timeout=5)
        win = app.window(handle=hwnd).wrapper_object()
        print(f"\n=== درخت کنترل‌های پنجره hwnd={hwnd} (backend={backend}) ===\n")
        dump_tree(app, win, max_depth, out_file)
    except Exception as e:
        print(f"خطا در وصل شدن به hwnd={hwnd}: {e}")
        if backend == "uia":
            print("امتحان کن با --backend win32 هم اجرا کنی.")


def dump_control_tree(title_substr: str = None, process_substr: str = None,
                       backend: str = "uia", max_depth: int = 6, out_file: str = None):
    """
    به پنجره‌ای که در عنوانش title_substr هست (یا نام exe‌اش process_substr رو داره) وصل
    می‌شه و درخت کنترل‌هاش رو چاپ می‌کنه. این خروجی رو نگه دار — برای نوشتن هر مرحله از
    فلوی عدم خسارت/الحاقیه لازممون می‌شه.
    """
    desktop = Desktop(backend=backend)
    target = None
    for w in desktop.windows():
        try:
            title_match = title_substr and title_substr in w.window_text()
            process_match = False
            if process_substr:
                exe = get_process_exe(w.process_id())
                process_match = exe and process_substr.lower() in exe.lower()
            if title_match or process_match:
                target = w
                break
        except Exception:
            continue

    if target is None:
        crit = title_substr or process_substr
        print(f"پنجره‌ای مطابق با '{crit}' پیدا نشد.")
        print("از --list استفاده کن تا عنوان دقیق / نام exe رو ببینی.")
        return

    print(f"\n=== درخت کنترل‌های پنجره: {target.window_text()!r} (backend={backend}) ===\n")
    try:
        app = Application(backend=backend).connect(handle=target.handle, timeout=5)
        dump_tree(app, target, max_depth, out_file)
    except Exception as e:
        print(f"خطا در چاپ درخت کنترل: {e}")
        print("اگه با uia جواب نداد، با --backend win32 دوباره امتحان کن.")


def main():
    parser = argparse.ArgumentParser(description="شناسایی پنجره و کنترل‌های فناوران")
    parser.add_argument("--list", action="store_true", help="لیست همه پنجره‌های باز")
    parser.add_argument("--title", type=str, default=None,
                         help="بخشی از عنوان پنجره فناوران برای دراپ کردن درخت کنترل‌ها")
    parser.add_argument("--process", type=str, default=None,
                         help="بخشی از نام exe فناوران (پایدارتر از عنوان چون عنوان عوض می‌شه)")
    parser.add_argument("--connect", type=str, default=None,
                         help="مستقیم به این exe وصل شو (مثلاً Bime.exe) - مطمئن‌تر از --process برای اپ‌هایی که پنجره‌شون توی لیست دسکتاپ دیده نمی‌شه")
    parser.add_argument("--hwnd", type=int, default=None,
                         help="مستقیم با یه window handle خام (از raw_enum_windows.py) وصل شو - مطمئن‌ترین روش")
    parser.add_argument("--backend", type=str, default="uia", choices=["uia", "win32"],
                         help="نوع backend برای pywinauto (پیش‌فرض uia)")
    parser.add_argument("--depth", type=int, default=6, help="عمق درخت کنترل‌ها")
    parser.add_argument("--delay", type=int, default=0,
                         help="چند ثانیه صبر قبل از گرفتن خروجی (تا مثلاً منوی بازشو رو باز کنی)")
    parser.add_argument("--all-windows", action="store_true",
                         help="با --connect: درخت همه‌ی پنجره‌های visible پروسه رو بگیر (PDF/پرینت/Save و ...)؛ "
                              "با --out هر پنجره توی فایل جدا (tree_0.txt, tree_1.txt, ...) ذخیره می‌شه")
    parser.add_argument("--out", type=str, default=None,
                         help="ذخیره خروجی درخت کنترل‌ها در فایل UTF-8 (مثلاً tree.txt) — برای فرستادن راحت‌تره")
    args = parser.parse_args()

    if args.delay:
        import time
        print(f"⏳ {args.delay} ثانیه وقت داری صفحه/منو رو آماده کنی...")
        for i in range(args.delay, 0, -1):
            print(f"  {i}...", flush=True)
            time.sleep(1)

    if args.list:
        list_windows(backend=args.backend)
    elif args.hwnd:
        dump_by_hwnd(args.hwnd, backend=args.backend, max_depth=args.depth, out_file=args.out)
    elif args.connect:
        connect_by_process(args.connect, backend=args.backend, max_depth=args.depth, out_file=args.out,
                           all_windows=args.all_windows)
    elif args.title or args.process:
        dump_control_tree(args.title, args.process, backend=args.backend, max_depth=args.depth, out_file=args.out)
    else:
        print("یکی از --list یا --title یا --process یا --connect یا --hwnd رو بده. مثال:")
        print("  python inspect_window.py --list")
        print('  python inspect_window.py --title "10065937"')
        print('  python inspect_window.py --connect "Bime.exe"')
        print('  python inspect_window.py --hwnd 1114302')


if __name__ == "__main__":
    main()
