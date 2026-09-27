# -*- coding: utf-8 -*-
"""
run_no_claim_flow.py
=====================
اجرای کامل فلوی «دریافت عدم خسارت» طبق ۱۷ مرحله‌ای که تعریف شده.

پیش‌نیاز:
    - coords.json باید نقاط 1 تا 17 (به‌جز 7، چون کیبوردیه) رو داشته باشه
      (با calibrate_coords.py --interactive ثبت شدن)
    - coords.json باید بخش "grid" رو هم داشته باشه (با calibrate_grid.py)
      برای مرحله 8 (پیدا کردن آخرین ردیف صفر)

اجرا:
    python run_no_claim_flow.py --excel "ورودی.xlsx"

به‌طور پیش‌فرض هر ردیفی که سلول ستون «بیمه گذار»‌ش با زرد هایلایت شده باشه
پردازش می‌شه (--select yellow). اگه روال کاری‌ت برعکسه و ردیف‌های بی‌رنگ/سفید
باید پردازش بشن، از --select white استفاده کن. بعد از هر ردیف (موفق یا ناموفق)،
ستون «وضعیت» نوشته می‌شه: موفق = سبز، ناموفق = قرمز + شماره مرحله‌ای که توش گیر کرد.

⚠️ در طول اجرا دست به موس/کیبورد نزن (چون کنترل کامل دست اسکریپته).
"""

import argparse
import ctypes
import json
import os
import sys
import time

# --- رفع ناهماهنگی DPI/مقیاس صفحه‌نمایش ---
# اگه این تنظیم انجام نشه، وقتی ویندوز روی مقیاسی غیر از ۱۰۰٪ (مثلاً ۱۲۵٪/۱۵۰٪)
# تنظیم شده باشه، win32gui.GetWindowRect مختصات رو مقیاس‌نشده برمی‌گردونه ولی
# pyautogui روی پیکسل واقعی کلیک می‌کنه -> هرچی نقطه از گوشه پنجره دورتر باشه
# خطای بیشتری می‌خوره (دقیقاً همون «بعضی کلیک‌ها درست، بقیه هرچی جلوتر بدتر»).
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
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

try:
    import pyperclip
except ImportError:
    pyperclip = None

from excel_utils import (
    get_white_rows,
    get_yellow_rows,
    write_status,
    to_western_digits,
    GREEN_FILL_ARGB,
    RED_FILL_ARGB,
)
import grid_reader
from grid_reader import find_last_zero_row_with_scroll
from folder_utils import current_persian_year_month, sanitize_folder_name

try:
    from pdf_printer import print_report_to_pdf, unique_pdf_path, PrintError
except ImportError:
    print_report_to_pdf = None

COORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coords.json")
DEBUG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug")
TARGET_EXE = "Bime.exe"
BASE_FOLDER = r"C:\Bime Ba Ma\فناوران اتومات\استعلام خسارت"

pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = True  # موس رو ببر گوشه بالا-چپ صفحه برای توقف اضطراری


class StepError(Exception):
    """خطایی که مشخص می‌کنه دقیقاً توی کدوم مرحله گیر کردیم."""
    def __init__(self, step, message):
        self.step = step
        self.message = message
        super().__init__(f"مرحله {step}: {message}")


# ---------- کمکی‌ها ----------

def load_coords():
    if not os.path.exists(COORDS_FILE):
        print(f"coords.json پیدا نشد: {COORDS_FILE}")
        sys.exit(1)
    with open(COORDS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def find_bime_window():
    """(hwnd, rect) پنجره اصلی Bime.exe رو برمی‌گردونه."""
    if psutil is None:
        print("psutil نصب نیست: pip install psutil")
        sys.exit(1)
    target_pids = {p.info["pid"] for p in psutil.process_iter(["pid", "name"])
                   if TARGET_EXE.lower() in (p.info.get("name") or "").lower()}
    if not target_pids:
        raise StepError(0, f"{TARGET_EXE} در حال اجرا نیست.")

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
        raise StepError(0, "پنجره اصلی Bime.exe پیدا نشد.")
    found.sort(key=lambda f: (f[1][2] - f[1][0]) * (f[1][3] - f[1][1]), reverse=True)
    return found[0]


def focus_window(hwnd):
    """پنجره رو میاره جلو (با ترفند Alt برای دور زدن محدودیت فوکوس ویندوز)."""
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # ترفند: یه Alt مجازی بزن تا محدودیت SetForegroundWindow دور بزنه
        pyautogui.keyDown("alt")
        win32gui.SetForegroundWindow(hwnd)
        pyautogui.keyUp("alt")
    time.sleep(0.3)


def click_point(coords, left, top, label, step_num):
    point = coords.get("points", {}).get(label)
    if point is None:
        raise StepError(step_num, f"نقطه‌ی '{label}' توی coords.json ثبت نشده.")
    x, y = left + point["x"], top + point["y"]
    pyautogui.click(x, y)
    print(f"  [مرحله {step_num}] کلیک روی '{label}' -> ({x},{y})")


def save_error_screenshot(excel_row: int, step) -> str:
    """از کل صفحه در لحظه‌ی خطا عکس می‌گیره (پوشه‌ی debug) تا معلوم باشه کجا گیر کرد."""
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        path = os.path.join(DEBUG_DIR, f"error_row{excel_row}_step{step}_{time.strftime('%H%M%S')}.png")
        pyautogui.screenshot().save(path)
        print(f"  🖼 عکس لحظه‌ی خطا: {path}")
        return path
    except Exception as e:
        print(f"  ⚠ عکس خطا ذخیره نشد: {e}")
        return ""


def paste_text(text: str, step_num: int):
    """
    متن رو از طریق کلیپ‌بورد پیست می‌کنه (Ctrl+V) به‌جای typewrite.
    دلیل: pyautogui.typewrite فقط کاراکترهای ASCII رو تایپ می‌کنه و حروف فارسی
    (مثل مسیر «استعلام خسارت» یا اسم فارسی شخص) رو بی‌صدا نادیده می‌گیره — یعنی
    اسم فایل ذخیره‌شده خراب/ناقص می‌شد. برای فیلدهای فارسی حتماً باید پیست کرد.
    """
    if pyperclip is None:
        raise StepError(step_num, "pyperclip نصب نیست ولی برای پیست متن فارسی لازمه: pip install pyperclip")
    pyperclip.copy(text)
    time.sleep(0.15)
    pyautogui.hotkey("ctrl", "v")


def find_person_folder(name: str) -> str:
    year, month = current_persian_year_month()
    safe_name = sanitize_folder_name(name) or "بدون-نام"
    folder = os.path.join(BASE_FOLDER, str(year), month, safe_name)
    os.makedirs(folder, exist_ok=True)
    return folder


# ---------- فلوی اصلی ----------

def run_one_row(coords, code: str, name: str, debug: bool = False, print_mode: str = "export"):
    """
    کل ۱۷ مرحله رو برای یه ردیف (یه کد بیمه‌گذار) اجرا می‌کنه.
    خطایی رخ بده، StepError با شماره مرحله raise می‌شه.
    """
    # پوشه شخص رو همین اول بساز (قبل از سوییچ به برنامه)
    person_folder = find_person_folder(name)
    print(f"  📁 پوشه: {person_folder}")

    if pyperclip is not None:
        pyperclip.copy(code)  # فقط به‌عنوان بک‌آپ (اگه خواستی دستی Ctrl+V بزنی) — دیگه لازم نیست نصب باشه

    hwnd, rect = find_bime_window()
    left, top, right, bottom = rect
    focus_window(hwnd)

    # مراحل ۱ تا ۵ - هر کدوم ۱ ثانیه فاصله
    for step_num in [1, 2, 3, 4, 5]:
        click_point(coords, left, top, str(step_num), step_num)
        time.sleep(1.0)

    # بعد از مرحله ۵: ۳ ثانیه صبر، F5، ۳ ثانیه صبر
    time.sleep(3.0)
    pyautogui.press("f5")
    print("  [F5] رفرش صفحه")
    time.sleep(3.0)

    # مرحله ۶: کلیک روی سلول ستون «بیمه گذار» برای انتخابش
    click_point(coords, left, top, "6", 6)
    time.sleep(0.5)

    # باز کردن پنجره‌ی جستجو با کلید Subtract روی Numpad (کنار علامت ضرب *)
    # نکته‌ی مهم: این کلید با «-» (هایفن معمولی روی صفحه‌کلید اصلی) فرق داره؛
    # pyautogui این دوتا رو با کد متفاوتی می‌فرسته (subtract=Numpad, "-"=OEM minus).
    # قبلاً به‌اشتباه "-" فرستاده می‌شد که گاهی همون کلید Numpad رو شبیه‌سازی
    # نمی‌کرد و باعث می‌شد پنجره جستجو گاهی باز نشه.
    pyautogui.press("subtract")
    print("  [کلید] Subtract (Numpad -) زده شد -> پنجره جستجو باید باز شده باشه")

    # مرحله ۷: تایپ مستقیم کد (نه پیست) توی فیلد جستجو
    # چون فناوران کنترل‌های سفارشی‌رندرشده داره، احتمال زیاد میان‌بر Ctrl+V رو
    # نمی‌شناسه. تایپ مستقیم کلید‌به‌کلید مطمئن‌تره چون از کیبورد واقعی تقلید
    # می‌کنه، نه از کلیپ‌بورد سیستم. فرض بر اینه که پنجره‌ی جستجو خودش input رو
    # فوکوس‌شده باز می‌کنه (معمول برای پنجره‌های quick-search). اگه اینطور نبود
    # و تایپ بازم جای اشتباه رفت، باید یه نقطه‌ی کلیک جدید برای همین فیلد
    # کالیبره کنیم (بگو تا اضافه کنم).
    time.sleep(2.0)

    pyautogui.hotkey("ctrl", "a")  # هر متن قبلی توی فیلد پاک بشه
    time.sleep(0.2)
    # کد همیشه عدد لاتینه (extract_code نرمال کرده)، پس typewrite امنه؛ اما اگه
    # ورودی به هر دلیلی رقم فارسی داشت، to_western_digits دوباره تضمینش می‌کنه.
    pyautogui.typewrite(to_western_digits(code), interval=0.05)
    print(f"  [مرحله 7] تایپ کد: {code}")
    time.sleep(1.0)
    pyautogui.press("enter")
    time.sleep(1.0)
    pyautogui.press("enter")
    time.sleep(1.0)

    # مرحله ۸: پیدا کردن آخرین ردیف با مقدار 0 (OCR)
    grid_config = coords.get("grid")
    if not grid_config:
        raise StepError(8, "بخش 'grid' توی coords.json نیست — اول calibrate_grid.py رو اجرا کن.")
    print("  [مرحله 8] در حال خوندن ردیف‌های گرید با OCR...")
    result = find_last_zero_row_with_scroll((left, top, right, bottom), grid_config, debug=debug)
    if result is None:
        raise StepError(8, "هیچ ردیفی با مقدار 0 پیدا نشد.")
    row_idx, click_x, click_y = result
    pyautogui.click(click_x, click_y)
    print(f"  [مرحله 8] ردیف {row_idx} انتخاب شد -> ({click_x},{click_y})")
    time.sleep(1.0)

    # مراحل ۹ و ۱۰
    click_point(coords, left, top, "9", 9)
    click_point(coords, left, top, "10", 10)

    if print_mode in ("export", "print"):
        # مراحل ۱۱ تا ذخیره بدون مختصات، با UI Automation (pdf_printer.py): منتظر
        # باز شدن واقعی نمایشگر گزارش می‌مونه (نه sleep ثابت)، Export To PDF (یا پرینت
        # با پرینتر PDF) رو می‌زنه، مسیر رو مستقیم توی پنجره‌ی Save می‌نویسه و ساخته
        # شدن فایل رو چک می‌کنه.
        if print_report_to_pdf is None:
            raise StepError(11, "pdf_printer.py یا pywinauto در دسترس نیست — pip install pywinauto")
        pdf_path = unique_pdf_path(person_folder)
        try:
            print_report_to_pdf(hwnd, pdf_path, method=print_mode)
        except PrintError as e:
            raise StepError(e.step, e.message)
        print(f"  [ذخیره] ✅ فایل ساخته شد: {pdf_path}")
        time.sleep(1.0)
        focus_window(hwnd)  # برای مراحل ۱۴ تا ۱۷ که هنوز با مختصات هستن
    else:
        # روش قدیمی مبتنی بر مختصات (--print-mode coords)
        # بعد از ۸ ثانیه، پنجره PDF استعلام خسارت باز می‌شه
        print("  ⏳ صبر ۸ ثانیه برای باز شدن پنجره PDF...")
        time.sleep(8.0)

        # مرحله ۱۱: (دکمه پرینت توی PDF viewer)
        click_point(coords, left, top, "11", 11)
        time.sleep(1.5)

        # مرحله ۱۲: انتخاب Microsoft Print to PDF
        click_point(coords, left, top, "12", 12)
        time.sleep(0.5)

        # مرحله ۱۳: دکمه Print
        click_point(coords, left, top, "13", 13)
        print("  ⏳ صبر ۲ ثانیه برای باز شدن پنجره Save...")
        time.sleep(2.0)

        # پنجره «Save Print Output As» - مستقیم توی فیلد File name مسیر کامل رو تایپ می‌کنیم
        # (پنجره Save همیشه فیلد File name رو فوکوس‌شده باز می‌کنه، نیازی به کلیک نیست)
        full_path_no_ext = os.path.join(person_folder, "استعلام خسارت")
        pyautogui.hotkey("ctrl", "a")  # هر چیزی از قبل توی فیلد بود پاک بشه
        time.sleep(0.2)
        # مسیر شامل حروف فارسی («استعلام خسارت» + اسم فارسی شخص + نام ماه) هست، پس
        # باید پیست بشه نه typewrite (typewrite حروف فارسی رو نادیده می‌گیره).
        paste_text(full_path_no_ext, 13)
        print(f"  [ذخیره] مسیر: {full_path_no_ext}.pdf")
        time.sleep(0.3)
        pyautogui.press("enter")  # فشردن Save
        print("  ⏳ صبر ۳ ثانیه برای ذخیره شدن...")
        time.sleep(3.0)

    # مراحل ۱۴ تا ۱۷ - هر کدوم ۱ ثانیه فاصله
    for step_num in [14, 15, 16, 17]:
        click_point(coords, left, top, str(step_num), step_num)
        time.sleep(1.0)

    return person_folder


def main():
    parser = argparse.ArgumentParser(description="اجرای کامل فلوی دریافت عدم خسارت")
    parser.add_argument("--excel", required=True, help="مسیر فایل اکسل ورودی")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--insurer-col", default="بیمه گذار")
    parser.add_argument("--header-row", type=int, default=1)
    parser.add_argument("--select", choices=["yellow", "white"], default="yellow",
                         help="کدوم ردیف‌ها پردازش بشن: 'yellow' = ردیف‌های هایلایت‌زرد "
                              "(پیش‌فرض، مطابق روال کاری فعلی)، 'white' = ردیف‌های بی‌رنگ/سفید.")
    parser.add_argument("--auto", action="store_true",
                         help="بدون مکث/تأیید بین ردیف‌ها اجرا کن (فقط بعد از اطمینان کامل!)")
    parser.add_argument("--print-mode", choices=["export", "print", "coords"], default="export",
                         help="ذخیره‌ی PDF: 'export' = دکمه‌ی Export To PDF بدون مختصات (پیش‌فرض)، "
                              "'print' = پرینت با Microsoft Print to PDF بدون مختصات، "
                              "'coords' = روش قدیمی با نقاط ۱۱ تا ۱۳")
    parser.add_argument("--ocr", choices=["tesseract", "paddle"], default="tesseract",
                         help="موتور OCR برای خوندن گرید (مرحله ۸). paddle نیاز به "
                              "pip install paddlepaddle paddleocr داره؛ اگه جواب نده، خودکار tesseract.")
    parser.add_argument("--debug", action="store_true",
                         help="چاپ جزئیات OCR گرید (برای عیب‌یابی مرحله ۸)")
    args = parser.parse_args()

    # اگه فایل اکسل همزمان توی Excel باز باشه، ویندوز قفلش می‌کنه و write_status
    # بعد از اولین ردیف با PermissionError از کار می‌افته؛ پس همین اول چک می‌کنیم.
    if not os.path.exists(args.excel):
        print(f"❌ فایل اکسل پیدا نشد: {args.excel}")
        return
    try:
        with open(args.excel, "a+b"):
            pass
    except PermissionError:
        print("❌ فایل اکسل الان توی Excel بازه (یا قفله). اول ببندش، بعد دوباره اجرا کن.")
        return

    coords = load_coords()
    grid_reader.OCR_ENGINE = args.ocr
    if args.debug:
        grid_reader.DEBUG_DIR = DEBUG_DIR

    # --- چک اندازه پنجره قبل از شروع ---
    # اگه اندازه‌ی پنجره الان با اندازه‌ی زمان کالیبراسیون فرق داشته باشه (یا DPI عوض
    # شده باشه)، همه‌ی مختصات ثبت‌شده اشتباه می‌شن. اینجا زودتر هشدار می‌دیم تا وسط
    # کار با کلیک‌های غلط سروکله نزنیم.
    saved_size = coords.get("window_size")
    try:
        _, rect0 = find_bime_window()
        win_w0, win_h0 = rect0[2] - rect0[0], rect0[3] - rect0[1]
        if saved_size and (saved_size[0] != win_w0 or saved_size[1] != win_h0):
            print(
                f"⚠ هشدار: اندازه پنجره با زمان کالیبراسیون فرق داره! "
                f"زمان کالیبراسیون: {saved_size}، الان: [{win_w0}, {win_h0}]."
            )
            print("این یعنی همه‌ی کلیک‌ها احتمالاً جای اشتباه می‌خورن. بهتره:")
            print("  1) پنجره Bime رو دقیقاً هم‌اندازه‌ی زمان کالیبراسیون کن (یا maximize کن)")
            print("  2) یا دوباره با calibrate_coords.py و calibrate_grid.py کالیبره کن")
            if input("با همین حال ادامه بدم؟ (y/n): ").strip().lower() != "y":
                return
    except StepError as e:
        print(f"❌ {e.message}")
        return

    if args.select == "yellow":
        rows = get_yellow_rows(args.excel, args.sheet, args.insurer_col, args.header_row)
        print(f"\n{len(rows)} ردیف زرد (برای پردازش) پیدا شد.\n")
    else:
        rows = get_white_rows(args.excel, args.sheet, args.insurer_col, args.header_row)
        print(f"\n{len(rows)} ردیف سفید (پردازش‌نشده) پیدا شد.\n")

    for idx, row in enumerate(rows, 1):
        print(f"\n{'=' * 55}")
        print(f"ردیف {idx}/{len(rows)}  (اکسل ردیف {row.row_index}):  نام={row.name}  کد={row.code}")
        print("=" * 55)

        if not row.code:
            write_status(args.excel, row.row_index, "خطا: کد استخراج نشد",
                         fill_argb=RED_FILL_ARGB, insurer_col_name=args.insurer_col)
            print("  ❌ کدی استخراج نشد.")
            continue

        try:
            folder = run_one_row(coords, row.code, row.name or "", debug=args.debug,
                                 print_mode=args.print_mode)
            write_status(args.excel, row.row_index, "انجام شد",
                         fill_argb=GREEN_FILL_ARGB, insurer_col_name=args.insurer_col)
            print(f"  ✅ موفق. ذخیره در: {folder}")
        except StepError as e:
            save_error_screenshot(row.row_index, e.step)
            write_status(args.excel, row.row_index, f"خطا در مرحله {e.step}: {e.message}",
                         fill_argb=RED_FILL_ARGB, insurer_col_name=args.insurer_col)
            print(f"  ❌ خطا در مرحله {e.step}: {e.message}")
        except Exception as e:
            save_error_screenshot(row.row_index, "x")
            write_status(args.excel, row.row_index, f"خطای غیرمنتظره: {e}",
                         fill_argb=RED_FILL_ARGB, insurer_col_name=args.insurer_col)
            print(f"  ❌ خطای غیرمنتظره: {e}")

        if not args.auto:
            resp = input("\n➡️  Enter برای ادامه به ردیف بعد، یا 'q' برای توقف: ").strip().lower()
            if resp == "q":
                print("متوقف شد.")
                break


if __name__ == "__main__":
    main()
