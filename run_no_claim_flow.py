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
import screen_locator
from grid_reader import find_last_zero_row_with_scroll
from folder_utils import current_persian_year_month, sanitize_folder_name

try:
    from pdf_printer import (print_report_to_pdf, unique_pdf_path, PrintError, close_leftover_dialogs,
                             report_viewer_open)
except ImportError:
    print_report_to_pdf = None
    close_leftover_dialogs = None
    report_viewer_open = None

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


STEP8_ATTEMPTS = 4       # چند بار گرید خونده بشه تا نتیجه‌ی جستجو لود بشه
STEP8_RETRY_WAIT = 2.5   # فاصله‌ی بین تلاش‌ها (ثانیه)
TEMPLATE_TIMEOUT = 8.0  # حداکثر صبر برای ظاهر شدن تصویر یه مرحله (مثلاً آیتم منو)


def click_point(coords, left, top, label, step_num):
    """
    روی نقطه‌ی مرحله کلیک می‌کنه. اگه برای این مرحله تصویر ساخته شده باشه
    (templates/<label>.png با capture_template.py)، دکمه رو از روی تصویر روی صفحه
    پیدا می‌کنه (بدون مختصات، و تا ظاهر شدنش صبر می‌کنه)؛ وگرنه مختصات coords.json.
    اگه تصویر هست ولی پیدا نشد، خطا می‌ده (کلیک کورکورانه روی مختصات خطرناکه).
    """
    if screen_locator.has_template(label):
        try:
            x, y, score = screen_locator.locate(label, timeout=TEMPLATE_TIMEOUT)
        except LookupError as e:
            raise StepError(step_num, str(e))
        pyautogui.click(x, y)
        print(f"  [مرحله {step_num}] کلیک روی تصویر '{label}' -> ({x},{y})  شباهت={score:.2f}")
        return

    point = coords.get("points", {}).get(label)
    if point is None:
        raise StepError(step_num, f"نقطه‌ی '{label}' توی coords.json ثبت نشده.")
    x, y = left + point["x"], top + point["y"]
    pyautogui.click(x, y)
    print(f"  [مرحله {step_num}] کلیک روی '{label}' -> ({x},{y})")


def save_error_screenshot(excel_row: int, step, attempt: int = 1) -> str:
    """از کل صفحه در لحظه‌ی خطا عکس می‌گیره (پوشه‌ی debug) تا معلوم باشه کجا گیر کرد."""
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        path = os.path.join(DEBUG_DIR, f"error_row{excel_row}_step{step}_try{attempt}_{time.strftime('%H%M%S')}.png")
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


SEARCH_TEMPLATES = ("search_dialog", "search_apply")            # لازم
SEARCH_OPTIONAL_TEMPLATES = ("search_clear", "search_field")     # اختیاری


def _grid_strip(left, top, grid_config):
    """عکس خاکستری ستون «شماره الحاقیه» گرید (برای فهمیدن اینکه گرید عوض شده یا نه)."""
    x1, x2 = left + grid_config["col_left"], left + grid_config["col_right"]
    y1 = top + grid_config["col_top"]
    y2 = top + grid_config.get("grid_bottom_y", grid_config["col_bottom"] + 20 * grid_config["row_height"])
    return pyautogui.screenshot(region=(x1, y1, x2 - x1, y2 - y1)).convert("L")


def _images_differ(a, b) -> bool:
    from PIL import ImageChops
    return ImageChops.difference(a, b).point(lambda v: 255 if v > 40 else 0).getbbox() is not None


def wait_grid_updated(left, top, grid_config, before, timeout=12.0) -> str:
    """
    بعد از «اعمال»، منتظر می‌مونه گرید نسبت به قبل (before) عوض بشه و بعد ثابت بمونه
    (یعنی نتیجه‌ی جستجو کامل لود شده). خروجی: "stable" / "changed" / "unchanged".
    "unchanged" معمولاً یعنی نتیجه همون قبلیه (مثلاً همون بیمه‌گذار دوباره).
    """
    end = time.time() + timeout
    prev, changed = None, False
    while time.time() < end:
        cur = _grid_strip(left, top, grid_config)
        if not changed:
            if _images_differ(before, cur):
                changed, prev = True, cur
        elif not _images_differ(prev, cur):
            return "stable"
        else:
            prev = cur
        time.sleep(0.6)
    return "changed" if changed else "unchanged"


def _click_template(label: str, step_num: int, timeout: float = 5.0):
    try:
        x, y, _ = screen_locator.locate(label, timeout=timeout)
    except LookupError as e:
        raise StepError(step_num, str(e))
    pyautogui.click(x, y)


def _search_once(code_typed: str, left: int, top: int, grid_config, suggest_wait: float) -> str:
    """یه‌بار: پنجره‌ی جستجو (باید باز شده باشه) → تایپ کد → انتخاب → «اعمال» → وضعیت گرید."""
    if not screen_locator.wait_visible("search_dialog", 10):
        raise StepError(7, "پنجره‌ی «جست و جو» بعد از زدن Numpad − باز نشد.")
    time.sleep(0.3)

    # فیلتر جستجوی قبلی (مثلاً بیمه‌گذار ردیف قبل) نباید بمونه، وگرنه نتیجه قاطی می‌شه
    if screen_locator.has_template("search_clear"):
        _click_template("search_clear", 7)
        time.sleep(0.8)
    if screen_locator.has_template("search_field"):
        _click_template("search_field", 7)
        time.sleep(0.3)

    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.2)
    pyautogui.typewrite(code_typed, interval=0.05)
    print(f"  [مرحله 7] تایپ کد: {code_typed}  (صبر {suggest_wait:g} ثانیه برای لیست پیشنهاد)")
    time.sleep(suggest_wait)          # فرصت رسیدن لیست پیشنهاد از سرور
    pyautogui.press("enter")          # انتخاب بیمه‌گذار از لیست
    time.sleep(1.0)

    before = _grid_strip(left, top, grid_config) if grid_config else None
    _click_template("search_apply", 7)
    if not screen_locator.wait_gone("search_dialog", 10):
        raise StepError(7, "بعد از زدن «اعمال»، پنجره‌ی جستجو بسته نشد.")
    print("  [مرحله 7] ✅ «اعمال» زده شد و پنجره‌ی جستجو بسته شد")
    if before is None:
        return "stable"
    return wait_grid_updated(left, top, grid_config, before)


def search_insurer(code: str, left: int, top: int, grid_config, reopen_search):
    """
    مرحله ۷: پنجره‌ی «جست و جو» (با Numpad − باز شده) → تایپ کد → انتخاب بیمه‌گذار از
    لیست پیشنهاد → «اعمال» → صبر تا گرید نتیجه‌ی واقعی رو نشون بده.

    به‌جای sleep کور، با تصاویر تأیید می‌شه:
        search_dialog = عنوان پنجره‌ی «جست و جو»  (باز شدن/بسته شدن پنجره)   ← لازم
        search_apply  = دکمه‌ی «اعمال»                                      ← لازم
        search_clear  = لینک «حذف همه فیلترها» (پاک کردن فیلتر ردیف قبل)      ← اختیاری
        search_field  = عنوان «بیمه گذار» + فیلدش (کلیک توی فیلد)             ← اختیاری

    محافظ اصلی: بعد از «اعمال» ستون الحاقیه‌ی گرید باید عوض بشه (گرید قبل از جستجو
    لیست پیش‌فرضه). اگه عوض نشد یعنی فیلتر بیمه‌گذار اعمال نشده (مثلاً لیست پیشنهاد
    دیر رسید و Enter هدر رفت)؛ یه‌بار با صبر بیشتر تکرار می‌شه و بعد خطا — تا هیچ‌وقت
    گرید اشتباه خونده نشه و PDF شخص دیگه‌ای ذخیره نشه.
    اگه تصاویر لازم ساخته نشده باشن، روش قدیمی (دو Enter با مکث ثابت) اجرا می‌شه.
    """
    code_typed = to_western_digits(code)
    missing = [t for t in SEARCH_TEMPLATES if not screen_locator.has_template(t)]
    if missing:
        print(f"  ⚠ تصاویر {', '.join(missing)} ساخته نشده؛ جستجو با روش قدیمی (بدون تأیید) انجام می‌شه.")
        time.sleep(2.0)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        pyautogui.typewrite(code_typed, interval=0.05)
        print(f"  [مرحله 7] تایپ کد: {code}")
        time.sleep(1.0)
        pyautogui.press("enter")
        time.sleep(1.0)
        pyautogui.press("enter")
        time.sleep(1.0)
        return

    for attempt, suggest_wait in ((1, 2.5), (2, 5.0)):
        status = _search_once(code_typed, left, top, grid_config, suggest_wait)
        if status == "stable":
            print("  [مرحله 7] ✅ نتیجه‌ی جستجو توی گرید لود شد")
            return
        if status == "changed":
            print("  ⚠ گرید عوض شد ولی هنوز ثابت نشده؛ مرحله‌ی ۸ در صورت نیاز دوباره می‌خونه.")
            return
        if attempt == 1:
            print("  ⚠ گرید بعد از «اعمال» عوض نشد (فیلتر اعمال نشد). جستجو با صبر بیشتر تکرار می‌شه...")
            reopen_search()
    raise StepError(7, f"بعد از «اعمال» گرید عوض نشد؛ یعنی فیلتر بیمه‌گذار با کد {code} اعمال نشده "
                       f"(لیست پیشنهاد نیومد یا کد توی فناوران نیست).")


def find_person_folder(name: str) -> str:
    year, month = current_persian_year_month()
    safe_name = sanitize_folder_name(name) or "بدون-نام"
    folder = os.path.join(BASE_FOLDER, str(year), month, safe_name)
    os.makedirs(folder, exist_ok=True)
    return folder


# ---------- فلوی اصلی ----------

def run_one_row(coords, code: str, name: str, debug: bool = False, print_mode: str = "export",
                progress: dict = None, navigate: bool = True):
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

    if navigate:
        # مراحل ۱ تا ۵ (رسیدن به صفحه‌ی «صدور بیمه نامه بدنه») + F5 — فقط برای ردیف
        # اول، یا بعد از خطا که معلوم نیست فناوران روی چه صفحه‌ایه. ردیف‌های بعدی چون
        # مراحل ۱۴ تا ۱۷ فرم گزارش رو بستن و فیلتر رو پاک کردن، روی همین صفحه‌ان و
        # مستقیم از مرحله‌ی ۶ شروع می‌کنن.
        for step_num in [1, 2, 3, 4, 5]:
            click_point(coords, left, top, str(step_num), step_num)
            time.sleep(1.0)

        # بعد از مرحله ۵: ۳ ثانیه صبر، F5، ۳ ثانیه صبر
        time.sleep(3.0)
        pyautogui.press("f5")
        print("  [F5] رفرش صفحه")
        time.sleep(3.0)
    else:
        print("  ⏭ مراحل ۱ تا ۵ و F5 رد شدن (فناوران از ردیف قبل روی صفحه‌ی صدوره)")
        time.sleep(1.0)

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

    # مرحله ۷: جستجوی بیمه‌گذار (با تأیید هر قدم اگه تصاویر search_* ساخته شده باشن)
    def reopen_search():
        click_point(coords, left, top, "6", 6)
        time.sleep(0.5)
        pyautogui.press("subtract")

    search_insurer(code, left, top, coords.get("grid"), reopen_search)

    # مرحله ۸: پیدا کردن آخرین ردیف با مقدار 0 (OCR)
    grid_config = coords.get("grid")
    if not grid_config:
        raise StepError(8, "بخش 'grid' توی coords.json نیست — اول calibrate_grid.py رو اجرا کن.")
    # نتیجه‌ی جستجو گاهی دیرتر از sleepهای ثابت لود می‌شه و گرید قبلی (لیست
    # ردیف‌های دیگه) هنوز روی صفحه‌ست؛ پس اگه ردیف هدف پیدا نشد، چند بار با فاصله
    # دوباره می‌خونیم تا نتیجه‌ی واقعی لود بشه.
    result = None
    for attempt in range(1, STEP8_ATTEMPTS + 1):
        print(f"  [مرحله 8] در حال خوندن ردیف‌های گرید با OCR... (تلاش {attempt}/{STEP8_ATTEMPTS})")
        result = find_last_zero_row_with_scroll((left, top, right, bottom), grid_config, debug=debug)
        if result is not None:
            break
        if attempt < STEP8_ATTEMPTS:
            print(f"  ⏳ ردیف هدف پیدا نشد؛ شاید نتیجه‌ی جستجو هنوز لود نشده. "
                  f"{STEP8_RETRY_WAIT:g} ثانیه صبر و دوباره...")
            time.sleep(STEP8_RETRY_WAIT)
    if result is None:
        raise StepError(8, f"بعد از {STEP8_ATTEMPTS} بار خوندن، ردیفی با الحاقیه = 0 و وضعیت "
                           f"«ارسال به مالی» پیدا نشد.")
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
        if progress is not None:
            progress["pdf_path"] = pdf_path
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
        if progress is not None and os.path.exists(full_path_no_ext + ".pdf"):
            progress["pdf_path"] = full_path_no_ext + ".pdf"

    # مراحل ۱۴ تا ۱۷: بستن فرم گزارش و پاک کردن فیلتر بیمه‌گذار (برای ردیف بعدی)
    #   ۱۴ = × کنار «چاپ استعلام خسارت بدنه»   ۱۵ = آیکون فیلتر بالای گرید
    #   ۱۶ = «حذف همه فیلترها»                   ۱۷ = «اعمال»
    for step_num in [14, 15, 16, 17]:
        click_point(coords, left, top, str(step_num), step_num)
        time.sleep(1.0)
        if step_num == 14 and report_viewer_open is not None:
            end = time.time() + 10
            while report_viewer_open(hwnd) and time.time() < end:
                time.sleep(0.5)
            if report_viewer_open(hwnd):
                raise StepError(14, "فرم «چاپ استعلام خسارت» بعد از زدن × بسته نشد.")

    return person_folder


def recover_ui():
    """
    بعد از خطا در یه ردیف: فناوران رو به حالت تمیز برمی‌گردونه تا ردیف از مرحله‌ی ۱
    دوباره شروع بشه — منوهای باز و پنجره‌های مودال (جستجو، Print، Save، پیغام خطا)
    با Esc و بستن مستقیم پنجره‌های اضافه‌ی Bime.exe. به پنجره‌ی اصلی دست نمی‌زنه.
    """
    print("  🔄 برگردوندن فناوران به حالت اولیه...")
    try:
        hwnd, _ = find_bime_window()
    except StepError as e:
        print(f"  ⚠ {e.message}")
        return
    try:
        focus_window(hwnd)
    except Exception:
        pass
    for _ in range(3):
        pyautogui.press("esc")
        time.sleep(0.4)
    # اگه فرم «چاپ استعلام خسارت» باز مونده، با همون × مرحله‌ی ۱۴ ببندش
    if report_viewer_open is not None and report_viewer_open(hwnd):
        if screen_locator.has_template("14"):
            try:
                x, y, _ = screen_locator.locate("14", timeout=3)
                pyautogui.click(x, y)
                print("  🔄 فرم «چاپ استعلام خسارت» بسته شد")
                time.sleep(1.5)
            except LookupError:
                print("  ⚠ فرم «چاپ استعلام خسارت» بازه ولی × اون پیدا نشد.")
        else:
            print("  ⚠ فرم «چاپ استعلام خسارت» بازه؛ برای بستن خودکارش تصویر مرحله‌ی ۱۴ رو بساز.")
    if close_leftover_dialogs is not None:
        try:
            n = close_leftover_dialogs(hwnd)
            if n:
                print(f"  🔄 {n} پنجره‌ی باز مونده بسته شد")
        except Exception as e:
            print(f"  ⚠ بستن پنجره‌های اضافه انجام نشد: {e}")
    time.sleep(1.5)


def safe_write_status(excel, row_index, text, fill_argb, insurer_col, sheet=None, header_row=1):
    """
    مثل write_status، ولی اگه اکسل وسط کار باز شده باشه (قفل)، چند بار صبر و تکرار
    می‌کنه و در نهایت فقط هشدار می‌ده — تا کل اجرای خودکار به‌خاطر یه ذخیره نخوابه.
    """
    for attempt in range(1, 6):
        try:
            write_status(excel, row_index, text, sheet_name=sheet, header_row=header_row,
                         fill_argb=fill_argb, insurer_col_name=insurer_col)
            return True
        except PermissionError:
            print(f"  ⚠ فایل اکسل قفله (احتمالاً توی Excel بازه) — ببندش؛ تلاش {attempt}/5 ...")
            time.sleep(5)
    print(f"  ❌ وضعیت ردیف {row_index} توی اکسل نوشته نشد: {text}")
    return False


def main():
    parser = argparse.ArgumentParser(description="اجرای کامل فلوی دریافت عدم خسارت")
    parser.add_argument("--excel", required=True, help="مسیر فایل اکسل ورودی")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--insurer-col", default="بیمه گذار")
    parser.add_argument("--header-row", type=int, default=1)
    parser.add_argument("--select", choices=["yellow", "white"], default="yellow",
                         help="کدوم ردیف‌ها پردازش بشن: 'yellow' = ردیف‌های هایلایت‌زرد "
                              "(پیش‌فرض، مطابق روال کاری فعلی)، 'white' = ردیف‌های بی‌رنگ/سفید.")
    parser.add_argument("--confirm", action="store_true",
                         help="بعد از هر ردیف منتظر Enter بمون (پیش‌فرض: همه‌ی ردیف‌ها خودکار و بدون تأیید)")
    parser.add_argument("--auto", action="store_true", help=argparse.SUPPRESS)  # سازگاری با قبل؛ الان پیش‌فرضه
    parser.add_argument("--retries", type=int, default=2,
                         help="اگه یه ردیف خطا داد، چند بار دیگه از اول تکرار بشه (پیش‌فرض ۲)")
    parser.add_argument("--max-consecutive-failures", type=int, default=3,
                         help="اگه این‌همه ردیف پشت‌سرهم (بعد از همه‌ی تکرارها) شکست خوردن، اجرا متوقف بشه")
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

    # اگه فقط بعضی فایل‌ها به‌روز شده باشن، وسط کار با AttributeError از کار می‌افته؛
    # پس همین اول چک می‌کنیم فایل‌های کمکی با این نسخه هماهنگ باشن.
    stale = []
    if not all(hasattr(screen_locator, f) for f in ("locate", "is_visible", "wait_visible", "wait_gone")):
        stale.append("screen_locator.py")
    if not hasattr(grid_reader, "is_target_row"):
        stale.append("grid_reader.py")
    if stale:
        print(f"❌ این فایل‌ها قدیمی‌ان و با run_no_claim_flow.py جور نیستن: {', '.join(stale)}")
        print("   همه‌ی فایل‌های .py رو از گیت‌هاب به‌روز کن (coords.json و پوشه‌ی templates رو دست نزن).")
        return

    coords = load_coords()
    with_templates = [str(n) for n in range(1, 18) if screen_locator.has_template(str(n))]
    with_templates += [t for t in SEARCH_TEMPLATES + SEARCH_OPTIONAL_TEMPLATES + ("status_ok",)
                       if screen_locator.has_template(t)]
    if with_templates:
        print(f"🖼 مراحلی که با تصویر پیدا می‌شن (بدون مختصات): {', '.join(with_templates)}")
        if not screen_locator.available():
            print("❌ برای این مراحل opencv-python لازمه: pip install opencv-python numpy")
            return
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

    max_attempts = 1 + max(0, args.retries)
    failsafe_exc = getattr(pyautogui, "FailSafeException", ())
    done_rows, failed_rows = [], []
    consecutive_failures = 0
    # آیا فناوران الان روی صفحه‌ی صدور (بعد از مراحل ۱۴ تا ۱۷ ردیف قبل) هست؟ فقط اون
    # موقع ردیف بعدی می‌تونه مراحل ۱ تا ۵ رو رد کنه. بعد از هر خطا False می‌شه.
    on_issue_page = False
    if not args.confirm:
        print(f"▶ اجرای خودکار همه‌ی ردیف‌ها (بدون تأیید). هر ردیف در صورت خطا تا {max_attempts} بار "
              f"از اول انجام می‌شه. توقف اضطراری: موس رو ببر گوشه‌ی بالا-چپ صفحه.")

    try:
        for idx, row in enumerate(rows, 1):
            print(f"\n{'=' * 55}")
            print(f"ردیف {idx}/{len(rows)}  (اکسل ردیف {row.row_index}):  نام={row.name}  کد={row.code}")
            print("=" * 55)

            if not row.code:
                safe_write_status(args.excel, row.row_index, "خطا: کد استخراج نشد", RED_FILL_ARGB,
                                  args.insurer_col, args.sheet, args.header_row)
                print("  ❌ کدی استخراج نشد.")
                failed_rows.append((row, "کد استخراج نشد"))
                continue

            last_error, folder, progress, attempt = None, None, {}, 0
            for attempt in range(1, max_attempts + 1):
                progress = {}
                if attempt > 1:
                    print(f"  🔁 تکرار ردیف از مرحله‌ی ۱ (تلاش {attempt}/{max_attempts})...")
                try:
                    folder = run_one_row(coords, row.code, row.name or "", debug=args.debug,
                                         print_mode=args.print_mode, progress=progress,
                                         navigate=not on_issue_page)
                    last_error = None
                    on_issue_page = True
                    break
                except failsafe_exc:
                    raise
                except StepError as e:
                    last_error = e
                except Exception as e:
                    last_error = StepError("x", f"خطای غیرمنتظره: {e}")

                on_issue_page = False  # بعد از خطا صفحه معلوم نیست؛ تلاش بعدی از مرحله‌ی ۱
                save_error_screenshot(row.row_index, last_error.step, attempt)
                print(f"  ❌ تلاش {attempt}/{max_attempts} — خطا در مرحله {last_error.step}: {last_error.message}")
                if last_error.step == 0:
                    break  # Bime.exe بسته‌ست؛ تکرار فایده نداره
                if progress.get("pdf_path"):
                    break  # PDF ذخیره شده؛ تکرار فقط یه PDF تکراری می‌سازه
                recover_ui()

            if last_error is None:
                note = "انجام شد" + (f" (تلاش {attempt})" if attempt > 1 else "")
                safe_write_status(args.excel, row.row_index, note, GREEN_FILL_ARGB,
                                  args.insurer_col, args.sheet, args.header_row)
                print(f"  ✅ موفق{' در تلاش ' + str(attempt) if attempt > 1 else ''}. ذخیره در: {folder}")
                done_rows.append(row)
                consecutive_failures = 0
            elif progress.get("pdf_path"):
                # خود کار اصلی (PDF) انجام شده؛ فقط مراحل پایانی خطا داشتن
                note = f"انجام شد؛ PDF ذخیره شد ولی مرحله {last_error.step} خطا داشت: {last_error.message}"
                safe_write_status(args.excel, row.row_index, note, GREEN_FILL_ARGB,
                                  args.insurer_col, args.sheet, args.header_row)
                print(f"  ✅ PDF ذخیره شد ({progress['pdf_path']})، ولی مراحل پایانی خطا داشتن.")
                done_rows.append(row)
                consecutive_failures = 0
                recover_ui()
            else:
                note = f"خطا در مرحله {last_error.step} (بعد از {attempt} تلاش): {last_error.message}"
                safe_write_status(args.excel, row.row_index, note, RED_FILL_ARGB,
                                  args.insurer_col, args.sheet, args.header_row)
                failed_rows.append((row, note))
                consecutive_failures += 1
                if last_error.step == 0:
                    print("⛔ فناوران (Bime.exe) در دسترس نیست — اجرا متوقف شد.")
                    break
                if consecutive_failures >= args.max_consecutive_failures:
                    print(f"⛔ {consecutive_failures} ردیف پشت‌سرهم شکست خوردن — احتمالاً مشکل کلی‌تره "
                          f"(فناوران قطع شده یا صفحه عوض شده). اجرا متوقف شد.")
                    break

            if args.confirm:
                resp = input("\n➡️  Enter برای ادامه به ردیف بعد، یا 'q' برای توقف: ").strip().lower()
                if resp == "q":
                    print("متوقف شد.")
                    break
            else:
                time.sleep(1.0)
    except failsafe_exc:
        print("\n⛔ توقف اضطراری (موس به گوشه‌ی صفحه رفت).")
    except KeyboardInterrupt:
        print("\n⛔ با Ctrl+C متوقف شد.")

    print(f"\n{'=' * 55}")
    print(f"پایان: {len(done_rows)} ردیف موفق، {len(failed_rows)} ردیف ناموفق، از {len(rows)} ردیف.")
    for row, note in failed_rows:
        print(f"  ❌ اکسل ردیف {row.row_index} ({row.name}): {note}")


if __name__ == "__main__":
    main()
