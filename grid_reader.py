# -*- coding: utf-8 -*-
"""
grid_reader.py
===============
با استفاده از مختصات کالیبره‌شده در calibrate_grid.py، هر ردیف ستون «شماره
الحاقیه» رو با OCR می‌خونه و آخرین ردیفی که مقدارش 0 هست رو پیدا می‌کنه.

نیازمند: pytesseract + نصب جداگانه Tesseract-OCR روی ویندوز
(https://github.com/UB-Mannheim/tesseract/wiki)
"""

import os
import sys
import time

try:
    import pyautogui
    from PIL import ImageOps
except ImportError:
    print("pyautogui یا Pillow نصب نیست: pip install pyautogui pillow")
    sys.exit(1)

try:
    import pytesseract
    # اگه tesseract روی PATH نیست، مسیر پیش‌فرض نصب ویندوز رو امتحان کن (fallback)
    _default_win_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(_default_win_path):
        pytesseract.pytesseract.tesseract_cmd = _default_win_path
except ImportError:
    pytesseract = None

try:
    from paddleocr import PaddleOCR
    import numpy as np
except ImportError:
    PaddleOCR = None

# PaddleOCR 3.x عوض کردن API (.ocr -> .predict با خروجی متفاوت) رو مدام تغییر داده و
# غیرقابل‌اعتماد شده؛ pytesseract با زبان فارسی دقت خوبی داد، پس فعلاً همینو اصلی می‌کنیم.
USE_PADDLE = False

_paddle_instance = None


def _get_paddle_ocr():
    """یه نمونه PaddleOCR می‌سازه (فقط بار اول - بارگذاری مدل کندتره، بارهای بعدی سریع)."""
    global _paddle_instance
    if _paddle_instance is None:
        print("  ⏳ بارگذاری مدل PaddleOCR (فقط بار اول کند است)...")
        try:
            # نسخه‌های جدید (3.x): show_log حذف شده، use_angle_cls -> use_textline_orientation
            _paddle_instance = PaddleOCR(use_textline_orientation=False, lang="ar")
        except TypeError:
            try:
                # نسخه‌های قدیمی‌تر
                _paddle_instance = PaddleOCR(use_angle_cls=False, lang="ar", show_log=False)
            except TypeError:
                _paddle_instance = PaddleOCR(lang="ar")
    return _paddle_instance

MAX_ROWS_TO_SCAN = 60

# اگه مقدار بگیره (run_no_claim_flow.py با --debug ستش می‌کنه)، از هر پاس خوندن
# گرید یه عکس علامت‌گذاری‌شده ذخیره می‌شه: کادر قرمز = ناحیه‌ای که OCR می‌خونه،
# نقطه‌ی سبز = جایی که برای انتخاب ردیف کلیک می‌شه، عدد کنارش = نتیجه‌ی OCR.
DEBUG_DIR = None
_debug_counter = 0


def _save_debug_image(screenshot, marks):
    """marks: لیست (crop_box, text, click_xy) نسبت به گوشه‌ی پنجره."""
    global _debug_counter
    try:
        from PIL import ImageDraw
        os.makedirs(DEBUG_DIR, exist_ok=True)
        img = screenshot.convert("RGB").copy()
        draw = ImageDraw.Draw(img)
        for box, text, (cx, cy) in marks:
            draw.rectangle(box, outline=(255, 0, 0), width=2)
            draw.ellipse((cx - 4, cy - 4, cx + 4, cy + 4), fill=(0, 200, 0))
            label = text if text else "-"
            draw.text((box[0] - 10 - 7 * len(label), box[1] + 4), label, fill=(255, 0, 0))
        _debug_counter += 1
        path = os.path.join(DEBUG_DIR, f"grid_{time.strftime('%H%M%S')}_{_debug_counter}.png")
        img.save(path)
        print(f"  🖼 عکس عیب‌یابی گرید: {path}")
    except Exception as e:
        print(f"  ⚠ ذخیره‌ی عکس عیب‌یابی نشد: {e}")
EMPTY_ROWS_TO_STOP = 3  # اگه این‌همه ردیف پشت‌سرهم خالی بود، یعنی به ته گرید رسیدیم
MAX_SCROLL_PASSES = 15  # سقف تعداد اسکرول برای جلوگیری از حلقه بی‌نهایت
SCROLL_BURSTS_PER_PASS = 3  # چندتا scroll(-15) پشت‌سرهم بزنیم (تشخیص همپوشانی خودش تکراری‌ها رو مدیریت می‌کنه)

PERSIAN_TO_WESTERN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
ARABIC_TO_WESTERN_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


def normalize_digits(text: str) -> str:
    """ارقام فارسی/عربی (۰۱۲.. یا ٠١٢..) رو به ارقام لاتین تبدیل می‌کنه و فقط عدد رو نگه می‌داره."""
    text = text.translate(PERSIAN_TO_WESTERN_DIGITS).translate(ARABIC_TO_WESTERN_DIGITS)
    digits_only = "".join(ch for ch in text if ch.isdigit())
    return digits_only


def ocr_digits(image):
    """
    یه تصویر (PIL Image، از قبل crop شده روی یه سلول تکی) رو OCR می‌کنه و فقط
    ارقام (لاتین‌شده) رو برمی‌گردونه. اول با PaddleOCR (دقت بهتر برای فارسی)،
    اگه نصب نبود، fallback به pytesseract.
    """
    # سلول‌های گرید معمولاً خیلی کوچیکن (۷۰x۳۷ پیکسل)؛ بزرگ‌نمایی + کنتراست دقت رو زیاد می‌کنه
    scale = 4
    image = image.resize((image.width * scale, image.height * scale))
    image = ImageOps.autocontrast(image.convert("L")).convert("RGB")

    if USE_PADDLE and PaddleOCR is not None:
        try:
            ocr = _get_paddle_ocr()
            img_array = np.array(image)
            try:
                result = ocr.ocr(img_array, det=False, cls=False)
            except TypeError:
                result = ocr.ocr(img_array, det=False)
            raw_text = ""
            if result and result[0]:
                raw_text = result[0][0][0]
            return normalize_digits(raw_text)
        except Exception as e:
            print(f"  ⚠ PaddleOCR خطا داد ({e})، فالبک به pytesseract...")

    if pytesseract is not None:
        config = "--psm 7 -c tessedit_char_whitelist=۰۱۲۳۴۵۶۷۸۹0123456789"
        try:
            raw_text = pytesseract.image_to_string(image, lang="fas", config=config)
        except Exception:
            raw_text = pytesseract.image_to_string(image, config=config)
        result = normalize_digits(raw_text)
        if result:
            return result
        # اگه psm 7 چیزی نگرفت، fallback به psm 10 (تک کاراکتر)
        config_fallback = "--psm 10 -c tessedit_char_whitelist=۰۱۲۳۴۵۶۷۸۹0123456789"
        try:
            raw_text = pytesseract.image_to_string(image, lang="fas", config=config_fallback)
        except Exception:
            raw_text = pytesseract.image_to_string(image, config=config_fallback)
        return normalize_digits(raw_text)

    raise RuntimeError("نه PaddleOCR نه pytesseract در دسترس نیست. حداقل یکی رو نصب کن.")


def read_grid_rows(window_rect, grid_config, max_rows=MAX_ROWS_TO_SCAN):
    """
    اسکرین‌شات کل پنجره رو می‌گیره و ردیف‌به‌ردیف سلول «شماره الحاقیه» رو OCR
    می‌کنه. لیستی از (row_index, text, abs_click_x, abs_click_y) برمی‌گردونه.
    """
    left, top, right, bottom = window_rect
    screenshot = pyautogui.screenshot(region=(left, top, right - left, bottom - top))

    col_left = grid_config["col_left"]
    col_right = grid_config["col_right"]
    col_top0 = grid_config["col_top"]
    col_bottom0 = grid_config["col_bottom"]
    row_height = grid_config["row_height"]
    row_click_x = grid_config["row_click_x"]
    row_click_y_offset = grid_config["row_click_y_offset"]
    cell_height = col_bottom0 - col_top0
    # مرز واقعی پایین گرید (قبل از فوتر/اسکرول‌بار)؛ اگه کالیبره نشده بود، کل پنجره
    grid_bottom_y = grid_config.get("grid_bottom_y", bottom - top)

    results = []
    marks = []
    consecutive_empty = 0

    for row_idx in range(max_rows):
        cell_top = col_top0 + row_idx * row_height
        cell_bottom = cell_top + cell_height

        if cell_bottom > grid_bottom_y:
            break  # به مرز واقعی گرید رسیدیم (قبل از فوتر/اسکرول‌بار)

        crop_box = (col_left, cell_top, col_right, cell_bottom)
        try:
            cell_img = screenshot.crop(crop_box)
            text = ocr_digits(cell_img)
        except Exception as e:
            print(f"  ⚠ خطای OCR توی ردیف {row_idx}: {e}")
            text = ""

        abs_click_x = left + row_click_x
        abs_click_y = top + col_top0 + row_idx * row_height + row_click_y_offset

        if text == "":
            consecutive_empty += 1
        else:
            consecutive_empty = 0

        results.append((row_idx, text, abs_click_x, abs_click_y))
        marks.append((crop_box, text, (abs_click_x - left, abs_click_y - top)))

        if consecutive_empty >= EMPTY_ROWS_TO_STOP:
            # چندتا ردیف خالی پشت‌سرهم یعنی به انتهای داده‌ها رسیدیم
            results = results[: -EMPTY_ROWS_TO_STOP]
            break

    if DEBUG_DIR:
        _save_debug_image(screenshot, marks)
    return results


def find_last_zero_row(window_rect, grid_config, max_rows=MAX_ROWS_TO_SCAN, debug=True):
    """آخرین ردیفی که مقدار OCR‌شده‌اش دقیقاً '0' هست رو برمی‌گردونه، یا None. (بدون اسکرول)"""
    rows = read_grid_rows(window_rect, grid_config, max_rows)

    if debug:
        print(f"  {len(rows)} ردیف اسکن شد:")
        for row_idx, text, x, y in rows:
            print(f"    ردیف {row_idx}: OCR='{text}'  ({x},{y})")

    zero_rows = [(idx, x, y) for idx, text, x, y in rows if text == "0"]
    if not zero_rows:
        return None
    return zero_rows[-1]  # آخرین ردیفی که صفره


def read_grid_rows_with_end_flag(window_rect, grid_config, max_rows=MAX_ROWS_TO_SCAN):
    """
    مثل read_grid_rows، ولی اضافه می‌گه آیا توی همین صفحه به «ته داده‌ها» رسیدیم
    یا نه (یعنی ردیف‌های خالی پیدا شدن، پس دیگه لازم نیست اسکرول کنیم).
    برمی‌گردونه: (rows, reached_end: bool)
    """
    left, top, right, bottom = window_rect
    col_top0 = grid_config["col_top"]
    row_height = grid_config["row_height"]
    grid_bottom_y = grid_config.get("grid_bottom_y", bottom - top)
    max_possible_rows = int((grid_bottom_y - col_top0) // row_height) + 1

    rows = read_grid_rows(window_rect, grid_config, max_rows)
    # اگه تعداد ردیف‌های خونده‌شده از حداکثر ممکن (بر اساس مرز گرید) کمتره،
    # یعنی خودِ read_grid_rows زودتر (به‌خاطر ردیف‌های خالی) متوقف شده -> ته داده‌هاست
    reached_end = len(rows) < max_possible_rows
    return rows, reached_end


def find_overlap(prev_texts, new_texts):
    """
    بزرگ‌ترین k رو پیدا می‌کنه که «k تای آخر prev_texts» با «k تای اول new_texts»
    یکی باشن - یعنی بعد از اسکرول چندتا ردیف مشترک (تکراری) هستن.
    """
    max_k = min(len(prev_texts), len(new_texts))
    for k in range(max_k, 0, -1):
        if prev_texts[-k:] == new_texts[:k]:
            return k
    return 0


def find_last_zero_row_with_scroll(window_rect, grid_config, debug=True):
    """
    آخرین ردیفی که مقدارش 0 هست رو پیدا می‌کنه، با اسکرول تطبیقی: اگه همون پاس
    اول به ته داده‌ها رسیده باشه اصلاً اسکرول نمی‌کنه؛ وگرنه هر پاس محتوای جدید
    رو با پاس قبلی مقایسه می‌کنه (تشخیص همپوشانی) تا هیچ ردیفی نه دوبار پردازش
    بشه نه جا بمونه. مختصات کلیک همیشه برای وضعیت *فعلی* اسکرول معتبره.
    """
    left, top, right, bottom = window_rect
    hover_x = left + (grid_config["col_left"] + grid_config["col_right"]) // 2
    hover_y = top + grid_config["col_top"] + grid_config["row_height"] * 3

    all_rows = []  # لیست دیده‌بان کل (بدون تکرار): (global_idx, text, click_x, click_y)
    prev_texts = []
    best_zero = None  # (global_idx, click_x, click_y)
    pass_idx = 0

    while pass_idx < MAX_SCROLL_PASSES:
        rows, reached_end = read_grid_rows_with_end_flag(window_rect, grid_config)
        pass_texts = [text for _, text, _, _ in rows]

        overlap = find_overlap(prev_texts, pass_texts) if prev_texts else 0
        new_rows = rows[overlap:]

        if debug:
            print(f"  --- پاس {pass_idx}: {len(rows)} ردیف دیده شد، {overlap} تا مشترک، {len(new_rows)} تا جدید ---")
            for _, text, _, _ in new_rows:
                print(f"    جدید: OCR='{text}'")

        for _, text, cx, cy in new_rows:
            global_idx = len(all_rows)
            all_rows.append((global_idx, text, cx, cy))
            if text == "0":
                best_zero = (global_idx, cx, cy)

        prev_texts = pass_texts

        if reached_end:
            if debug:
                print("  ✅ به ته داده‌ها رسیدیم (فضای خالی دیده شد).")
            break

        if not new_rows and pass_idx > 0:
            # اسکرول کردیم ولی هیچ ردیف جدیدی نیومد -> احتمالاً واقعاً ته گرید بدون فضای خالی
            if debug:
                print("  ✅ اسکرول تغییری در محتوا ایجاد نکرد - توقف.")
            break

        # اسکرول به پایین (چندتا burst پشت‌سرهم چون یه scroll تنها فقط ~۱ ردیف جابه‌جا می‌کرد)
        pyautogui.moveTo(hover_x, hover_y)
        for _ in range(SCROLL_BURSTS_PER_PASS):
            pyautogui.scroll(-15)
            time.sleep(0.05)
        time.sleep(0.5)
        pass_idx += 1

    if best_zero is None:
        return None

    # مختصات click ثبت‌شده برای اون ردیف مال آخرین پاسی هست که توش دیده شد. اگه از اون
    # موقع اسکرول بیشتری انجام شده (یعنی ردیف صفر مال یه پاس قبلی بود، نه آخرین پاس)،
    # باید به همون مقدار برگردیم بالا تا دوباره قابل کلیک باشه.
    global_idx, click_x, click_y = best_zero
    rows_scrolled_past_target = (len(all_rows) - 1) - global_idx
    # هر burst تقریباً ۱ ردیف جابه‌جا می‌کنه (طبق مشاهده)
    if pass_idx > 0 and rows_scrolled_past_target > 0:
        if debug:
            print(f"  ↩ برگشت به بالا برای رسیدن به ردیف هدف (global_idx={global_idx})...")
        pyautogui.moveTo(hover_x, hover_y)
        for _ in range(rows_scrolled_past_target + 2):  # کمی بیشتر برای اطمینان
            pyautogui.scroll(15)
            time.sleep(0.05)
        time.sleep(0.5)
        rows, _ = read_grid_rows_with_end_flag(window_rect, grid_config)
        # دنبال ردیفی با همون متن '0' که به‌ترتیب با global_idx فاصله منطقی داره بگرد؛
        # ساده‌ترین حالت قابل‌اعتماد: آخرین ردیف صفر توی همین دید فعلی
        zero_rows_now = [(idx, x, y) for idx, text, x, y in rows if text == "0"]
        if zero_rows_now:
            _, click_x, click_y = zero_rows_now[-1]

    return (global_idx, click_x, click_y)





if __name__ == "__main__":
    # تست مستقل
    import argparse
    import json
    import os
    import win32gui
    import win32process
    import psutil

    COORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "coords.json")
    TARGET_EXE = "Bime.exe"

    def find_bime_window_rect():
        target_pids = {p.info["pid"] for p in psutil.process_iter(["pid", "name"])
                       if TARGET_EXE.lower() in (p.info.get("name") or "").lower()}
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
        found.sort(key=lambda f: (f[1][2] - f[1][0]) * (f[1][3] - f[1][1]), reverse=True)
        return found[0][1]

    with open(COORDS_FILE, "r", encoding="utf-8") as f:
        coords = json.load(f)

    rect = find_bime_window_rect()
    result = find_last_zero_row_with_scroll(rect, coords["grid"])
    if result:
        print(f"\n✅ آخرین ردیف صفر: index={result[0]}  کلیک روی ({result[1]}, {result[2]})")
    else:
        print("\n❌ هیچ ردیفی با مقدار 0 پیدا نشد.")
