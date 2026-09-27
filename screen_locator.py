# -*- coding: utf-8 -*-
"""
screen_locator.py
==================
پیدا کردن دکمه‌ها/آیتم‌های منو روی صفحه از روی تصویرشون (template matching)،
به‌جای مختصات ثابت.

چرا: فرم‌های داخلی فناوران به UI Automation معرفی نشدن، و چیدمان نوار ابزار
بسته به ردیف انتخاب‌شده عوض می‌شه (مثلاً «چاپ بیمه نامه» ↔ «چاپ آزمایشی»)، پس
مختصات ثابت جابه‌جا می‌خوره. با تصویر، دکمه هرجا باشه پیدا می‌شه.

تصویر هر مرحله با capture_template.py ساخته می‌شه و اینجا ذخیره می‌شه:
    templates/<label>.png   ← خود تصویر
    templates/<label>.json  ← نقطه‌ی کلیک نسبت به گوشه‌ی تصویر (و تنظیمات اختیاری)
<label> همون اسم نقطه در coords.json هست (مثلاً "9"، "10").

نیازمند: pip install opencv-python numpy
"""

import json
import os
import time

try:
    import cv2
    import numpy as np
except ImportError:
    cv2 = None
    np = None

try:
    import pyautogui
except ImportError:
    pyautogui = None

TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
DEFAULT_CONFIDENCE = 0.85


def template_path(label: str) -> str:
    return os.path.join(TEMPLATES_DIR, f"{label}.png")


def meta_path(label: str) -> str:
    return os.path.join(TEMPLATES_DIR, f"{label}.json")


def has_template(label: str) -> bool:
    return os.path.exists(template_path(label))


def available() -> bool:
    return cv2 is not None and pyautogui is not None


def load_meta(label: str) -> dict:
    try:
        with open(meta_path(label), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _load_gray(path: str):
    # cv2.imread روی ویندوز مسیرهای غیرلاتین رو باز نمی‌کنه؛ پس fromfile + imdecode
    data = np.fromfile(path, dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_GRAYSCALE)
    if img is None:
        raise IOError(f"تصویر خونده نشد: {path}")
    return img


def to_gray(pil_image):
    return cv2.cvtColor(np.array(pil_image.convert("RGB")), cv2.COLOR_RGB2GRAY)


def match(template_gray, screen_gray):
    """بهترین تطابق: (score, (x, y)) — x, y گوشه‌ی بالا-چپ نسبت به screen_gray."""
    th, tw = template_gray.shape[:2]
    sh, sw = screen_gray.shape[:2]
    if sh < th or sw < tw:
        return 0.0, None
    res = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)
    return float(max_val), max_loc


def count_other_matches(template_gray, screen_gray, best_xy, confidence=DEFAULT_CONFIDENCE) -> int:
    """
    چند جای دیگه‌ی صفحه (غیر از best_xy) هم با این تصویر تطابق دارن؟ اگه > 0 باشه،
    تصویر مبهمه (مثلاً یه فلش ▼ تنها که چند جای صفحه هست).
    """
    th, tw = template_gray.shape[:2]
    res = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
    ys, xs = np.where(res >= confidence)
    others = []
    for x, y in zip(xs, ys):
        if abs(x - best_xy[0]) <= tw // 2 and abs(y - best_xy[1]) <= th // 2:
            continue
        if any(abs(x - ox) <= tw // 2 and abs(y - oy) <= th // 2 for ox, oy in others):
            continue
        others.append((x, y))
    return len(others)


def locate(label: str, region=None, confidence=None, timeout=8.0, interval=0.3):
    """
    تصویر مرحله‌ی label رو روی صفحه پیدا می‌کنه و نقطه‌ی کلیکش رو برمی‌گردونه:
    (x, y, score) با مختصات مطلق صفحه. تا timeout ثانیه منتظر ظاهر شدنش می‌مونه
    (مثلاً آیتم منویی که تازه داره باز می‌شه). اگه پیدا نشد LookupError.
    region: (left, top, width, height) برای محدود کردن جستجو (اختیاری).
    """
    if not available():
        raise LookupError("opencv-python یا pyautogui نصب نیست: pip install opencv-python numpy")
    meta = load_meta(label)
    conf = confidence or meta.get("confidence", DEFAULT_CONFIDENCE)
    template = _load_gray(template_path(label))
    th, tw = template.shape[:2]
    dx, dy = meta.get("click", [tw // 2, th // 2])
    ox, oy = (region[0], region[1]) if region else (0, 0)

    best = 0.0
    end = time.time() + timeout
    while True:
        screen = to_gray(pyautogui.screenshot(region=region))
        score, loc = match(template, screen)
        best = max(best, score)
        if loc is not None and score >= conf:
            # اگه همین تصویر جای دیگه‌ی صفحه هم هست، کلیک نکن: ممکنه روی دکمه‌ی
            # اشتباه (مثلاً یه ▼ دیگه) بخوره. تصویر رو باید با متن کنارش دوباره ساخت.
            if not meta.get("allow_multiple"):
                others = count_other_matches(template, screen, loc, conf)
                if others:
                    raise LookupError(
                        f"تصویر «{label}» مبهمه: {others + 1} جای صفحه پیدا شد و معلوم نیست کدوم درسته. "
                        f"دوباره با capture_template.py --label {label} بساز و کادر رو بزرگ‌تر بگیر "
                        f"(همراه متن/آیکون کنارش) تا یکتا بشه."
                    )
            return ox + loc[0] + dx, oy + loc[1] + dy, score
        if time.time() >= end:
            raise LookupError(
                f"تصویر «{label}» تا {timeout:.0f} ثانیه روی صفحه پیدا نشد "
                f"(بهترین شباهت: {best:.2f}، حداقل لازم: {conf:.2f})."
            )
        time.sleep(interval)


def is_visible(label: str, confidence=None) -> bool:
    """آیا تصویر این مرحله همین الان روی صفحه دیده می‌شه؟ (یه‌بار چک، بدون صبر)"""
    meta = load_meta(label)
    conf = confidence or meta.get("confidence", DEFAULT_CONFIDENCE)
    template = _load_gray(template_path(label))
    score, loc = match(template, to_gray(pyautogui.screenshot()))
    return loc is not None and score >= conf


def wait_visible(label: str, timeout: float, interval: float = 0.3) -> bool:
    """تا timeout ثانیه منتظر می‌مونه تصویر ظاهر بشه. خروجی: True اگه ظاهر شد."""
    end = time.time() + timeout
    while True:
        if is_visible(label):
            return True
        if time.time() >= end:
            return False
        time.sleep(interval)


def wait_gone(label: str, timeout: float, interval: float = 0.3) -> bool:
    """تا timeout ثانیه منتظر می‌مونه تصویر از صفحه بره (مثلاً پنجره‌ای بسته بشه)."""
    end = time.time() + timeout
    while True:
        if not is_visible(label):
            return True
        if time.time() >= end:
            return False
        time.sleep(interval)
