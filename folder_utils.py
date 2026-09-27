# -*- coding: utf-8 -*-
"""
folder_utils.py
================
توابع کمکی برای ساخت مسیر پوشه‌ی ذخیره‌سازی بر اساس تاریخ شمسی جاری، و پاک‌سازی
اسم شخص برای استفاده به‌عنوان اسم پوشه (حذف کاراکترهای غیرمجاز ویندوز).

این فایل هیچ وابستگی بیرونی (مثل jdatetime یا persiantools) نداره — تبدیل
میلادی به شمسی با یک الگوریتم مستقل پیاده‌سازی شده تا نیازی به نصب پکیج
اضافه نباشه.

توابعی که run_no_claim_flow.py از این فایل استفاده می‌کنه:
    - current_persian_year_month()
    - sanitize_folder_name(name)
"""

import datetime
import re

PERSIAN_MONTH_NAMES = [
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
]

# کاراکترهایی که ویندوز توی اسم فایل/پوشه اجازه نمی‌ده
_INVALID_CHARS_PATTERN = re.compile(r'[\\/:*?"<>|]')
_MULTI_SPACE_PATTERN = re.compile(r"\s+")


def _is_leap_gregorian(year: int) -> bool:
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


def gregorian_to_jalali(g_year: int, g_month: int, g_day: int):
    """
    تبدیل تاریخ میلادی به شمسی (الگوریتم استاندارد و شناخته‌شده، بدون نیاز به
    کتابخونه بیرونی). خروجی: (jalali_year, jalali_month, jalali_day)
    """
    g_days_in_month = [31, 29 if _is_leap_gregorian(g_year) else 28,
                        31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    gy = g_year - 1600
    gm = g_month - 1
    gd = g_day - 1

    g_day_no = 365 * gy + (gy + 3) // 4 - (gy + 99) // 100 + (gy + 399) // 400
    for i in range(gm):
        g_day_no += g_days_in_month[i]
    g_day_no += gd

    j_day_no = g_day_no - 79

    j_np = j_day_no // 12053
    j_day_no %= 12053

    jy = 979 + 33 * j_np + 4 * (j_day_no // 1461)
    j_day_no %= 1461

    if j_day_no >= 366:
        jy += (j_day_no - 1) // 365
        j_day_no = (j_day_no - 1) % 365

    j_days_in_month = [31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
    jm = 0
    for i in range(11):
        if j_day_no < j_days_in_month[i]:
            jm = i
            break
        j_day_no -= j_days_in_month[i]
        jm = i + 1
    jd = j_day_no + 1

    return jy, jm + 1, jd


def current_persian_year_month():
    """
    سال و ماه شمسی جاری رو برمی‌گردونه (برای ساخت مسیر پوشه).
    خروجی: (year: int, month_folder_name: str)
    مثال: (1404, "07-مهر")
    """
    today = datetime.date.today()
    jy, jm, _ = gregorian_to_jalali(today.year, today.month, today.day)
    month_name = PERSIAN_MONTH_NAMES[jm - 1]
    month_folder_name = f"{jm:02d}-{month_name}"
    return jy, month_folder_name


def sanitize_folder_name(name: str) -> str:
    """
    اسم بیمه‌گذار رو برای استفاده به‌عنوان اسم پوشه پاک‌سازی می‌کنه:
    - کاراکترهای غیرمجاز ویندوز (\\ / : * ? " < > |) حذف می‌شن
    - فاصله‌های اضافه/چندتایی یکی می‌شن
    - از ابتدا/انتها فاصله و نقطه حذف می‌شه (ویندوز اسم ختم‌شده به نقطه/فاصله نمی‌پذیره)
    - طول نهایی به ۱۰۰ کاراکتر محدود می‌شه (جلوگیری از خطای طول مسیر ویندوز)
    """
    if not name:
        return ""

    cleaned = _INVALID_CHARS_PATTERN.sub("", name)
    cleaned = _MULTI_SPACE_PATTERN.sub(" ", cleaned).strip()
    cleaned = cleaned.strip(" .")
    return cleaned[:100]


if __name__ == "__main__":
    # تست سریع و مستقل
    y, m = current_persian_year_month()
    print(f"سال/ماه شمسی جاری: {y} / {m}")

    tests = [
        "علی رضایی",
        'محمد "حسینی"',
        "زهرا/کریمی:۱۲۳",
        "   نام با فاصله زیاد   ",
        "پایان با نقطه.",
    ]
    for t in tests:
        print(f"{t!r} -> {sanitize_folder_name(t)!r}")
