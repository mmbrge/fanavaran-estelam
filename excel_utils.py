# -*- coding: utf-8 -*-
"""
excel_utils.py
==============
قدم دوم پروژه: خواندن ردیف‌های زرد از اکسل ورودی و استخراج «کد بیمه‌گذار».

فرض‌ها (اگه فرق داره بگو تا اصلاح کنم):
    - رنگ زرد هایلایت سلول‌ها به شکل استاندارد اکسل (fill آبی-زرد PatternFill) هست.
    - ستونی به اسم چیزی شبیه «بیمه گذار» وجود داره که مقدارش مثل:
          "علی رضایی - 1234567"   یا   "علی رضایی (1234567)"
      هست و کد بعد از اسم اومده. الگوی استخراج کد رو با --code-pattern می‌شه عوض کرد.
    - یه ستون «وضعیت» هم برای نوشتن نتیجه نهایی (خسارت داشت/نداشت و ...) در نظر می‌گیریم؛
      اگه نبود، خودش ستون رو می‌سازه.

این ماژول مستقل تست می‌شه (بدون نیاز به باز بودن فناوران) تا مطمئن بشیم
تشخیص ردیف‌های زرد و استخراج کد درست کار می‌کنه، قبل از قاطی کردنش با اتوماسیون UI.

اجرا برای تست:
    python excel_utils.py --file "ورودی.xlsx" --sheet "Sheet1" --insurer-col "بیمه گذار"
"""

import argparse
import re
import sys
from dataclasses import dataclass
from typing import List, Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# رنگ‌های رایج زرد در اکسل (ARGB). اگه رنگ دقیق فایل شما فرق داره،
# با اجرای --debug-colors رنگ واقعی سلول‌های هایلایت‌شده رو ببین و اینجا اضافه کن.
YELLOW_ARGB_CANDIDATES = {
    "FFFFFF00",  # زرد استاندارد
    "FFFFFF00".lower(),
    "00FFFF00",
    "FFFFF200",
    "FFFFEB00",
}

DEFAULT_CODE_PATTERN = re.compile(r"کد\s*[:\-]?\s*(\d{4,})")  # بعد از کلمه «کد»
FALLBACK_CODE_PATTERN = re.compile(r"(\d{4,})")  # اگه «کد» نبود، اولین رشته عددی

# جدول تبدیل ارقام فارسی/عربی به لاتین. کد استخراج‌شده همیشه به لاتین نرمال می‌شه
# چون بعداً یا توی کلیپ‌بورد کپی می‌شه یا با pyautogui تایپ می‌شه؛ و pyautogui
# فقط کاراکترهای ASCII رو تایپ می‌کنه (ارقام فارسی ۰۱۲.. رو نادیده می‌گیره).
_PERSIAN_TO_WESTERN = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def to_western_digits(text: str) -> str:
    """ارقام فارسی/عربی داخل متن رو به لاتین تبدیل می‌کنه (بقیه کاراکترها دست‌نخورده)."""
    return text.translate(_PERSIAN_TO_WESTERN) if text else text

GREEN_FILL_ARGB = "FF00B050"  # سبز برای ردیف‌های موفق
RED_FILL_ARGB = "FFFF0000"    # قرمز برای ردیف‌هایی که با خطا مواجه شدن


@dataclass
class InsurerRow:
    row_index: int          # شماره ردیف در اکسل (۱-based، شامل هدر)
    raw_value: str          # مقدار خام سلول بیمه‌گذار
    code: Optional[str]     # کد استخراج‌شده
    name: Optional[str]     # بخش اسمی (بدون کد) — تقریبی


def is_yellow(cell) -> bool:
    """چک می‌کنه سلول با رنگ زرد فیل شده یا نه."""
    fill = cell.fill
    if fill is None or fill.fgColor is None:
        return False
    argb = getattr(fill.fgColor, "rgb", None)
    if argb is None:
        return False
    return str(argb).upper() in {c.upper() for c in YELLOW_ARGB_CANDIDATES}


def is_white(cell) -> bool:
    """
    چک می‌کنه سلول رنگی نخورده (سفید/پیش‌فرض) — یعنی fill_type نداره، رنگش
    سفید صریح (FFFFFFFF) هست، یا رنگش از نوع «theme»/«indexed»/«auto» هست (یعنی
    کاربر با دکمه‌ی Fill Color هایلایتش نکرده، فقط رنگ پیش‌فرض تم اکسل رو داره).
    این برای ردیف‌هایی استفاده می‌شه که هنوز پردازش نشدن (بعد از پردازش موفق،
    با رنگ سبز مارک می‌شن پس دیگه سفید نیستن).

    نکته: بعضی فایل‌های اکسل به‌جای رنگ RGB مستقیم، از رنگ «theme» (پیش‌فرض تم)
    برای سلول‌های بدون‌هایلایت استفاده می‌کنن. openpyxl برای این حالت گاهی موقع
    خوندن fgColor.rgb به‌جای None یه مقدار غیرمنتظره برمی‌گردونه، پس این تابع
    اول نوع رنگ (type) رو چک می‌کنه، نه فقط مقدار rgb.
    """
    fill = cell.fill
    if fill is None:
        return True
    fill_type = getattr(fill, "fill_type", None)
    if fill_type is None or fill_type == "none":
        return True

    fg = fill.fgColor
    if fg is None:
        return True

    color_type = getattr(fg, "type", None)
    if color_type != "rgb":
        # theme/indexed/auto: هایلایت صریح RGB نیست، پس سفید/بدون‌هایلایت در نظرش می‌گیریم
        return True

    try:
        argb = fg.rgb
    except Exception:
        return True

    if not isinstance(argb, str):
        return True

    return argb.upper() in {"FFFFFFFF", "00FFFFFF"}


def extract_code(raw_value: str, pattern: re.Pattern = DEFAULT_CODE_PATTERN):
    """
    از متن ستون بیمه‌گذار، کد رو بیرون می‌کشه. اول دنبال الگوی «کد ۱۲۳۴۵۶» می‌گرده
    (که دقیق‌تره چون شماره تلفن/تاریخ داخل اسم رو اشتباهی کد نمی‌گیره)، اگه نبود
    fallback به اولین رشته عددی ۴ رقمی یا بیشتر.
    مثال: 'احمد کمندی کد 12328388' -> کد=12328388، نام='احمد کمندی'
    """
    if not raw_value:
        return None, None

    match = pattern.search(raw_value)
    if match:
        code = to_western_digits(match.group(1))
        name = raw_value[: match.start()].strip(" -()،,")
        return code, name

    # fallback: اگه کلمه «کد» توی متن نبود
    match = FALLBACK_CODE_PATTERN.search(raw_value)
    if match:
        code = to_western_digits(match.group(1))
        name = raw_value[: match.start()].strip(" -()،,")
        return code, name

    return None, raw_value.strip()


def _get_insurer_col_idx(ws, insurer_col_name, header_row):
    target = insurer_col_name.strip()
    # اول دنبال تطابق دقیق هدر می‌گردیم (تا مثلاً موقع جستجوی «بیمه گذار» اشتباهی
    # ستون «کد ملی بیمه گذار» انتخاب نشه که همین رشته رو به‌عنوان زیررشته داره).
    for col_idx in range(1, ws.max_column + 1):
        header_val = ws.cell(row=header_row, column=col_idx).value
        if header_val and str(header_val).strip() == target:
            return col_idx
    # اگه تطابق دقیق نبود، به تطابق زیررشته‌ای برمی‌گردیم
    for col_idx in range(1, ws.max_column + 1):
        header_val = ws.cell(row=header_row, column=col_idx).value
        if header_val and target in str(header_val).strip():
            return col_idx
    raise ValueError(
        f"ستونی با نام شامل '{insurer_col_name}' توی هدر ردیف {header_row} پیدا نشد. "
        f"هدرهای موجود: {[ws.cell(row=header_row, column=c).value for c in range(1, ws.max_column + 1)]}"
    )


def get_yellow_rows(
    filepath: str,
    sheet_name: Optional[str] = None,
    insurer_col_name: str = "بیمه گذار",
    header_row: int = 1,
    code_pattern: re.Pattern = DEFAULT_CODE_PATTERN,
) -> List[InsurerRow]:
    """ردیف‌هایی که سلول ستون بیمه‌گذارشون زرده رو برمی‌گردونه (با کد/نام استخراج‌شده)."""
    return _get_rows_by_predicate(filepath, sheet_name, insurer_col_name, header_row, code_pattern, is_yellow)


def get_white_rows(
    filepath: str,
    sheet_name: Optional[str] = None,
    insurer_col_name: str = "بیمه گذار",
    header_row: int = 1,
    code_pattern: re.Pattern = DEFAULT_CODE_PATTERN,
) -> List[InsurerRow]:
    """
    ردیف‌هایی که سلول ستون بیمه‌گذارشون هنوز سفید/بی‌رنگه (یعنی هنوز پردازش نشدن)
    رو برمی‌گردونه. بعد از پردازش موفق، write_status با رنگ سبز مارکشون می‌کنه
    و دیگه توی اجرای بعدی سفید نیستن.
    """
    return _get_rows_by_predicate(filepath, sheet_name, insurer_col_name, header_row, code_pattern, is_white)


def _get_rows_by_predicate(filepath, sheet_name, insurer_col_name, header_row, code_pattern, predicate):
    wb = load_workbook(filepath, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    insurer_col_idx = _get_insurer_col_idx(ws, insurer_col_name, header_row)

    results: List[InsurerRow] = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        cell = ws.cell(row=row_idx, column=insurer_col_idx)
        if cell.value is None:
            continue
        if predicate(cell):
            code, name = extract_code(str(cell.value), code_pattern)
            results.append(InsurerRow(row_index=row_idx, raw_value=str(cell.value), code=code, name=name))

    return results


def write_status(
    filepath: str,
    row_index: int,
    status_text: str,
    sheet_name: Optional[str] = None,
    status_col_name: str = "وضعیت",
    header_row: int = 1,
    fill_argb: Optional[str] = None,
    also_color_insurer_cell: bool = True,
    insurer_col_name: str = "بیمه گذار",
):
    """
    نتیجه نهایی رو توی ستون «وضعیت» همون ردیف می‌نویسه. اگه fill_argb داده بشه
    (مثلاً GREEN_FILL_ARGB بعد از موفقیت)، هم سلول وضعیت هم (پیش‌فرض) سلول
    بیمه‌گذار همون ردیف رو با اون رنگ فیل می‌کنه — تا دفعه بعد دیگه «سفید» شمرده
    نشه و get_white_rows دوباره پردازشش نکنه.
    """
    from openpyxl.styles import PatternFill

    wb = load_workbook(filepath)
    ws = wb[sheet_name] if sheet_name else wb.active

    status_col_idx = None
    for col_idx in range(1, ws.max_column + 1):
        header_val = ws.cell(row=header_row, column=col_idx).value
        if header_val and status_col_name.strip() in str(header_val).strip():
            status_col_idx = col_idx
            break

    if status_col_idx is None:
        status_col_idx = ws.max_column + 1
        ws.cell(row=header_row, column=status_col_idx, value=status_col_name)

    ws.cell(row=row_index, column=status_col_idx, value=status_text)

    if fill_argb:
        fill = PatternFill(start_color=fill_argb, end_color=fill_argb, fill_type="solid")
        ws.cell(row=row_index, column=status_col_idx).fill = fill
        if also_color_insurer_cell:
            insurer_col_idx = _get_insurer_col_idx(ws, insurer_col_name, header_row)
            ws.cell(row=row_index, column=insurer_col_idx).fill = fill

    wb.save(filepath)


def debug_print_colors(filepath: str, sheet_name: Optional[str], insurer_col_name: str, header_row: int = 1):
    """کمک برای دیباگ: رنگ ARGB واقعی همه سلول‌های ستون بیمه‌گذار رو چاپ می‌کنه."""
    wb = load_workbook(filepath, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    insurer_col_idx = None
    for col_idx in range(1, ws.max_column + 1):
        header_val = ws.cell(row=header_row, column=col_idx).value
        if header_val and insurer_col_name.strip() in str(header_val).strip():
            insurer_col_idx = col_idx
            break
    if insurer_col_idx is None:
        print("ستون بیمه‌گذار پیدا نشد.")
        return
    col_letter = get_column_letter(insurer_col_idx)
    print(f"ستون بیمه‌گذار: {col_letter}\n")
    for row_idx in range(header_row + 1, min(ws.max_row, header_row + 40) + 1):
        cell = ws.cell(row=row_idx, column=insurer_col_idx)
        argb = getattr(cell.fill.fgColor, "rgb", None) if cell.fill else None
        print(f"ردیف {row_idx}: مقدار={cell.value!r}  رنگ={argb}")


def main():
    parser = argparse.ArgumentParser(description="تست خواندن ردیف‌های زرد اکسل و استخراج کد بیمه‌گذار")
    parser.add_argument("--file", required=True, help="مسیر فایل اکسل")
    parser.add_argument("--sheet", default=None, help="نام شیت (پیش‌فرض: شیت فعال)")
    parser.add_argument("--insurer-col", default="بیمه گذار", help="نام ستون بیمه‌گذار در هدر")
    parser.add_argument("--header-row", type=int, default=1, help="شماره ردیف هدر")
    parser.add_argument("--debug-colors", action="store_true", help="چاپ رنگ واقعی سلول‌ها برای دیباگ")
    args = parser.parse_args()

    if args.debug_colors:
        debug_print_colors(args.file, args.sheet, args.insurer_col, args.header_row)
        return

    try:
        rows = get_yellow_rows(args.file, args.sheet, args.insurer_col, args.header_row)
    except ValueError as e:
        print(f"خطا: {e}")
        sys.exit(1)

    print(f"\n{len(rows)} ردیف زرد پیدا شد:\n")
    for r in rows:
        print(f"  ردیف {r.row_index}: خام={r.raw_value!r}  |  کد={r.code}  |  نام={r.name}")


if __name__ == "__main__":
    main()
