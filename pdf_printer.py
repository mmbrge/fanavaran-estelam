# -*- coding: utf-8 -*-
"""
pdf_printer.py
===============
پرینت و ذخیره‌ی PDF استعلام خسارت (مراحل ۱۱، ۱۲، ۱۳ و پنجره‌ی Save) بدون
مختصات، با UI Automation.

خروجی inspect_window.py نشون داد که برخلاف فرم‌های داخلی فناوران (که به UIA
معرفی نشدن)، نمایشگر گزارش، پنجره‌ی Print و پنجره‌ی Save ویندوز کامل
قابل‌شناسایی‌ان:
    - نمایشگر گزارش:  auto_id="FastReportViewerUserControl"
    - دکمه‌ی Print:   داخل auto_id="toolBar"، با title="Print"
    - پنجره‌ی Print:  auto_id="PrinterSetupForm"  (پرینتر: cbxPrinter، تأیید: btnOk)
    - پنجره‌ی Save:   فیلد نام فایل auto_id="1001"، دکمه‌ی Save با auto_id="1"

مزیت‌ها نسبت به روش مختصاتی:
    - به اندازه‌ی پنجره، رزولوشن و جای دیالوگ‌ها وابسته نیست.
    - به‌جای sleepهای ثابت، منتظر ظاهر/بسته شدن واقعی هر پنجره می‌مونه.
    - مسیر فارسی مستقیم توی فیلد نوشته می‌شه (نه typewrite، نه کلیپ‌بورد).
    - آخر کار چک می‌کنه فایل PDF واقعاً روی دیسک ساخته شده باشه.
    - دکمه‌های دیالوگ‌ها با پیام ویندوز (BM_CLICK) زده می‌شن، نه حرکت موس؛
      فقط اگه جواب نداد، کلیک واقعی روی محل همون دکمه (که UIA پیدا کرده) انجام می‌شه.

تست مستقل (با پنجره‌ی نمایشگر استعلام خسارت باز):
    python pdf_printer.py --out "C:\\temp\\test.pdf"
"""

import os
import time

try:
    from pywinauto import Application
except ImportError:
    Application = None

try:
    import win32api
    import win32gui
except ImportError:
    win32api = None
    win32gui = None

PDF_PRINTER_NAME = "Microsoft Print to PDF"
BM_CLICK = 0x00F5


class PrintError(Exception):
    """خطایی که مشخص می‌کنه توی کدوم مرحله‌ی پرینت/ذخیره گیر کردیم."""
    def __init__(self, step, message):
        self.step = step
        self.message = message
        super().__init__(f"مرحله {step}: {message}")


# ---------- کمکی‌ها ----------

def _wait_until(predicate, timeout, interval=0.25):
    """تا وقتی predicate درست بشه (یا timeout برسه) صبر می‌کنه. خروجی: True/False."""
    end = time.time() + timeout
    while True:
        try:
            if predicate():
                return True
        except Exception:
            pass
        if time.time() >= end:
            return False
        time.sleep(interval)


def _is_gone(wrapper) -> bool:
    """آیا این پنجره/کنترل بسته شده (یا دیگه دیده نمی‌شه)؟"""
    handle = getattr(wrapper, "handle", None)
    if handle and win32gui is not None:
        return not (win32gui.IsWindow(handle) and win32gui.IsWindowVisible(handle))
    try:
        return not wrapper.is_visible()
    except Exception:
        return True


def _press(ctrl, done, timeout, step, label):
    """
    یه دکمه رو می‌زنه و منتظر نتیجه‌ش (done) می‌مونه.
    اول با پیام BM_CLICK (بدون حرکت موس)؛ اگه تا ۳ ثانیه نتیجه نداد، کلیک واقعی
    روی مستطیل همون دکمه. از invoke() استفاده نمی‌کنیم چون روی دکمه‌هایی که
    پنجره‌ی مودال باز می‌کنن ممکنه تا بسته شدن اون پنجره قفل بمونه.
    """
    handle = getattr(ctrl, "handle", None)
    if handle and win32api is not None:
        try:
            try:
                ctrl.set_focus()
            except Exception:
                pass
            win32api.PostMessage(handle, BM_CLICK, 0, 0)
            if _wait_until(done, 3):
                return
        except Exception:
            pass

    ctrl.click_input()
    if not _wait_until(done, timeout):
        raise PrintError(step, f"بعد از زدن دکمه‌ی «{label}» اتفاق مورد انتظار نیفتاد.")


def _combo_text(cbx) -> str:
    try:
        return (cbx.selected_text() or "").strip()
    except Exception:
        return (cbx.window_text() or "").strip()


def _find_window_spec(app, main, timeout, **criteria):
    """
    یه پنجره‌ی دیالوگ رو پیدا می‌کنه؛ هم به‌عنوان فرزند پنجره‌ی اصلی (همون‌طور که
    توی خروجی inspect دیده شد) و هم به‌عنوان پنجره‌ی سطح‌بالای جدا.
    """
    specs = [main.child_window(**criteria), app.window(**criteria)]
    found = []

    def check():
        for s in specs:
            if s.exists(timeout=0):
                found.append(s)
                return True
        return False

    if _wait_until(check, timeout):
        return found[0]
    return None


def _ancestor_window(ctrl):
    """نزدیک‌ترین والد از نوع Window (مثلاً خود دیالوگ Save برای فیلد نام فایل)."""
    p = ctrl.parent()
    while p is not None:
        try:
            if p.element_info.control_type == "Window":
                return p
        except Exception:
            pass
        p = p.parent()
    return None


def unique_pdf_path(folder: str, base_name: str = "استعلام خسارت") -> str:
    """
    مسیر فایل PDF رو برمی‌گردونه؛ اگه از قبل وجود داشت، (2)، (3)، ... اضافه می‌کنه تا
    فایل قبلی بازنویسی نشه (و دیالوگ «Confirm Save As» هم باز نشه).
    """
    path = os.path.join(folder, base_name + ".pdf")
    i = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{base_name} ({i}).pdf")
        i += 1
    return path


# ---------- مراحل ----------

def connect_main(hwnd):
    if Application is None:
        raise PrintError(11, "pywinauto نصب نیست: pip install pywinauto")
    app = Application(backend="uia").connect(handle=hwnd, timeout=10)
    main = app.window(handle=hwnd)
    return app, main


def open_print_dialog(app, main, viewer_timeout=30):
    """مرحله ۱۱: منتظر باز شدن نمایشگر گزارش می‌مونه و دکمه‌ی Print رو می‌زنه."""
    viewer = main.child_window(auto_id="FastReportViewerUserControl")
    if not viewer.exists(timeout=viewer_timeout):
        raise PrintError(11, f"نمایشگر گزارش استعلام خسارت تا {viewer_timeout} ثانیه باز نشد.")

    print_btn = viewer.child_window(auto_id="toolBar", control_type="ToolBar") \
                      .child_window(title="Print", control_type="Button")
    # تا وقتی گزارش کامل رندر نشده، دکمه‌ی Print غیرفعاله
    if not _wait_until(lambda: print_btn.exists(timeout=0) and print_btn.is_enabled(), viewer_timeout):
        raise PrintError(11, "دکمه‌ی Print نمایشگر گزارش پیدا نشد یا فعال نشد.")

    dlg_holder = []

    def dialog_open():
        spec = _find_window_spec(app, main, 0, auto_id="PrinterSetupForm", control_type="Window")
        if spec is not None:
            dlg_holder.append(spec)
            return True
        return False

    _press(print_btn.wrapper_object(), dialog_open, 15, 11, "Print")
    # spec برمی‌گردونیم نه wrapper، چون child_window فقط روی spec وجود داره
    return dlg_holder[-1]


def select_printer(dlg, printer_name=PDF_PRINTER_NAME):
    """مرحله ۱۲: انتخاب پرینتر PDF و گزینه‌ی «All» (همه‌ی صفحات). dlg = spec پنجره‌ی Print."""
    cbx = dlg.child_window(auto_id="cbxPrinter", control_type="ComboBox").wrapper_object()
    if _combo_text(cbx) != printer_name:
        try:
            cbx.select(printer_name)
        except Exception:
            try:
                available = [t for t in cbx.texts() if t]
            except Exception:
                available = []
            raise PrintError(12, f"پرینتر «{printer_name}» توی لیست نبود. پرینترهای موجود: {available}")
        if not _wait_until(lambda: _combo_text(cbx) == printer_name, 3):
            raise PrintError(12, f"پرینتر «{printer_name}» انتخاب نشد (الان: «{_combo_text(cbx)}»).")

    try:
        rb_all = dlg.child_window(auto_id="rbAll", control_type="RadioButton").wrapper_object()
        if not rb_all.is_selected():
            rb_all.select()
    except Exception:
        pass  # اختیاریه؛ پیش‌فرض خود فرم معمولاً All هست


def confirm_print(dlg):
    """مرحله ۱۳: زدن دکمه‌ی Print پنجره‌ی پرینت (تا بسته شدن همون پنجره)."""
    dlg_wrapper = dlg.wrapper_object()
    ok_btn = dlg.child_window(auto_id="btnOk", control_type="Button").wrapper_object()
    _press(ok_btn, lambda: _is_gone(dlg_wrapper), 10, 13, "Print")


def save_pdf(app, pdf_path, dialog_timeout=40, file_timeout=60):
    """پنجره‌ی «Save Print Output As»: نوشتن مسیر، زدن Save و چک ساخته شدن فایل."""
    holder = []

    def find_edit():
        for w in app.windows():
            spec = app.window(handle=w.handle).child_window(auto_id="1001", control_type="Edit")
            if spec.exists(timeout=0):
                holder.append(spec.wrapper_object())
                return True
        return False

    if not _wait_until(find_edit, dialog_timeout, interval=0.5):
        raise PrintError("ذخیره", f"پنجره‌ی Save تا {dialog_timeout} ثانیه باز نشد.")
    edit = holder[-1]

    save_dlg = _ancestor_window(edit)
    if save_dlg is None:
        raise PrintError("ذخیره", "پنجره‌ی Save پیدا شد ولی دکمه‌ی Save‌اش پیدا نشد.")

    # نوشتن مستقیم مسیر (با ValuePattern؛ فارسی بدون مشکل)
    edit.set_edit_text(pdf_path)
    try:
        written = edit.get_value()
    except Exception:
        written = pdf_path
    if written != pdf_path:
        raise PrintError("ذخیره", f"مسیر درست توی فیلد نوشته نشد: {written!r}")

    save_btns = [b for b in save_dlg.descendants(control_type="Button")
                 if b.element_info.automation_id == "1"]
    if not save_btns:
        raise PrintError("ذخیره", "دکمه‌ی Save پیدا نشد.")
    _press(save_btns[0], lambda: _is_gone(save_dlg), 10, "ذخیره", "Save")

    def file_ready():
        if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
            return False
        size = os.path.getsize(pdf_path)
        time.sleep(0.5)
        return os.path.getsize(pdf_path) == size  # نوشتن فایل تموم شده

    if not _wait_until(file_ready, file_timeout, interval=0.5):
        raise PrintError("ذخیره", f"فایل PDF تا {file_timeout} ثانیه ساخته نشد: {pdf_path}")


def print_report_to_pdf(hwnd, pdf_path, printer_name=PDF_PRINTER_NAME, log=print):
    """
    کل مسیر: نمایشگر گزارش → Print → انتخاب پرینتر PDF → Print → Save → چک فایل.
    hwnd = پنجره‌ی اصلی Bime.exe. در صورت خطا PrintError با شماره‌ی مرحله.
    """
    app, main = connect_main(hwnd)
    log("  [مرحله 11] منتظر نمایشگر گزارش و زدن Print...")
    dlg = open_print_dialog(app, main)
    log(f"  [مرحله 12] انتخاب پرینتر «{printer_name}»...")
    select_printer(dlg, printer_name)
    log("  [مرحله 13] تأیید پرینت...")
    confirm_print(dlg)
    log("  [ذخیره] منتظر پنجره‌ی Save...")
    save_pdf(app, pdf_path)
    return pdf_path


if __name__ == "__main__":
    import argparse
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="تست پرینت/ذخیره‌ی PDF از نمایشگر گزارش باز فناوران")
    parser.add_argument("--out", required=True, help="مسیر کامل فایل PDF خروجی")
    parser.add_argument("--printer", default=PDF_PRINTER_NAME)
    args = parser.parse_args()

    from run_no_claim_flow import find_bime_window
    hwnd, _ = find_bime_window()
    try:
        print_report_to_pdf(hwnd, args.out, args.printer)
        print(f"\n✅ ذخیره شد: {args.out}")
    except PrintError as e:
        print(f"\n❌ {e}")
