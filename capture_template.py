# -*- coding: utf-8 -*-
"""
capture_template.py
====================
ساخت تصویر یه مرحله (دکمه/آیتم منو) برای screen_locator.py — جایگزین کالیبراسیون
مختصات.

روش کار (مشکل «منو با Alt+Tab بسته می‌شه» رو نداره):
    1) اجرا کن. چند ثانیه (پیش‌فرض ۵) وقت داری برگردی به فناوران و صفحه رو دقیقاً
       به حالتی ببری که اون دکمه دیده می‌شه (مثلاً منوی ▼ رو باز کنی). موس رو روی
       خود دکمه نگه ندار (رنگ hover توی تصویر نیفته)؛ یه جای خالی بذارش.
    2) عکس کل صفحه خودکار گرفته می‌شه و روی یه پنجره‌ی تمام‌صفحه نشون داده می‌شه.
    3) روی همون عکس ثابت، با موس دور دکمه کادر بکش (کلیک-نگه‌دار-بکش-رها کن).
       ترجیحاً کمی از متن/آیکون کنارش رو هم بگیر تا تصویر یکتا باشه.
    4) اگه نقطه‌ی کلیک نباید وسط کادر باشه (مثلاً فقط روی ▼)، یه کلیک روی همون
       نقطه بزن. وگرنه لازم نیست.
    5) Enter = ذخیره، Esc = لغو، کلیک راست = از نو.

اجرا:
    python capture_template.py --label 9
    python capture_template.py --label 10 --delay 8

label همون اسم مرحله در coords.json هست ("9"، "10"، ...). از این به بعد
run_no_claim_flow.py اون مرحله رو با تصویر پیدا می‌کنه، نه مختصات.
"""

import argparse
import ctypes
import json
import os
import sys
import time

# باید دقیقاً مثل run_no_claim_flow.py باشه تا پیکسل‌های عکس = پیکسل‌های واقعی صفحه
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

try:
    import tkinter as tk
    import pyautogui
    from PIL import ImageTk
except ImportError as e:
    print(f"پیش‌نیاز نصب نیست ({e}): pip install pyautogui pillow")
    sys.exit(1)

import screen_locator as sl

INSTRUCTIONS = ("Drag a box around the button   |   optional: click the exact click point inside it   |   "
                "Enter = save   Esc = cancel   Right-click = restart")


class Selector:
    def __init__(self, shot, label):
        self.shot = shot
        self.label = label
        self.box = None        # (x1, y1, x2, y2)
        self.click = None      # (x, y) مطلق روی عکس
        self.start = None
        self.saved = False

        self.root = tk.Tk()
        self.root.attributes("-fullscreen", True)
        self.root.attributes("-topmost", True)
        self.photo = ImageTk.PhotoImage(shot)
        self.canvas = tk.Canvas(self.root, width=shot.width, height=shot.height,
                                highlightthickness=0, cursor="cross")
        self.canvas.pack()
        self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
        self.canvas.create_rectangle(0, 0, shot.width, 28, fill="black", outline="")
        self.status = self.canvas.create_text(10, 14, anchor="w", fill="yellow",
                                              font=("Segoe UI", 11, "bold"),
                                              text=f"[{label}]  {INSTRUCTIONS}")
        self.rect_id = None
        self.mark_id = None

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<ButtonPress-3>", lambda e: self.reset())
        self.root.bind("<Return>", lambda e: self.save())
        self.root.bind("<Escape>", lambda e: self.root.destroy())

    def set_status(self, text):
        self.canvas.itemconfigure(self.status, text=f"[{self.label}]  {text}")

    def reset(self):
        for item in (self.rect_id, self.mark_id):
            if item:
                self.canvas.delete(item)
        self.rect_id = self.mark_id = None
        self.box = self.click = self.start = None
        self.set_status(INSTRUCTIONS)

    def on_press(self, e):
        self.start = (e.x, e.y)

    def on_drag(self, e):
        if self.start is None:
            return
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(*self.start, e.x, e.y, outline="red", width=2)

    def on_release(self, e):
        if self.start is None:
            return
        x1, y1 = self.start
        self.start = None
        if abs(e.x - x1) < 5 and abs(e.y - y1) < 5:
            # کلیک ساده (نه کشیدن) = انتخاب نقطه‌ی کلیک داخل کادر
            if self.box and self.box[0] <= e.x <= self.box[2] and self.box[1] <= e.y <= self.box[3]:
                self.click = (e.x, e.y)
                if self.mark_id:
                    self.canvas.delete(self.mark_id)
                self.mark_id = self.canvas.create_oval(e.x - 5, e.y - 5, e.x + 5, e.y + 5,
                                                       fill="lime", outline="black")
                self.set_status("Click point set.  Enter = save   Right-click = restart")
            if self.rect_id and not self.box:
                self.canvas.delete(self.rect_id)
                self.rect_id = None
            return
        self.box = (min(x1, e.x), min(y1, e.y), max(x1, e.x), max(y1, e.y))
        self.click = None
        if self.mark_id:
            self.canvas.delete(self.mark_id)
            self.mark_id = None
        self.set_status("Box set.  Optional: click the exact click point inside it.  Enter = save")

    def save(self):
        if not self.box:
            self.set_status("Draw a box first!  " + INSTRUCTIONS)
            return
        x1, y1, x2, y2 = self.box
        crop = self.shot.crop(self.box)
        os.makedirs(sl.TEMPLATES_DIR, exist_ok=True)
        crop.save(sl.template_path(self.label))
        cx, cy = self.click if self.click else ((x1 + x2) // 2, (y1 + y2) // 2)
        meta = {"click": [cx - x1, cy - y1], "size": [x2 - x1, y2 - y1],
                "captured_screen": [self.shot.width, self.shot.height]}
        with open(sl.meta_path(self.label), "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        self.saved = True
        self.root.destroy()

    def run(self):
        self.root.mainloop()
        return self.saved


def check_unique(shot, label):
    """بعد از ذخیره: چک می‌کنه تصویر روی همون عکس، جای دیگه‌ای تکرار نشده باشه."""
    if not sl.available():
        print("⚠ opencv-python نصب نیست؛ چک یکتا بودن انجام نشد (pip install opencv-python).")
        return
    tpl = sl._load_gray(sl.template_path(label))
    screen = sl.to_gray(shot)
    score, loc = sl.match(tpl, screen)
    others = sl.count_other_matches(tpl, screen, loc)
    if others:
        print(f"⚠ این تصویر {others} جای دیگه‌ی صفحه هم پیدا می‌شه — ممکنه اشتباهی روی اون‌ها کلیک بشه.")
        print("  دوباره بساز و کادر رو بزرگ‌تر بگیر (مثلاً همراه متن کنار دکمه) تا یکتا بشه.")
    else:
        print(f"✅ تصویر یکتاست (شباهت با خودش: {score:.2f}).")


def main():
    parser = argparse.ArgumentParser(description="ساخت تصویر یه مرحله برای پیدا کردن بدون مختصات")
    parser.add_argument("--label", required=True, help="اسم مرحله، مثل coords.json (مثلاً 9 یا 10)")
    parser.add_argument("--delay", type=int, default=5,
                        help="چند ثانیه صبر قبل از عکس گرفتن (تا صفحه/منو رو آماده کنی)")
    args = parser.parse_args()

    print(f"⏳ {args.delay} ثانیه وقت داری برگردی به فناوران و صفحه رو آماده کنی "
          f"(مثلاً منو رو باز کنی). موس رو روی خود دکمه نذار.")
    for i in range(args.delay, 0, -1):
        print(f"  {i}...", flush=True)
        time.sleep(1)
    shot = pyautogui.screenshot()
    print("📸 عکس گرفته شد. روی عکسی که باز شد دور دکمه کادر بکش؛ Enter = ذخیره، Esc = لغو.")

    if Selector(shot, args.label).run():
        print(f"✅ ذخیره شد: {sl.template_path(args.label)}")
        check_unique(shot, args.label)
    else:
        print("لغو شد.")


if __name__ == "__main__":
    main()
