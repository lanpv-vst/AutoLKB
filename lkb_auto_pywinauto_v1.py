#!/usr/bin/env python3
"""
Tabmis automation with a simple Tkinter form using pywinauto for keyboard actions.

Replaces pyautogui with pywinauto.keyboard.send_keys (Windows-only).
- Enter start row and end row (1-based). Defaults: 2 to 2.
- Enter key delay (seconds) between key actions. Default: 0.25
- Choose data file (CSV or Excel).
- Option: Wait while mouse cursor is hourglass (Windows only).
- Press OK to start, Exit to quit.

Requires (Windows):
    pip install pywinauto pyperclip pandas
Note: pywinauto works on Windows. This script keeps clipboard paste via pyperclip
and sends Ctrl+V through pywinauto.send_keys.

Copyright (c) lanpv@vst.gov.vn
"""
import csv
import math
import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import platform
import ctypes

import pandas as pd
import pyperclip

# Try import pywinauto keyboard send_keys and Application
try:
    from pywinauto import keyboard
    from pywinauto import Application
except Exception:
    keyboard = None  # handled at runtime
    Application = None  # handled at runtime

# ---------- Automation functions ----------
class TabmisAutomator:
    def __init__(self, csv_path, start_row, end_row, key_delay,
                 between_rows_delay=0.6, start_delay=3.0, wait_cursor=False):
        self.csv_path = csv_path
        self.start_row = start_row
        self.end_row = end_row
        self.key_delay = float(key_delay)
        self.between_rows_delay = float(between_rows_delay)
        self.start_delay = float(start_delay)
        self.wait_cursor = bool(wait_cursor)
        self._stop_requested = False

    def stop(self):
        self._stop_requested = True

    def focus_tabmis_window(self, status_callback=None):
        """
        Tìm và focus vào cửa sổ Tabmis có title "Các ứng dụng Oracle - Môi trường sản xuất TABMIS 2018".
        Trả về True nếu thành công, False nếu không tìm thấy.
        """
        if Application is None:
            if status_callback:
                status_callback("pywinauto.Application not available. Cannot focus window.")
            return False

        window_title = "Các ứng dụng Oracle - Môi trường sản xuất TABMIS 2018"
        
        try:
            # Cách 1: Thử connect theo title chính xác
            try:
                app = Application(backend="win32").connect(title=window_title, timeout=2)
                window = app.window(title=window_title)
                if window.exists():
                    window.set_focus()
                    window.restore()  # Đảm bảo cửa sổ không bị minimize
                    if status_callback:
                        status_callback(f"Đã focus vào cửa sổ Tabmis")
                    self._sleep_with_cancel(0.5)  # Đợi một chút để cửa sổ focus xong
                    return True
            except Exception:
                pass
            
            # Cách 2: Thử tìm bằng Desktop (tìm tất cả cửa sổ)
            try:
                from pywinauto import Desktop
                desktop = Desktop(backend="win32")
                windows = desktop.windows()
                for w in windows:
                    if w.is_visible() and window_title in w.window_text():
                        w.set_focus()
                        w.restore()
                        if status_callback:
                            status_callback(f"Đã focus vào cửa sổ Tabmis")
                        self._sleep_with_cancel(0.5)
                        return True
            except Exception:
                pass
            
            # Cách 3: Tìm bằng từ khóa TABMIS và Oracle (fuzzy match)
            try:
                from pywinauto import Desktop
                desktop = Desktop(backend="win32")
                windows = desktop.windows()
                for w in windows:
                    text = w.window_text()
                    if w.is_visible() and "TABMIS" in text and ("Oracle" in text or "Môi trường" in text):
                        w.set_focus()
                        w.restore()
                        if status_callback:
                            status_callback(f"Đã focus vào cửa sổ Tabmis (tìm kiếm mờ)")
                        self._sleep_with_cancel(0.5)
                        return True
            except Exception:
                pass
            
            if status_callback:
                status_callback(f"Không tìm thấy cửa sổ Tabmis. Vui lòng mở cửa sổ trước.")
            return False
            
        except Exception as e:
            if status_callback:
                status_callback(f"Lỗi khi tìm cửa sổ Tabmis: {e}")
            return False

    def _sleep_with_cancel(self, total_seconds):
        """Ngủ nhưng vẫn kiểm tra cờ dừng để có thể dừng gần như ngay lập tức."""
        end_time = time.time() + float(total_seconds)
        while time.time() < end_time:
            if self._stop_requested:
                return
            # ngủ từng đoạn rất ngắn để phản ứng nhanh
            time.sleep(0.02)

    def is_cursor_busy(self):
        """
        Kiểm tra con trỏ chuột có đang là Hourglass / AppStarting hay không.
        Hiện thực Windows-only (sử dụng Win32 API). Trả về True nếu đang là cursor 'wait' hoặc 'appstarting'.
        Trên hệ khác trả về False.
        """
        if not self.wait_cursor:
            return False
        if platform.system() != "Windows":
            # Không hỗ trợ trên non-Windows trong phiên bản này
            return False
        try:
            user32 = ctypes.windll.user32

            class POINT(ctypes.Structure):
                _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

            class CURSORINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", ctypes.c_uint),
                    ("flags", ctypes.c_uint),
                    ("hCursor", ctypes.c_void_p),
                    ("ptScreenPos", POINT),
                ]

            ci = CURSORINFO()
            ci.cbSize = ctypes.sizeof(CURSORINFO)
            res = user32.GetCursorInfo(ctypes.byref(ci))
            if not res:
                return False
            cur = ci.hCursor

            # IDC_WAIT = 32514, IDC_APPSTARTING = 32650
            IDC_WAIT = 32514
            IDC_APPSTARTING = 32650
            try:
                wait_h = user32.LoadCursorW(0, IDC_WAIT)
                appstart_h = user32.LoadCursorW(0, IDC_APPSTARTING)
            except Exception:
                # Nếu LoadCursorW không nhận integer (hiếm), fallback False
                wait_h = None
                appstart_h = None

            # So sánh handle; nếu trùng với một trong các handle hệ thống, coi là busy
            if cur and (cur == wait_h or cur == appstart_h):
                return True
            return False
        except Exception:
            # Nếu có lỗi, không block (tránh treo)
            return False

    def wait_while_cursor_busy(self, status_callback=None):
        """
        Nếu tùy chọn chờ được bật thì lặp kiểm tra con trỏ cho tới khi không còn busy.
        Gọi status_callback(text) để cập nhật UI khi cần.
        """
        if not self.wait_cursor:
            return
        # Nếu không phải Windows, thông báo 1 lần (không spam)
        if platform.system() != "Windows":
            if status_callback:
                status_callback("Wait-cursor tính năng chỉ hỗ trợ Windows — bỏ qua.")
            return
        # Lặp cho đến khi cursor không busy hoặc người dùng dừng
        waited = 0.0
        while not self._stop_requested and self.is_cursor_busy():
            if status_callback:
                status_callback("Đang chờ con trỏ chuột hết busy...")
            self._sleep_with_cancel(0.1)
            waited += 0.1
            # Sau một khoảng thời gian dài, cập nhật status để người dùng biết
            if status_callback and (int(waited) % 5 == 0):
                status_callback(f"Đang chờ con trỏ... {int(waited)}s")
        if status_callback and not self._stop_requested:
            status_callback("Con trỏ đã sẵn sàng, tiếp tục...")

    def paste_text(self, text, status_callback=None):
        if self._stop_requested:
            return
        if keyboard is None:
            if status_callback:
                status_callback("pywinauto not available. Install via: pip install pywinauto")
            return

        # Nếu bật tùy chọn chờ con trỏ, đợi trước khi thao tác clipboard/dán
        self.wait_while_cursor_busy(status_callback=status_callback)

        if text is None:
            text = ""
        pyperclip.copy(str(text))
        self._sleep_with_cancel(min(0.1, self.key_delay / 4))
        if self._stop_requested:
            return
        try:
            # Ctrl+V
            keyboard.send_keys("^v")
        except Exception:
            # fallback: send characters one by one
            try:
                s = str(text)
                for ch in s:
                    if self._stop_requested:
                        return
                    keyboard.send_keys(ch)
                    self._sleep_with_cancel(0.005)
            except Exception:
                pass
        self._sleep_with_cancel(self.key_delay)

    def _normalize_key_name(self, key):
        """Chuẩn hóa tên phím đầu vào"""
        key = str(key).lower().strip()
        key_map = {
            'pagedown': 'pagedown',
            'page down': 'pagedown',
            'pgdn': 'pagedown',
            'pageup': 'pageup',
            'page up': 'pageup',
            'pgup': 'pageup',
            'ctrl': 'ctrl',
            'control': 'ctrl',
            'alt': 'alt',
            'shift': 'shift',
            'win': 'win',
            'command': 'win',
            'cmd': 'win',
            'enter': 'enter',
            'esc': 'esc',
            'tab': 'tab',
            'down': 'down',
            'up': 'up',
            'left': 'left',
            'right': 'right',
            'space': 'space',
            'f4': 'f4',
        }
        return key_map.get(key, key)

    def _token_for_key(self, key):
        """
        Trả về token phù hợp cho pywinauto.keyboard.send_keys
        - letters/digits returned as-is
        - special keys mapped to {KEY}
        - page keys and function keys provided with braces
        """
        k = key.lower()
        special = {
            'enter': '{ENTER}',
            'esc': '{ESC}',
            'tab': '{TAB}',
            'down': '{DOWN}',
            'up': '{UP}',
            'left': '{LEFT}',
            'right': '{RIGHT}',
            'pagedown': '{PGDN}',
            'pageup': '{PGUP}',
            'f1': '{F1}',
            'f2': '{F2}',
            'f3': '{F3}',
            'f4': '{F4}',
            'f5': '{F5}',
            'f6': '{F6}',
            'f7': '{F7}',
            'f8': '{F8}',
            'f9': '{F9}',
            'f10': '{F10}',
            'f11': '{F11}',
            'f12': '{F12}',
            'space': ' ',
        }
        return special.get(k, k)

    def press(self, key, count=1):
        if keyboard is None:
            return
        token = self._token_for_key(self._normalize_key_name(key))
        for _ in range(count):
            if self._stop_requested:
                return
            if self.wait_cursor:
                self.wait_while_cursor_busy()
            try:
                keyboard.send_keys(token)
            except Exception:
                # fallback: try raw
                try:
                    keyboard.send_keys(str(token))
                except Exception:
                    pass
            self._sleep_with_cancel(self.key_delay)

    def hotkey(self, *keys):
        """
        Nhấn tổ hợp phím sử dụng pywinauto send_keys:
        - Ctrl -> '^', Shift -> '+', Alt -> '!' prefix.
        Ví dụ:
            hotkey('ctrl','s') -> '^s'
            hotkey('shift','pagedown') -> '+{PGDN}'
            hotkey('alt','c') -> '%c'
        """
        if keyboard is None:
            return
        if self._stop_requested:
            return

        self.wait_while_cursor_busy()

        normalized = [self._normalize_key_name(k) for k in keys]
        MODIFIERS = {'shift', 'ctrl', 'alt', 'win', 'command'}

        modifiers = [k for k in normalized if k in MODIFIERS]
        main_keys = [k for k in normalized if k not in MODIFIERS]

        prefix_map = {
            'ctrl': '^',
            'shift': '+',
            'alt': '%',
            # 'win' is not supported as prefix in send_keys; use {LWIN} if needed
        }
        prefix = ''.join(prefix_map.get(m, '') for m in modifiers)

        if not main_keys:
            # No main key — try to send modifiers alone (rare)
            try:
                # send a modifier press-release by sending prefix alone (may be ignored)
                keyboard.send_keys(prefix)
            except Exception:
                pass
            self._sleep_with_cancel(self.key_delay)
            return

        for mk in main_keys:
            if self._stop_requested:
                return
            token = self._token_for_key(mk)
            seq = f"{prefix}{token}"
            try:
                keyboard.send_keys(seq)
            except Exception:
                # fallback: try without braces
                try:
                    keyboard.send_keys(f"{prefix}{mk}")
                except Exception:
                    pass
            self._sleep_with_cancel(0.02)

        self._sleep_with_cancel(self.key_delay)

    def get_cell(self, row, col_1based):
        idx = col_1based - 1
        if idx < 0:
            return ""
        if idx < len(row):
            return row[idx]
        return ""

    def process_row(self, row):
        # Follow the user's specified sequence exactly
        if self._stop_requested:
            return

        self.paste_text(self.get_cell(row, 1))
        self.press('down')

        self.paste_text(self.get_cell(row, 2))
        self.press('tab')
        self.press('tab')

        self.paste_text(self.get_cell(row, 3))
        for _ in range(5):
            self.press('tab')

        self.paste_text(self.get_cell(row, 4))
        self.press('tab')

        self.paste_text(self.get_cell(row, 5))
        self.press('enter')

        self.press('tab')
        self.press('tab')

        self.paste_text(self.get_cell(row, 6))
        self.press('tab')

        self.paste_text(self.get_cell(row, 7))
        self.press('tab')

        self.paste_text(self.get_cell(row, 8))
        self.press('down')

        self.paste_text(self.get_cell(row, 9))
        self.press('tab')
        self.press('tab')

        self.paste_text(self.get_cell(row, 10))

        self.hotkey('ctrl', 's')
        self.press('enter')
        self.hotkey('alt', 'c')

        #self.press('down', count=4)
        self.press('down')
        self.press('down')
        self.press('down')
        self.press('down')
        self.press('enter')

        self.paste_text(self.get_cell(row, 11))
        self.press('tab')
        self.press('tab')
        self.press('tab')

        self.paste_text(self.get_cell(row, 12))
        self.press('tab')
        self.press('tab')

        self.paste_text(self.get_cell(row, 13))
        self.press('tab')
        self.press('tab')

        # Paste col15 then col16 immediately (as specified)
        self.paste_text(self.get_cell(row, 14))
        self.paste_text(self.get_cell(row, 15))
        self.press('tab')

        self.paste_text(self.get_cell(row, 16))
        self.hotkey('shift', 'pagedown')
        self.hotkey('shift', 'pagedown')
        self.press('tab')

        self.paste_text(self.get_cell(row, 17))
        self.press('tab')
        self.press('tab')
        self.press('tab')

        self.paste_text(self.get_cell(row, 18))
        self.hotkey('ctrl', 's')
        self.press('f4')
        self.hotkey('shift', 'pageup')
        self.press('down')

    def run(self, status_callback=None):
        # status_callback(text) to update UI
        if keyboard is None:
            if status_callback:
                status_callback("pywinauto not installed. Please run: pip install pywinauto")
            return
        try:
            ext = os.path.splitext(self.csv_path)[1].lower()
            if ext == ".csv":
                with open(self.csv_path, newline='', encoding='utf-8') as f:
                    reader = list(csv.reader(f))
            elif ext in (".xlsx", ".xls"):
                # Read all cells (no header) and convert NaN to empty string
                df = pd.read_excel(self.csv_path, header=None)
                df = df.where(df.notna(), "")
                reader = df.astype(str).values.tolist()
            else:
                if status_callback:
                    status_callback(f"Unsupported file type: {ext}. Please use CSV or Excel.")
                return
        except FileNotFoundError:
            if status_callback:
                status_callback(f"File not found: {self.csv_path}")
            return
        except Exception as e:
            if status_callback:
                status_callback(f"Error reading file: {e}")
            return

        total_rows = len(reader)
        if status_callback:
            status_callback(f"Loaded {total_rows} rows. Đang tìm cửa sổ Tabmis...")

        # Tự động focus vào cửa sổ Tabmis
        if not self.focus_tabmis_window(status_callback):
            if status_callback:
                status_callback("Không tìm thấy cửa sổ Tabmis. Vui lòng mở cửa sổ trước khi chạy.")
            # Vẫn tiếp tục, có thể người dùng sẽ focus thủ công trong thời gian countdown
        else:
            if status_callback:
                status_callback(f"Đã focus vào cửa sổ Tabmis. Bắt đầu sau {self.start_delay} giây...")

        # Give user time to focus Tabmis window (nếu chưa focus được tự động)
        countdown = int(self.start_delay)
        for t in range(countdown, 0, -1):
            if self._stop_requested:
                if status_callback:
                    status_callback("Stopped before start.")
                return
            if status_callback:
                status_callback(f"Starting in {t}...")
            self._sleep_with_cancel(1)

        for i in range(self.start_row, self.end_row + 1):
            if self._stop_requested:
                if status_callback:
                    status_callback("Stopped by user.")
                return

            idx = i - 1
            if idx < 0 or idx >= total_rows:
                if status_callback:
                    status_callback(f"Skipping row {i}: not in CSV")
                continue

            row = reader[idx]
            if status_callback:
                status_callback(f"Processing row {i}...")
            try:
                self.process_row(row)
            except Exception as e:
                if status_callback:
                    status_callback(f"Error on row {i}: {e}")
                return

            if status_callback:
                status_callback(f"Finished row {i}. Waiting {self.between_rows_delay}s")
            self._sleep_with_cancel(self.between_rows_delay)

        if status_callback:
            status_callback("All done.")


# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        root.title("LKB Auto")
        root.resizable(False, False)

        # --- Color theme ---
        self.primary_color = "#95031B"   # main background color
        self.button_color = "#C51F3A"    # button background color
        self.button_active = "#7A0113"   # button active color
        self.text_color = "#FFFFFF"      # main text color

        root.configure(bg=self.primary_color)

        frm = tk.Frame(root, padx=10, pady=10, bg=self.primary_color)
        frm.pack()

        tk.Label(frm, text="File dữ liệu:", fg=self.text_color, bg=self.primary_color).grid(row=0, column=0, sticky="e")
        self.csv_var = tk.StringVar(value="")
        self.csv_entry = tk.Entry(frm, width=40, textvariable=self.csv_var, bg="white", fg="black")
        self.csv_entry.grid(row=0, column=1, columnspan=2, sticky="w")
        tk.Button(
            frm,
            text="Browse",
            command=self.browse_csv,
            bg=self.button_color,
            fg=self.text_color,
            activebackground=self.button_active,
            activeforeground=self.text_color
        ).grid(row=0, column=3, padx=(6,0))

        tk.Label(frm, text="Start_r", fg=self.text_color, bg=self.primary_color).grid(row=1, column=0, sticky="e")
        self.start_var = tk.StringVar(value="2")
        tk.Entry(frm, textvariable=self.start_var, width=10, bg="white", fg="black").grid(row=1, column=1, sticky="w")

        tk.Label(frm, text="End_r", fg=self.text_color, bg=self.primary_color).grid(row=1, column=1, sticky="e")
        self.end_var = tk.StringVar(value="2")
        tk.Entry(frm, textvariable=self.end_var, width=10, bg="white", fg="black").grid(row=1, column=2, sticky="w")

        tk.Label(frm, text="Delay_k(s)", fg=self.text_color, bg=self.primary_color).grid(row=2, column=0, sticky="e")
        self.delay_var = tk.StringVar(value="0.25")
        tk.Entry(frm, textvariable=self.delay_var, width=10, bg="white", fg="black").grid(row=2, column=1, sticky="w")

        tk.Label(frm, text="Delay_r(s)", fg=self.text_color, bg=self.primary_color).grid(row=2, column=1, sticky="e")
        self.between_var = tk.StringVar(value="0.25")
        tk.Entry(frm, textvariable=self.between_var, width=10, bg="white", fg="black").grid(row=2, column=2, sticky="w")

        # Checkbox: Wait while mouse cursor is hourglass
        self.wait_cursor_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            frm,
            text="Chờ Tabmis phản hồi",
            variable=self.wait_cursor_var,
            fg=self.text_color,
            bg=self.primary_color,
            activebackground=self.primary_color,
            selectcolor=self.primary_color
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(6,0))

        btn_frame = tk.Frame(frm, pady=8, bg=self.primary_color)
        btn_frame.grid(row=4, column=0, columnspan=4)

        self.ok_btn = tk.Button(
            btn_frame,
            text="Chạy",
            width=12,
            command=self.on_ok,
            bg=self.button_color,
            fg=self.text_color,
            activebackground=self.button_active,
            activeforeground=self.text_color
        )
        self.ok_btn.pack(side="left", padx=6)

        self.exit_btn = tk.Button(
            btn_frame,
            text="Thoát",
            width=12,
            command=self.on_exit,
            bg=self.button_color,
            fg=self.text_color,
            activebackground=self.button_active,
            activeforeground=self.text_color
        )
        self.exit_btn.pack(side="left", padx=6)

        self.stop_btn = tk.Button(
            btn_frame,
            text="Dừng",
            width=12,
            command=self.on_stop,
            state="disabled",
            bg=self.button_color,
            fg=self.text_color,
            activebackground=self.button_active,
            activeforeground=self.text_color
        )
        self.stop_btn.pack(side="left", padx=6)

        self.status_label = tk.Label(
            frm,
            text="Ready. © lanpv@vst.gov.vn",
            anchor="w",
            justify="left",
            fg=self.text_color,
            bg=self.primary_color
        )
        self.status_label.grid(row=5, column=0, columnspan=4, sticky="w")

        # Thêm ngôi sao vàng 5 cánh (cờ Việt Nam) ở góc trên bên phải
        self._add_vietnam_flag_star(root)

        self.automator = None
        self.worker_thread = None

        # Phím tắt: nhấn ESC để dừng ngay lập tức
        self.root.bind("<Escape>", self._on_esc_pressed)

    def _draw_star_5_points(self, canvas, center_x, center_y, outer_radius, inner_radius):
        """Vẽ ngôi sao 5 cánh"""
        points = []
        for i in range(10):  # 10 điểm (5 cánh, mỗi cánh có 2 điểm)
            angle = (i * math.pi / 5) - math.pi / 2  # Bắt đầu từ trên cùng
            if i % 2 == 0:  # Điểm ngoài
                radius = outer_radius
            else:  # Điểm trong
                radius = inner_radius
            x = center_x + radius * math.cos(angle)
            y = center_y + radius * math.sin(angle)
            points.extend([x, y])
        return canvas.create_polygon(points, fill="#FFCD00", outline="#FFCD00", width=1)

    def _add_vietnam_flag_star(self, root):
        """Thêm ngôi sao vàng 5 cánh (cờ Việt Nam) ở góc trên bên phải với viền vàng"""
        # Tạo Canvas nhỏ ở góc trên bên phải (vị trí không che nút Browse)
        canvas_width = 70
        canvas_height = 45
        star_canvas = tk.Canvas(
            root,
            width=canvas_width,
            height=canvas_height,
            bg=self.primary_color,
            highlightthickness=0,
            borderwidth=0
        )
        # Đặt ở góc trên bên phải, cách lề 5px từ trên và 5px từ phải, tránh che nút Browse
        star_canvas.place(relx=1.0, rely=0.0, anchor="ne", x=-75, y=55)
        
        # Vẽ hình chữ nhật viền vàng
        border_width = 2
        star_canvas.create_rectangle(
            border_width, border_width,
            canvas_width - border_width, canvas_height - border_width,
            outline="#FFCD00", width=border_width
        )
        
        # Vẽ ngôi sao vàng 5 cánh ở giữa canvas
        center_x, center_y = canvas_width / 2, canvas_height / 2
        outer_radius = 18  # Bán kính ngoài
        inner_radius = 7   # Bán kính trong
        self._draw_star_5_points(star_canvas, center_x, center_y, outer_radius, inner_radius)

    def _on_esc_pressed(self, event=None):
        """Handler khi nhấn phím ESC – dừng automation ngay."""
        self.on_stop()

    def browse_csv(self):
        path = filedialog.askopenfilename(
            title="Chọn file dữ liệu",
            filetypes=[
                ("Data files (CSV, Excel)", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.csv_var.set(path)

    def set_status(self, text):
        # update status label in main thread
        def _update():
            self.status_label.config(text=text)
        self.root.after(0, _update)

    def on_ok(self):
        try:
            start_row = int(self.start_var.get())
            end_row = int(self.end_var.get())
            if start_row <= 0 or end_row <= 0:
                raise ValueError("Row numbers must be >= 1")
            if end_row < start_row:
                raise ValueError("End row must be >= start row")
            key_delay = float(self.delay_var.get())
            between = float(self.between_var.get())
        except Exception as e:
            messagebox.showerror("Invalid input", f"Please check inputs:\n{e}")
            return

        csv_path = self.csv_var.get().strip()
        if not csv_path:
            messagebox.showerror("File dữ liệu", "Vui lòng chọn file dữ liệu (CSV hoặc Excel).")
            return

        # If user enabled wait-cursor on non-Windows, warn once
        if self.wait_cursor_var.get() and platform.system() != "Windows":
            if not messagebox.askyesno("Chú ý", "Tùy chọn 'Wait while mouse cursor is hourglass' chỉ hỗ trợ Windows. Tiếp tục và bỏ qua tính năng trên hệ thống này?"):
                return

        # If pywinauto not available, warn and abort
        if keyboard is None:
            messagebox.showerror("pywinauto not available", "Module 'pywinauto' not found. Please install it via:\n\npip install pywinauto\n\nThis script requires pywinauto (Windows).")
            return

        # Disable buttons and start thread
        self.ok_btn.config(state="disabled")
        self.exit_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.set_status("Preparing...")

        self.automator = TabmisAutomator(
            csv_path, start_row, end_row, key_delay,
            between_rows_delay=between, start_delay=3.0,
            wait_cursor=self.wait_cursor_var.get()
        )
        self.worker_thread = threading.Thread(target=self._run_worker, daemon=True)
        self.worker_thread.start()

    def _run_worker(self):
        try:
            self.automator.run(status_callback=self.set_status)
        finally:
            # Re-enable buttons when done
            def _finish():
                self.ok_btn.config(state="normal")
                self.exit_btn.config(state="normal")
                self.stop_btn.config(state="disabled")
                messagebox.showinfo("Done", "Automation finished (or stopped).")
            self.root.after(0, _finish)

    def on_stop(self):
        if self.automator:
            self.automator.stop()
            self.set_status("Stop requested...")

    def on_exit(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("Exit", "Automation is running. Exit anyway?"):
                return
        self.root.quit()


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()