# AutoLKB
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
