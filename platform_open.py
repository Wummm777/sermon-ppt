# -*- coding: utf-8 -*-
"""跨平台小工具：用系统默认程序打开文件/文件夹（Win/macOS/Linux）。"""
import os
import subprocess
import sys


def open_path(path):
    """用系统默认方式打开文件或文件夹。"""
    if not os.path.exists(path):
        return False
    if sys.platform == "win32":
        try:
            os.startfile(path)
            return True
        except Exception:
            return False
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False


def message_box(text, title="讲道稿转PPT", error=False):
    """跨平台消息弹窗：Windows 用原生 MessageBox；其他平台用 tkinter。"""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, text, title, 0x30 if error else 0x40)
        except Exception:
            print(("[错误] " if error else "") + text)
        return
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        if error:
            messagebox.showerror(title, text)
        else:
            messagebox.showinfo(title, text)
        r.destroy()
    except Exception:
        print(("[错误] " if error else "") + text)
