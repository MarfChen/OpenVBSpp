'''
Copyright 2027 Marf Chen(MarfChen)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Native "open / save as" dialogs on Windows, Linux and macOS.

Browsers cannot pop up the system save dialog on a page's behalf, so the local server
calls Tk's filedialog instead. The IDE only listens on 127.0.0.1, so the dialog appears
on the user's own desktop and feels like a local program.

  * tkinter ships with Python, but some Linux distributions omit python3-tk;
  * without a graphical environment (server, container) `available()` returns False and
    the frontend falls back to the browser flow.

中文：系统原生「打开 / 另存为」对话框（Windows / Linux / macOS）。浏览器不能替网页
A browser cannot raise a system save dialog, so the local server calls Tk's filedialog instead. The IDE binds
中文：弹出系统保存对话框，所以由本地服务端调用 Tk 的 filedialog 完成；IDE 只监听
127.0.0.1 only, so the dialog appears on the user's own desktop. tkinter ships with Python, yet some Linux distributions omit
中文：127.0.0.1，弹窗出现在用户自己的桌面上。注意 tkinter 是标准库但部分 Linux 发行版没装
python3-tk; without a display `available()` returns False and the front end falls back to the browser.
中文：python3-tk；无图形环境时 `available()` 返回 False，前端自动退回浏览器方案。
'''

import sys
from pathlib import Path


FILETYPES = [
    ("VBScript", "*.vbs"),
    ("OpenVBS++", "*.ovbs"),
    ("Text", "*.txt"),
    ("All files", "*.*"),
]


def _import_tk():
    try:
        import tkinter
        from tkinter import filedialog
    # ImportError or TclError, which is what a headless environment raises.
    # 中文：ImportError / TclError（无显示环境）。
    except Exception:
        return None, None
    return tkinter, filedialog


def available():
    """Whether this machine can pop native dialogs (tkinter present and a display available).

    中文：本机能不能弹原生对话框（有 tkinter 且有图形环境）。
    """
    tkinter, _ = _import_tk()
    if tkinter is None:
        return False
    try:
        root = tkinter.Tk()
    except Exception:
        return False
    try:
        root.destroy()
    except Exception:
        pass
    return True


def _hidden_root(tkinter):
    root = tkinter.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    return root


def ask_open(initial_dir=None, title="Open VBScript"):
    """Show the open dialog and return an absolute path, or None on cancel or failure.

    中文：弹出"打开"对话框，返回绝对路径；取消或失败返回 None。
    """
    tkinter, filedialog = _import_tk()
    if tkinter is None:
        return None
    root = None
    try:
        root = _hidden_root(tkinter)
        options = {"title": title, "filetypes": FILETYPES}
        if initial_dir:
            options["initialdir"] = str(initial_dir)
        path = filedialog.askopenfilename(**options)
        return path or None
    except Exception:
        return None
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


def ask_save(default_name="untitled.vbs", initial_dir=None,
             title="Save VBScript"):
    """Show the save-as dialog and return an absolute path, or None on cancel or failure.

    中文：弹出"另存为"对话框，返回绝对路径；取消或失败返回 None。
    """
    tkinter, filedialog = _import_tk()
    if tkinter is None:
        return None
    root = None
    try:
        root = _hidden_root(tkinter)
        options = {"title": title, "defaultextension": ".vbs",
                   "filetypes": FILETYPES}
        if default_name:
            options["initialfile"] = default_name
        if initial_dir:
            options["initialdir"] = str(initial_dir)
        path = filedialog.asksaveasfilename(**options)
        return path or None
    except Exception:
        return None
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


# ------------------------------------------------------------------- read / write
def read_script(path):
    """Read a script: strip a BOM and normalise newlines to \\n.

    The lexer accepts both, and normalising keeps everything simpler.

    中文：读脚本时自动去掉 BOM，换行统一成 \\n（词法器两种都能吃，统一更省心）。"""
    text = Path(path).read_text(encoding="utf-8-sig")
    return text.replace("\r\n", "\n").replace("\r", "\n")


def write_script(path, source):
    """Write a script as UTF-8 with LF so Windows and Linux agree byte for byte.

    中文：写脚本：固定 UTF-8 + LF，保证 Windows/Linux 下内容一致。
    """
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(source)
    return str(target)


def platform_hint():
    """Platform information for frontends and diagnostics.

    中文：给前端/诊断用的平台信息。
    """
    return {
        "platform": sys.platform,
        "native_dialogs": available(),
        "newline": "LF (normalized on read/write)",
    }
