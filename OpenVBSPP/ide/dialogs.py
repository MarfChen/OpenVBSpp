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

系统原生"打开 / 另存为"对话框（Windows / Linux / macOS）。

浏览器出于安全限制**不能**替网页弹出系统的保存对话框，所以由本地服务端
调用 Tk 的 filedialog 完成。因为 IDE 只监听 127.0.0.1，弹窗出现在用户自己
的桌面上，体验和本地程序一致。

  * tkinter 是标准库，但不是所有 Linux 发行版都装了 python3-tk；
  * 无图形环境（服务器/容器）时 `available()` 返回 False，前端自动退回浏览器方案。
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
    except Exception:            # ImportError / TclError（无显示环境）
        return None, None
    return tkinter, filedialog


def available():
    """本机能不能弹原生对话框（有 tkinter 且有图形环境）。"""
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
    """弹出"打开"对话框，返回绝对路径；取消或失败返回 None。"""
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
    """弹出"另存为"对话框，返回绝对路径；取消或失败返回 None。"""
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


# --------------------------------------------------------------------- 读写
def read_script(path):
    """读脚本：自动去掉 BOM，换行统一成 \\n（词法器两种都能吃，统一更省心）。"""
    text = Path(path).read_text(encoding="utf-8-sig")
    return text.replace("\r\n", "\n").replace("\r", "\n")


def write_script(path, source):
    """写脚本：固定 UTF-8 + LF，保证 Windows/Linux 下内容一致。"""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(source)
    return str(target)


def platform_hint():
    """给前端/诊断用的平台信息。"""
    return {
        "platform": sys.platform,
        "native_dialogs": available(),
        "newline": "LF (normalized on read/write)",
    }
