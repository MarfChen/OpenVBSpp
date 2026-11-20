'''
Copyright 2026 Marf Chen(MarfChen)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

GUI dialogs for `MsgBox` and `InputBox`, backed by Python's `tkinter`.

This file is a bundled over package: scripts never import it. The runtime loads
it from `packages/ui-support/` (a project install, so a version can be pinned) or
from the copy shipped in `OpenVBSPP/vendor/ui-support/`, then
`OpenVBSPP/ovpprt/ui.py` decides when it is used.

The mapping is deliberately thin:

  * each of the six VBScript button sets maps onto one `tk_messageBox -type`;
  * the four icon constants map onto `-icon`;
  * `vbDefaultButtonN` maps onto `-default`;
  * the string `tk_messageBox` returns maps back onto the VBScript codes 1-7.

Function names follow the com-support convention (`is_available`,
`show_message`, `ask_text`) so the dispatcher can treat both packages the same
way.

中文：`MsgBox` 与 `InputBox` 的 GUI 对话框，底层用 Python 的 `tkinter`。

这个文件是一个内置 over 包：脚本里不需要 import。运行期会从
`packages/ui-support/`（项目内安装，可钉版本）或随包分发的
`OpenVBSPP/vendor/ui-support/` 加载它，再由 `OpenVBSPP/ovpprt/ui.py` 决定何时
使用。

映射刻意做得很薄：六种 VBScript 按钮组合各对应一个 `tk_messageBox -type`；
四个图标常量对应 `-icon`；`vbDefaultButtonN` 对应 `-default`；
`tk_messageBox` 返回的字符串再映射回 VBScript 的 1-7。

函数名沿用 com-support 的约定（`is_available`、`show_message`、`ask_text`），
这样分发层可以用同样的方式对待两个包。
'''

import sys

#: Version of this bundled over package.
#: 中文：这个内置 over 包的版本号。
VERSION = "1.0.0"

#: VBScript button sets -> `tk_messageBox -type`.
#: 中文：VBScript 按钮组合 -> `tk_messageBox -type`。
BUTTON_TYPES = {
    0: "ok",                        # vbOKOnly
    1: "okcancel",                  # vbOKCancel
    2: "abortretryignore",          # vbAbortRetryIgnore
    3: "yesnocancel",               # vbYesNoCancel
    4: "yesno",                     # vbYesNo
    5: "retrycancel",               # vbRetryCancel
}

#: VBScript icon bits -> `tk_messageBox -icon`.
#: 中文：VBScript 图标位 -> `tk_messageBox -icon`。
ICON_NAMES = {
    16: "error",                    # vbCritical
    32: "question",                 # vbQuestion
    48: "warning",                  # vbExclamation
    64: "info",                     # vbInformation
}

#: `tk_messageBox` result string -> VBScript return code.
#: 中文：`tk_messageBox` 返回字符串 -> VBScript 返回码。
RESULT_CODES = {
    "ok": 1,                        # vbOK
    "cancel": 2,                    # vbCancel
    "abort": 3,                     # vbAbort
    "retry": 4,                     # vbRetry
    "ignore": 5,                    # vbIgnore
    "yes": 6,                       # vbYes
    "no": 7,                        # vbNo
}

#: The buttons each set shows, in VBScript order (index of `vbDefaultButtonN`).
#: 中文：每种按钮组合显示的按钮（顺序对应 `vbDefaultButtonN` 的序号）。
DEFAULT_BUTTONS = {
    0: ("ok",),
    1: ("ok", "cancel"),
    2: ("abort", "retry", "ignore"),
    3: ("yes", "no", "cancel"),
    4: ("yes", "no"),
    5: ("retry", "cancel"),
}

#: Dialog title used when the script passes an empty one.
#: 中文：脚本传入空标题时使用的默认标题。
DEFAULT_TITLE = "OpenVBS++"

#: The one hidden Tk root shared by every dialog in this process.
#: 中文：本进程所有对话框共用的一个隐藏 Tk 根。
_ROOT = None


def is_available():
    """True when `tkinter` can be imported at all.

    中文：`tkinter` 是否至少能被导入。
    """
    try:
        import tkinter  # noqa: F401
    except Exception:
        return False
    return True


def dialog_spec(buttons=0):
    """Split a VBScript `buttons` value into `tk_messageBox` options.

    Returns a dict with `type`, `icon`, `default` and `topmost`, which is what
    the caller feeds to `_show`. Unknown or reserved bits are ignored, and an
    out-of-range button set falls back to `vbOKOnly` the way VBScript does.

    中文：把 VBScript 的 `buttons` 值拆成 `tk_messageBox` 选项。

    返回含 `type`、`icon`、`default`、`topmost` 的字典，调用方直接喂给 `_show`。
    未知或保留位被忽略；按钮组合越界时按 VBScript 的行为退回 `vbOKOnly`。
    """
    try:
        value = int(buttons or 0)
    except (TypeError, ValueError):
        value = 0

    kind = value & 0x07
    if kind not in BUTTON_TYPES:
        kind = 0

    icon_bits = value & 0x70
    default_bits = value & 0x300
    names = DEFAULT_BUTTONS[kind]

    # vbDefaultButton1 is 0, so the first button is the default unless a
    # higher bit says otherwise; beyond the last button we clamp.
    # 中文：vbDefaultButton1 就是 0，所以默认是第一个按钮；
    # 超出最后一个按钮时向上取最后一个。
    index = min(default_bits >> 8, len(names) - 1)

    return {
        "type": BUTTON_TYPES[kind],
        "icon": ICON_NAMES.get(icon_bits, ""),
        "default": names[index],
        "topmost": bool(value & 0x3000),   # vbApplicationModal / vbSystemModal
    }


def _root():
    """Create (once) and return the hidden Tk root.

    中文：创建（只创建一次）并返回隐藏的 Tk 根。
    """
    global _ROOT
    if _ROOT is None:
        import tkinter
        root = tkinter.Tk()
        root.withdraw()
        _ROOT = root
    return _ROOT


def show_message(message, buttons=0, title=""):
    """Show a modal message box and return the VBScript code 1-7.

    中文：弹出模态消息框，返回 VBScript 的 1-7 返回码。
    """
    from tkinter import messagebox

    spec = dialog_spec(buttons)
    options = {"_type": spec["type"]}
    if spec["icon"]:
        options["_icon"] = spec["icon"]
    if spec["default"]:
        options["default"] = spec["default"]

    _root()
    if spec["topmost"]:
        try:
            _ROOT.attributes("-topmost", True)
        except Exception:
            pass

    result = messagebox._show(title=str(title or ""), message=str(message),
                              **options)
    return RESULT_CODES.get(str(result).strip().lower(), 1)


def ask_text(prompt, title="", default=""):
    """Show a modal single-line prompt and return the typed text.

    Cancel maps to `""` because that is what VBScript's `InputBox` returns; the
    caller never sees `None`.

    中文：弹出模态单行输入框，返回输入的文本。

    取消映射为 `""`，因为 VBScript 的 `InputBox` 就是这样；调用方永远看不到
    `None`。
    """
    from tkinter import simpledialog

    parent = _root()
    answer = simpledialog.askstring(str(title or DEFAULT_TITLE), str(prompt),
                                    initialvalue=str(default or ""),
                                    parent=parent)
    if answer is None:
        return ""
    return str(answer)


def describe():
    """One-line status used by tests and diagnostics.

    中文：给测试与诊断用的一句话状态。
    """
    if not is_available():
        return f"ui-support {VERSION}: tkinter is not available"
    try:
        import tkinter
        version = str(tkinter.TkVersion)
        patch = str(tkinter.TclVersion)
    except Exception:
        version = patch = "?"
    return (f"ui-support {VERSION}: ready "
            f"(Tk {version} / Tcl {patch}, {sys.platform})")
