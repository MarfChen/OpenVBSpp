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

VBS 的错误运行时：Err 对象 + On Error Resume Next 的状态机。

实现要点：
  * 每个过程一帧（enter/leave），`On Error Resume Next` 只影响当前帧，
    所以子过程里的错误处理不会泄漏给调用方；
  * 打开 Resume Next 后，编译器会把每条语句包进 try/except，
    except 里调用 capture()：记录错误信息并继续执行（"恢复"的语义）；
  * 没打开错误处理时 capture() 会把异常重新抛出，由 runtime.vbs_run 翻译成
    OVPP* 错误。
'''

import os

from OpenVBSPP.exceptions import (OVPPError, OVPPRuntimeError)

_PACKAGE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class _Frame:
    __slots__ = ("handler",)

    def __init__(self):
        self.handler = False


_FRAMES = [_Frame()]

# Try 块的处理栈：非空表示"有结构化异常块在等待"，
# 此时 On Error Resume Next 不再吞异常（Try 优先）。
_HANDLERS = []
_RESUME = [""]   # "" / "next" / "retry"

_STATE = {
    "number": 0,
    "description": "",
    "source": "",
    "helpfile": "",
    "helpcontext": 0,
    "line": 0,
}

# 错误号字典：Number -> 人类可读的描述（VBS 常用错误号）
ERRORS = {
    5: "Invalid procedure call or argument",
    6: "Overflow",
    7: "Out of memory",
    9: "Subscript out of range",
    10: "This array is fixed or temporarily locked",
    11: "Division by zero",
    13: "Type mismatch",
    14: "Out of string space",
    17: "Cannot perform requested operation",
    28: "Out of stack space",
    35: "Sub or Function not defined",
    48: "Error in loading DLL",
    51: "Internal error",
    52: "Bad file name or number",
    53: "File not found",
    54: "Bad file mode",
    55: "File already open",
    57: "Device I/O error",
    58: "File already exists",
    61: "Disk full",
    62: "Input past end of file",
    67: "Too many files",
    68: "Device unavailable",
    70: "Permission denied",
    71: "Disk not ready",
    75: "Path/File access error",
    76: "Path not found",
    91: "Object variable or With block variable not set",
    92: "For loop not initialized",
    93: "Invalid pattern string",
    94: "Invalid use of Null",
    322: "Unable to create necessary temporary file",
    424: "Object required",
    429: "ActiveX component can't create object",
    430: "Class does not support Automation",
    432: "File name or class name not found",
    438: "Object doesn't support this property or method",
    440: "Automation error",
    445: "Object doesn't support this action",
    446: "Object doesn't support named arguments",
    447: "Object doesn't support current locale setting",
    448: "Named argument not found",
    449: "Argument not optional",
    450: "Wrong number of arguments or invalid property assignment",
    451: "Object not a collection",
    453: "Specified DLL function not found",
    457: "This key is already associated with an element of this collection",
    1001: "Variable is undefined",
    1002: "Syntax error",
    1003: "Cannot assign to constant",
    1004: "Type mismatch in declaration",
}


def describe_number(number):
    """错误号 → 人类可读描述（查不到就返回 'Error N'）。"""
    try:
        number = int(number)
    except (TypeError, ValueError):
        return str(number)
    return ERRORS.get(number, f"Error {number}")


# Python 异常 → VBS 错误号
_PY_MAP = (
    (ZeroDivisionError, 11, "Division by zero"),
    (NameError, 1001, "Variable is undefined"),
    (UnboundLocalError, 1001, "Variable is undefined"),
    (TypeError, 13, "Type mismatch"),
    (OverflowError, 6, "Overflow"),
    (RecursionError, 6, "Overflow"),
    (IndexError, 9, "Subscript out of range"),
    (KeyError, 9, "Subscript out of range"),
    (FileNotFoundError, 53, "File not found"),
    (PermissionError, 70, "Permission denied"),
    (AttributeError, 438, "Object doesn't support this property or method"),
    (ValueError, 5, "Invalid procedure call"),
)


def _describe(exc):
    if isinstance(exc, OVPPError):
        text = exc.message or exc.name or ""
        if not text or text.startswith("Error "):
            text = describe_number(exc.number)
        return exc.number, text
    for exc_type, number, label in _PY_MAP:
        if isinstance(exc, exc_type):
            text = str(exc)
            return number, f"{label}: {text}" if text else label
    return 5, f"Runtime error: {exc}"


def _clear():
    _STATE["number"] = 0
    _STATE["line"] = 0
    _STATE["description"] = ""
    _STATE["source"] = ""
    _STATE["helpfile"] = ""
    _STATE["helpcontext"] = 0


class ErrObject:
    """Err 对象：`Err.Number` / `Err.Description` / `Err.Clear` / `Err.Raise`。"""

    @property
    def number(self):
        return _STATE["number"]

    @number.setter
    def number(self, value):
        _STATE["number"] = int(value)

    @property
    def description(self):
        return _STATE["description"]

    @description.setter
    def description(self, value):
        _STATE["description"] = "" if value is None else str(value)

    @property
    def source(self):
        return _STATE["source"]

    @source.setter
    def source(self, value):
        _STATE["source"] = "" if value is None else str(value)

    @property
    def helpfile(self):
        return _STATE["helpfile"]

    @property
    def helpcontext(self):
        return _STATE["helpcontext"]

    def clear(self):
        _clear()

    def raise_error(self, number, source="", description="", helpfile="", helpcontext=0):
        _STATE["number"] = int(number)
        _STATE["source"] = "" if source is None else str(source)
        _STATE["description"] = "" if description is None else str(description)
        _STATE["helpfile"] = "" if helpfile is None else str(helpfile)
        _STATE["helpcontext"] = int(helpcontext)

        error = OVPPRuntimeError(
            message=_STATE["description"] or describe_number(number),
            description=f"Raised by script (Err.Raise {int(number)})")
        # 错误号用脚本指定的值覆盖类默认值
        error.number = int(number)
        raise error

    @property
    def Line(self):
        """出错行号（超集扩展，VBS 不提供）。"""
        return _STATE["line"]

    def __str__(self):
        return f"Err({self.Number}, {self.Description!r})"

    __repr__ = __str__


ERR = ErrObject()


# --------------------------------------------------------------------- 状态机
def enter():
    """进入一个过程：压入新的错误处理帧（默认关闭）。"""
    _FRAMES.append(_Frame())


def leave():
    """离开一个过程：弹出帧。"""
    if len(_FRAMES) > 1:
        _FRAMES.pop()


def resume_next():
    """On Error Resume Next"""
    _FRAMES[-1].handler = True


def goto_zero():
    """On Error GoTo 0"""
    _FRAMES[-1].handler = False


def handler_enabled():
    return _FRAMES[-1].handler


def _guess_line(exc):
    """从 traceback 里找出"用户脚本那一行"。

    生成的 Python 代码每个语句都带着 VBS 行号，所以取最深的那个
    "不在本包内部"的栈帧，就是脚本出错的真实行号。
    """
    line = getattr(exc, "line_number", 0) or 0
    if line:
        return line
    candidate = 0
    tb = exc.__traceback__
    while tb is not None:
        filename = tb.tb_frame.f_code.co_filename
        if not filename.startswith(_PACKAGE_DIR) and not filename.startswith("<frozen"):
            candidate = tb.tb_lineno
        tb = tb.tb_next
    return candidate


def record(exc):
    """记录错误状态，但不抛出（Try 块用）。"""
    number, description = _describe(exc)
    _STATE["number"] = number
    _STATE["description"] = description
    _STATE["source"] = "OpenVBS++"
    _STATE["line"] = _guess_line(exc)
    return number, description


def capture(exc):
    """记录错误；没有接手者时继续往上抛。

    任何一层 Try 块都算接手者——所以块内的 On Error Resume Next 不会
    把异常吞掉，异常会交给 Try 的 Catch。
    """
    record(exc)
    if _HANDLERS or not _FRAMES[-1].handler:
        raise


# --------------------------------------------------------------------- Try 块
def push_handler():
    """进入 Try 块。"""
    _HANDLERS.append(True)


def pop_handler():
    """离开 Try 块（尽力而为，异常路径上也不会崩）。"""
    if _HANDLERS:
        _HANDLERS.pop()


def handler_active():
    return bool(_HANDLERS)


def set_resume(mode):
    """处理段里执行到 Resume / Resume Next 时记录动作。"""
    _RESUME[0] = mode or ""


def take_resume():
    """取出并清空最近的 Resume 动作。"""
    mode = _RESUME[0]
    _RESUME[0] = ""
    return mode


class _Snapshot:
    """Catch 变量 e：错误信息的一份不可变快照。"""
    __slots__ = ("number", "description", "source", "line")

    def __init__(self, state):
        self.number = state["number"]
        self.description = state["description"]
        self.source = state["source"]
        self.line = state.get("line", 0)

    Number = property(lambda self: self.number)
    Description = property(lambda self: self.description)
    Source = property(lambda self: self.source)
    Line = property(lambda self: self.line)

    def __repr__(self):
        return f"ErrSnapshot({self.number}, {self.description!r})"


def snapshot():
    """当前错误的一份快照（给 Catch e 的 e 用）。"""
    return _Snapshot(_STATE)


def reset():
    """把错误状态和帧栈恢复到初始情况（测试用）。"""
    del _FRAMES[1:]
    del _HANDLERS[:]
    _RESUME[0] = ""
    _FRAMES[0].handler = False
    _clear()
