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

The VBScript error runtime: the Err object plus the On Error Resume Next state
machine.

中文：VBS 的错误运行时 —— Err 对象 + On Error Resume Next 的状态机。

Implementation notes:

中文：实现要点：
  * one frame per procedure (enter/leave); `On Error Resume Next` affects only the current
    frame,
中文：每个过程一帧（enter/leave），`On Error Resume Next` 只影响当前帧，
    so a handler inside a callee never leaks to its caller;
    中文：所以子过程里的错误处理不会泄漏给调用方；
  * with Resume Next on, the compiler wraps each statement in try/except,
    中文：打开 Resume Next 后，编译器会把每条语句包进 try/except，
    and the except clause calls capture(): it records the error and carries on, which
    is what resuming means;
    中文：except 里调用 capture()，记录错误信息并继续执行（「恢复」的语义）；
  * without a handler, capture() re-raises and runtime.vbs_run turns it into an
    中文：没打开错误处理时 capture() 会把异常重新抛出，由 runtime.vbs_run 翻译成
    OVPP* error.
    中文：OVPP* 错误。
'''

import os

from OpenVBSPP.exceptions import (OVPPError, OVPPRuntimeError)

_PACKAGE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class _Frame:
    __slots__ = ("handler",)

    def __init__(self):
        self.handler = False


_FRAMES = [_Frame()]

# Stack of Try blocks: a non-empty stack means a structured handler is waiting, so
# On Error Resume Next stops swallowing errors (Try wins).
# 中文：Try 块的处理栈 —— 非空表示有结构化异常块在等待，此时 On Error Resume Next
# stop swallowing: Try wins.
# 中文：不再吞异常（Try 优先）。
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

# Error numbers: Number -> human readable description (common VBScript errors).
# 中文：错误号字典 —— Number -> 人类可读的描述（VBS 常用错误号）。
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
    """Turn an error number into a readable description, falling back to
    `Error N`.

    中文：错误号 → 人类可读描述（查不到就返回 `Error N`）。
    """
    try:
        number = int(number)
    except (TypeError, ValueError):
        return str(number)
    return ERRORS.get(number, f"Error {number}")


# Python exception -> VBScript error number.
# 中文：Python 异常 → VBS 错误号。
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
    """The Err object: `Err.Number`, `Err.Description`, `Err.Clear`, `Err.Raise`.

    中文：Err 对象 —— `Err.Number` / `Err.Description` / `Err.Clear` / `Err.Raise`。
    """

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
        # The script-supplied number overrides the class default.
        # 中文：错误号用脚本指定的值覆盖类默认值。
        error.number = int(number)
        raise error

    @property
    def Line(self):
        """The failing line number (a superset extension; VBScript has none).

    中文：出错行号（超集扩展，VBS 不提供）。
    """
        return _STATE["line"]

    def __str__(self):
        return f"Err({self.Number}, {self.Description!r})"

    __repr__ = __str__


ERR = ErrObject()


# --------------------------------------------------------------------- State machine / 中文：状态机
def enter():
    """Enter a procedure: push a fresh error-handling frame (handler off).

    中文：进入一个过程 —— 压入新的错误处理帧（默认关闭）。
    """
    _FRAMES.append(_Frame())


def leave():
    """Leave a procedure: pop the frame.

    中文：离开一个过程 —— 弹出帧。
    """
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
    """Find the user script line inside a traceback.

    Every generated Python statement carries a VBScript line number, so the deepest
    frame outside this package is the real line that failed.

    中文：从 traceback 里找出「用户脚本那一行」。生成的 Python 代码每个语句都带着
    VBScript line number, so the deepest frame that is not inside this package is the real line where the script failed.
    中文：VBS 行号，所以取最深的那个「不在本包内部」的栈帧，就是脚本出错的真实行号。
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
    """Record the error state without raising (used by Try blocks).

    中文：记录错误状态，但不抛出（Try 块用）。
    """
    number, description = _describe(exc)
    _STATE["number"] = number
    _STATE["description"] = description
    _STATE["source"] = "OpenVBS++"
    _STATE["line"] = _guess_line(exc)
    return number, description


def capture(exc):
    """Record the error, re-raising it when nothing takes over.

    Any enclosing Try block counts as a taker, so On Error Resume Next inside a Try
    never swallows the exception: it is handed to the Catch clause.

    中文：记录错误；没有接手者时继续往上抛。任何一层 Try 块都算接手者，所以块内的
    On Error Resume Next does not swallow here: the exception goes to the Try block's Catch.
    中文：On Error Resume Next 不会把异常吞掉，异常会交给 Try 的 Catch。
    """
    record(exc)
    if _HANDLERS or not _FRAMES[-1].handler:
        raise


# --------------------------------------------------------------------- Try blocks / 中文：Try 块
def push_handler():
    """Enter a Try block.

    中文：进入 Try 块。
    """
    _HANDLERS.append(True)


def pop_handler():
    """Leave a Try block (best effort, safe on error paths).

    中文：离开 Try 块（尽力而为，异常路径上也不会崩）。
    """
    if _HANDLERS:
        _HANDLERS.pop()


def handler_active():
    return bool(_HANDLERS)


def set_resume(mode):
    """Record the action taken by `Resume` / `Resume Next` inside a handler.

    中文：处理段里执行到 Resume / Resume Next 时记录动作。
    """
    _RESUME[0] = mode or ""


def take_resume():
    """Take and clear the latest Resume action.

    中文：取出并清空最近的 Resume 动作。
    """
    mode = _RESUME[0]
    _RESUME[0] = ""
    return mode


class _Snapshot:
    """The Catch variable `e`: an immutable snapshot of the error.

    中文：Catch 变量 e —— 错误信息的一份不可变快照。
    """
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
    """A snapshot of the current error, handed to `Catch e`.

    中文：当前错误的一份快照（给 Catch e 的 e 用）。
    """
    return _Snapshot(_STATE)


def reset():
    """Reset the error state and the frame stack (used by tests).

    中文：把错误状态和帧栈恢复到初始情况（测试用）。
    """
    del _FRAMES[1:]
    del _HANDLERS[:]
    _RESUME[0] = ""
    _FRAMES[0].handler = False
    _clear()
