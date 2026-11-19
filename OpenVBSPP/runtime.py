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

OpenVBS++ 运行时：编译 + 执行，并把 Python 异常翻译成 VBS 风格的错误。
'''


import re

from OpenVBSPP.compiler import vbs_compile, COMPILED_FILENAME
from OpenVBSPP.exceptions import (OVPPError, OVPPRuntimeError, OVPPTypeError,
                                  OVPPVariableNotDefinedError,
                                  OVPPDivisionByZeroError, OVPPOverflowError,
                                  OVPPArgumentError, OVPPObjectNotSupportedError,
                                  OVPPPermissionDeniedError, OVPPFileNotFoundError,
                                  OVPPOutOfRangeError, OVPPInvalidProcedureCallError)


# 顺序有意义：子类要排在父类前面
_ERROR_MAP = [
    (ZeroDivisionError,       OVPPDivisionByZeroError),
    (NameError,               OVPPVariableNotDefinedError),
    (UnboundLocalError,       OVPPVariableNotDefinedError),
    (TypeError,               OVPPTypeError),
    (OverflowError,           OVPPOverflowError),
    (RecursionError,          OVPPOverflowError),
    (IndexError,              OVPPOutOfRangeError),
    (KeyError,                OVPPOutOfRangeError),
    (FileNotFoundError,       OVPPFileNotFoundError),
    (PermissionError,         OVPPPermissionDeniedError),
    (AttributeError,          OVPPObjectNotSupportedError),
    (ValueError,              OVPPInvalidProcedureCallError),
]

_NAME_RE = re.compile(r"name '([^']+)' is not defined")


def _error_location(exc: BaseException, filename: str):
    """找到异常在 .vbs 源码里对应的行号（取最深的一帧）。"""
    line = None
    tb = exc.__traceback__
    while tb is not None:
        if tb.tb_frame.f_code.co_filename == filename:
            line = tb.tb_lineno
        tb = tb.tb_next
    return line


def _error_message(exc: BaseException) -> str:
    if isinstance(exc, (NameError, UnboundLocalError)):
        m = _NAME_RE.search(str(exc))
        if m:
            return f"Variable '{m.group(1)}' is not defined"
    if isinstance(exc, ZeroDivisionError):
        return "Division by zero"
    if isinstance(exc, RecursionError):
        return "Procedure call nesting is too deep (stack overflow)"
    return str(exc) or exc.__class__.__name__


def wrap_runtime_error(exc: BaseException, filename: str, source: str) -> OVPPError:
    """把 Python 运行时异常包装成带行号的 VBS 风格错误。"""
    error_type = OVPPRuntimeError
    for py_type, mapped in _ERROR_MAP:
        if isinstance(exc, py_type):
            error_type = mapped
            break

    line = _error_location(exc, filename)
    source_line = None
    if line is not None:
        lines = source.splitlines()
        if 0 < line <= len(lines):
            source_line = lines[line - 1].rstrip()

    return error_type(
        message=_error_message(exc),
        description=f"Python {type(exc).__name__}: {exc}",
        line_number=line,
        source_line=source_line,
    )


def _bundle_line(code_obj, exc):
    """多文件程序：从 traceback 里取"用户脚本"那一行。"""
    line = None
    tb = exc.__traceback__
    while tb is not None:
        filename = tb.tb_frame.f_code.co_filename
        if filename and not filename.startswith("<frozen"):
            line = tb.tb_lineno
        tb = tb.tb_next
    return line


def vbs_run(code: str, globals: dict = None, filename: str = None,
            mode: str = "exec", save_dir: str = None, verbose: bool = False):
    """编译并执行 OpenVBS++ 源码。

    :returns: 执行所用的 globals 字典
    :raises OVPPError: 语法错误或运行期错误（带行号与源码行）
    """
    code_obj = vbs_compile(code, mode=mode, filename=filename,
                           save_dir=save_dir, verbose=verbose)

    g = {} if globals is None else globals
    compiled_name = str(filename) if filename else COMPILED_FILENAME

    try:
        if hasattr(code_obj, "exec"):
            code_obj.exec(g)
        else:
            exec(code_obj, g)
    except (SystemExit, KeyboardInterrupt):
        raise
    except OVPPError:
        raise
    except Exception as exc:
        error = wrap_runtime_error(exc, compiled_name, code)
        if getattr(error, "line_number", None) is None:
            error.line_number = _bundle_line(code_obj, exc)
        raise error from exc

    return g
