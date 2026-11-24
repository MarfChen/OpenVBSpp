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

The IDE backend service layer. It depends on no web framework; every frontend (the
built-in page, Gradio, anything else) calls into it.

Design constraints:
  * standard library only — the core package pulls in no third-party dependencies;
  * every call returns a JSON-serialisable dict so any frontend can reuse it;
  * file access is confined to one workspace directory (default ~/.openvbspp/scripts).

中文：IDE 的后端服务层 —— 不依赖任何 Web 框架，前端（内置网页 / Gradio / 其它）
everything goes through here. Constraints: standard library only, no third-party dependency for the core package, every entry point returns a JSON-serialisable
中文：都调这里。约束：只用标准库，核心包不引入第三方依赖；所有接口返回可 JSON 序列化的
dict, and file access is confined to one workspace directory (default `~/.openvbspp/scripts`).
中文：dict；文件读写被限制在一个工作目录里（默认 `~/.openvbspp/scripts`）。
'''

import io
import time
import contextlib
import traceback
from pathlib import Path

from OpenVBSPP import __version__
from OpenVBSPP.compiler import vbs_compile
from OpenVBSPP.exceptions import OVPPError
from OpenVBSPP.runtime import vbs_run
from OpenVBSPP.ide.packages import capabilities as package_capabilities
from OpenVBSPP.source import read_source


DEFAULT_WORKSPACE = Path.home() / ".openvbspp" / "scripts"


# --------------------------------------------------------------- result objects
# --------------------------------------------------------------- result objects
    # 中文：结果对象
def _error_payload(exc):
    """Turn an OVPPError into a structure a frontend can render directly.

    中文：把 OVPPError 变成前端能直接渲染的结构。
    """
    payload = {
        "kind": type(exc).__name__,
        "number": getattr(exc, "error_number", None),
        "message": getattr(exc, "message", str(exc)),
        "line": getattr(exc, "line_number", None),
        "column": getattr(exc, "column", None),
        "source_line": getattr(exc, "source_line", None),
    }
    return payload


def _collect_output(source, filename):
    """Execute some source, capturing stdout and exceptions.

    中文：执行一段源码，捕获 stdout 与异常。
    """
    buffer = io.StringIO()
    started = time.perf_counter()
    error = None
    ok = True
    try:
        with contextlib.redirect_stdout(buffer):
            vbs_run(source, filename=filename)
    except OVPPError as exc:
        ok = False
        error = _error_payload(exc)
    # Last resort: this should not happen, but it has to be visible.
    # 中文：兜底 —— 不该发生，但要能看到。
    except Exception as exc:
        ok = False
        error = {
            "kind": type(exc).__name__,
            "number": None,
            "message": str(exc),
            "line": None,
            "column": None,
            "source_line": None,
            "traceback": traceback.format_exc(limit=3),
        }
    elapsed = (time.perf_counter() - started) * 1000.0
    return {
        "ok": ok,
        "output": buffer.getvalue(),
        "error": error,
        "duration_ms": round(elapsed, 2),
    }


# ------------------------------------------------------------- compile / run
# ------------------------------------------------------------- compile / run
    # 中文：编译 / 运行
def run_source(source, filename="<ide>"):
    """Run VBS source and return output, errors and elapsed time.

    中文：运行一段 VBS 源码，返回输出 + 错误 + 耗时。
    """
    if not source.strip():
        return {"ok": True, "output": "", "error": None, "duration_ms": 0.0}
    return _collect_output(source, filename)


def check_source(source, filename="<ide>"):
    """Compile only: collect syntax and type errors (Option Typing checks happen here too).

    中文：只编译不运行：拿语法/类型错误（Option Typing 的检查也在这里）。
    """
    started = time.perf_counter()
    try:
        vbs_compile(source, filename=filename)
    except OVPPError as exc:
        return {
            "ok": False,
            "diagnostics": [_error_payload(exc)],
            "duration_ms": round((time.perf_counter() - started) * 1000.0, 2),
        }
    return {
        "ok": True,
        "diagnostics": [],
        "duration_ms": round((time.perf_counter() - started) * 1000.0, 2),
    }


# ------------------------------------------------------------------- files
# ------------------------------------------------------------------- files
    # 中文：文件
class Workspace:
    """A confined workspace: only .vbs files inside this directory can be read or written.

    中文：受限的工作目录：只能读写这个目录里的 .vbs 文件。
    """

    def __init__(self, root=None):
        self.root = Path(root) if root else DEFAULT_WORKSPACE
        self.root.mkdir(parents=True, exist_ok=True)

    def _resolve(self, name):
        candidate = (self.root / f"{name}.vbs").resolve()
        if self.root.resolve() not in candidate.parents and candidate.parent != self.root.resolve():
            raise ValueError("path escapes the workspace")
        return candidate

    def list(self):
        return sorted(path.stem for path in self.root.glob("*.vbs"))

    def load(self, name):
        path = self._resolve(name)
        if not path.is_file():
            return {"ok": False, "message": f"script '{name}' not found"}
        return {"ok": True, "name": name,
                "source": read_source(path),
                "path": str(path)}

    def save(self, name, source):
        path = self._resolve(name)
        path.write_text(source, encoding="utf-8")
        return {"ok": True, "name": name, "path": str(path)}


# ---------------------------------------------------------------- examples
# ---------------------------------------------------------------- examples
    # 中文：示例
DEMOS = {
    "hello": ('echo "Hello, VBScript!"\n'
              'echo "running at " & Now()\n'),
    "fizzbuzz": (
        "For i = 1 To 15\n"
        "    If i Mod 15 = 0 Then\n"
        '        echo "FizzBuzz"\n'
        "    ElseIf i Mod 3 = 0 Then\n"
        '        echo "Fizz"\n'
        "    ElseIf i Mod 5 = 0 Then\n"
        '        echo "Buzz"\n'
        "    Else\n"
        "        echo i\n"
        "    End If\n"
        "Next\n"),
    "class": (
        "Class Counter\n"
        "    Private n\n"
        "    Public Property Get Value()\n"
        "        Value = n\n"
        "    End Property\n"
        "    Public Sub Bump(Optional step = 1)\n"
        "        n = n + step\n"
        "    End Sub\n"
        "    Private Sub Class_Initialize()\n"
        "        n = 0\n"
        "    End Sub\n"
        "End Class\n"
        "\n"
        "Set c = New Counter\n"
        "c.Bump 5\n"
        "c.Bump\n"
        'echo "counter=" & c.Value\n'),
    "typing": (
        "Option Typing\n"
        "Dim total -> Integer\n"
        "total = 0\n"
        "Function Add(a -> Integer, b -> Integer) -> Integer\n"
        "    Add = a + b\n"
        "End Function\n"
        "total = Add(20, 22)\n"
        "echo total\n"
        "\n"
        "rem Deliberate type error: uncomment the next line to see the compile-time error\n"
        "rem 中文：故意写错类型 —— 取消注释下面这行看编译期报错\n"
        "rem total = \"not a number\"\n"),
}


def examples():
    return [{"name": name, "source": source} for name, source in DEMOS.items()]


# ------------------------------------------------------------ capabilities
# ------------------------------------------------------------ capabilities
    # 中文：能力声明
def _dialog_info():
    try:
        from OpenVBSPP.ide import dialogs
        return dialogs.platform_hint()
    except Exception:
        return {"platform": "unknown", "native_dialogs": False}



def capabilities():
    """Declare to frontends and plugins what the IDE supports (package fields are reserved).

    中文：向前端/插件声明本 IDE 支持什么（包管理器相关字段预留给以后）。
    """
    dialog_info = _dialog_info()
    return {
        "project": "OpenVBS++",
        "version": __version__,
        "features": {
            "run": True,
            "check": True,
            "save": True,
            "load": True,
            "examples": True,
            "native_dialogs": dialog_info.get("native_dialogs", False),
            # Reserved: the package manager.
            # 中文：预留 —— 包管理器。
            "packages": False,
            # Reserved: a language server.
            # 中文：预留 —— 语言服务器。
            "lsp": False,
        },
        "dialogs": dialog_info,
        "packages": package_capabilities(),
    }
