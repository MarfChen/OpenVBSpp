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

IDE 的后端服务层：不依赖任何 Web 框架，前端（内置网页 / Gradio / 其它）都调这里。

设计约束：
  * 只用标准库，核心包不引入第三方依赖；
  * 所有接口返回可 JSON 序列化的 dict，方便任何前端复用；
  * 文件读写被限制在一个工作目录里（默认 ~/.openvbspp/scripts）。
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


DEFAULT_WORKSPACE = Path.home() / ".openvbspp" / "scripts"


# --------------------------------------------------------------------- 结果对象
def _error_payload(exc):
    """把 OVPPError 变成前端能直接渲染的结构。"""
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
    """执行一段源码，捕获 stdout 与异常。"""
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
    except Exception as exc:                      # 兜底：不该发生，但要能看到
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


# --------------------------------------------------------------------- 编译 / 运行
def run_source(source, filename="<ide>"):
    """运行一段 VBS 源码，返回输出 + 错误 + 耗时。"""
    if not source.strip():
        return {"ok": True, "output": "", "error": None, "duration_ms": 0.0}
    return _collect_output(source, filename)


def check_source(source, filename="<ide>"):
    """只编译不运行：拿语法/类型错误（Option Typing 的检查也在这里）。"""
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


# --------------------------------------------------------------------- 文件
class Workspace:
    """受限的工作目录：只能读写这个目录里的 .vbs 文件。"""

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
                "source": path.read_text(encoding="utf-8-sig"),
                "path": str(path)}

    def save(self, name, source):
        path = self._resolve(name)
        path.write_text(source, encoding="utf-8")
        return {"ok": True, "name": name, "path": str(path)}


# --------------------------------------------------------------------- 示例
DEMOS = {
    "hello": ('echo "Hello, VBScript!"\n'
              'echo "运行在 " & Now()\n'),
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
        "rem 故意写错类型：把下面这行取消注释看编译期报错\n"
        "rem total = \"not a number\"\n"),
}


def examples():
    return [{"name": name, "source": source} for name, source in DEMOS.items()]


# --------------------------------------------------------------------- 能力声明
def _dialog_info():
    try:
        from OpenVBSPP.ide import dialogs
        return dialogs.platform_hint()
    except Exception:
        return {"platform": "unknown", "native_dialogs": False}



def capabilities():
    """向前端/插件声明本 IDE 支持什么（包管理器相关字段预留给以后）。"""
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
            "packages": False,        # 预留：v0.9 的包管理器
            "lsp": False,             # 预留：v1.0 的语言服务器
        },
        "dialogs": dialog_info,
        "packages": package_capabilities(),
    }
