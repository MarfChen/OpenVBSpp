#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""OpenVBS++ 基准测试 / Benchmark.

两件事：
1. **进程内**：同一份逻辑分别用 VBS 源码（编译一次、反复执行）和等价的 Python 源码跑，
   看「编译到 CPython 字节码」本身的额外开销。这才是语言实现的开销。
2. **端到端**（`--e2e`）：连同解释器启动一起计时，和原生 `cscript.exe` 对照。
   这是用户实际感受到的数字。

用法：
    python tools/bench.py              # 只跑进程内
    python tools/bench.py --e2e        # 加跑端到端（Windows 上会自动找 cscript.exe）
    python tools/bench.py --repeat 7   # 每个用例取几次的中位数（默认 5）

注意：这是**单机开发基准**，不是严谨的基准框架。数字用于发现量级问题（比如 O(n²)），
不要拿它做跨机器的精确比较。
"""

from __future__ import annotations

import argparse
import os
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.runtime import vbs_compile  # noqa: E402


# --------------------------------------------------------------------------
# 用例：每项包含 VBS 源码与「应该等价」的 Python 源码
# --------------------------------------------------------------------------

CASES = [
    (
        "1M 算术循环（过程内）",
        """
Option Explicit
Sub Work()
    Dim i, s
    s = 0
    For i = 1 To 1000000
        s = s + i
    Next
End Sub
Work
""",
        """
def work():
    s = 0
    for i in range(1, 1000001):
        s = s + i
work()
""",
    ),
    (
        "1M 算术循环（模块级）",
        """
Option Explicit
Dim i, s
s = 0
For i = 1 To 1000000
    s = s + i
Next
""",
        """
s = 0
for i in range(1, 1000001):
    s = s + i
""",
    ),
    (
        "fib(24) 递归",
        """
Option Explicit
Function Fib(n)
    If n < 2 Then
        Fib = n
    Else
        Fib = Fib(n - 1) + Fib(n - 2)
    End If
End Function
Dim r
r = Fib(24)
""",
        """
def fib(n):
    if n < 2:
        return n
    return fib(n - 1) + fib(n - 2)
r = fib(24)
""",
    ),
    (
        "20k 字符串拼接",
        """
Option Explicit
Dim i, s
s = ""
For i = 1 To 20000
    s = s & "x"
Next
""",
        """
s = ""
for i in range(1, 20001):
    s = s + "x"
""",
    ),
    (
        "100k 数组读写",
        """
Option Explicit
Dim i
Dim a(100000)
For i = 0 To 99999
    a(i) = i
Next
Dim t
t = 0
For i = 0 To 99999
    t = t + a(i)
Next
""",
        """
a = [0] * 100001
for i in range(100000):
    a[i] = i
t = 0
for i in range(100000):
    t = t + a[i]
""",
    ),
    (
        "50k Dictionary 插入",
        """
Option Explicit
Dim i, d
Set d = CreateObject("Scripting.Dictionary")
For i = 1 To 50000
    d.Add i, i
Next
""",
        """
d = {}
for i in range(1, 50001):
    d[i] = i
""",
    ),
    (
        "200k 函数调用",
        """
Option Explicit
Function Double(x)
    Double = x * 2
End Function
Dim i, t
t = 0
For i = 1 To 200000
    t = Double(i)
Next
""",
        """
def double(x):
    return x * 2
t = 0
for i in range(1, 200001):
    t = double(i)
""",
    ),
]

E2E_CASES = [
    ("只 echo 一行（启动开销）", "echo 1\n"),
    (
        "1M 算术循环",
        "Option Explicit\nSub Work()\n  Dim i, s\n  s = 0\n  For i = 1 To 1000000\n"
        "    s = s + i\n  Next\nEnd Sub\nWork\n",
    ),
    (
        "fib(24) 递归",
        "Option Explicit\nFunction Fib(n)\n  If n < 2 Then\n    Fib = n\n  Else\n"
        "    Fib = Fib(n - 1) + Fib(n - 2)\n  End If\nEnd Function\nDim r\nr = Fib(24)\n",
    ),
]


def median_ms(fn, repeat: int) -> float:
    samples = []
    for _ in range(repeat):
        start = time.perf_counter()
        fn()
        samples.append((time.perf_counter() - start) * 1000.0)
    return statistics.median(samples)


def compile_vbs(source: str):
    """把 VBS 源码编译成 Python code object（不含执行）。"""
    compiled = vbs_compile(source)
    if isinstance(compiled, types.CodeType):
        return compiled
    if isinstance(compiled, str):
        return compile(compiled, "<bench>", "exec")
    code = getattr(compiled, "code", None)
    if code is not None and not isinstance(code, str):
        return code
    if isinstance(compiled, dict):
        for value in compiled.values():
            if isinstance(value, str):
                return compile(value, "<bench>", "exec")
    raise SystemExit("无法从 vbs_compile 的返回值里取出 code object：%r" % (type(compiled),))


def run_in_process(repeat: int) -> list:
    rows = []
    for name, vbs_src, py_src in CASES:
        compile_ms = median_ms(lambda src=vbs_src: compile_vbs(src), repeat)
        vbs_code = compile_vbs(vbs_src)

        def run_vbs(code=vbs_code):
            exec(code, {"__name__": "__main__"})

        py_code = compile(py_src, "<bench>", "exec")

        def run_py(code=py_code):
            exec(code, {"__name__": "__main__"})

        vbs_ms = median_ms(run_vbs, repeat)
        py_ms = median_ms(run_py, repeat)
        rows.append((name, compile_ms, vbs_ms, py_ms, vbs_ms / py_ms if py_ms else 0.0))
    return rows


def _median_cmd_ms(cmd: list, repeat: int, env=None) -> float:
    def once():
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)

    return median_ms(once, repeat)


def run_e2e(repeat: int) -> list:
    rows = []
    cscript = shutil.which("cscript") or shutil.which("cscript.exe")
    cwd = Path(tempfile.mkdtemp(prefix="ovpp-bench-"))
    for name, source in E2E_CASES:
        script = cwd / ("case_%d.vbs" % len(rows))
        script.write_text(source, encoding="utf-8")
        ovpp_ms = _median_cmd_ms([sys.executable, "-m", "OpenVBSPP", str(script)], repeat)
        cs_ms = _median_cmd_ms([cscript, "//nologo", str(script)], repeat) if cscript else None
        rows.append((name, ovpp_ms, cs_ms))
    return rows


def print_inproc(rows) -> None:
    print("\n## 进程内（编译一次，反复执行；中位数）\n")
    print("| 用例 | 编译 | OpenVBS++ | CPython | 倍数 |")
    print("|---|---:|---:|---:|---:|")
    for name, compile_ms, vbs_ms, py_ms, ratio in rows:
        print("| %s | %.1f ms | %.1f ms | %.1f ms | **%.2f×** |"
              % (name, compile_ms, vbs_ms, py_ms, ratio))
    worst = max(rows, key=lambda r: r[4])
    print("\n最差的是「%s」（%.2f× CPython）。" % (worst[0], worst[4]))


def print_e2e(rows) -> None:
    has_cscript = any(r[2] for r in rows)
    print("\n## 端到端（含进程启动；中位数）\n")
    if has_cscript:
        print("| 用例 | OpenVBS++ | cscript.exe | CPython |")
        print("|---|---:|---:|---:|")
    else:
        print("| 用例 | OpenVBS++ | cscript.exe |")
        print("|---|---:|---:|")
    for name, ovpp_ms, cs_ms in rows:
        if cs_ms:
            print("| %s | %.0f ms | %.0f ms | - |" % (name, ovpp_ms, cs_ms))
        else:
            print("| %s | %.0f ms | 未安装 |" % (name, ovpp_ms))


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="OpenVBS++ 基准测试")
    parser.add_argument("--e2e", action="store_true", help="额外跑端到端（子进程）基准")
    parser.add_argument("--repeat", type=int, default=5, help="每个用例取样次数，取中位数（默认 5）")
    args = parser.parse_args(argv)

    print("OpenVBS++ 基准 / 平台 %s / Python %s" % (sys.platform, sys.version.split()[0]))
    print_inproc(run_in_process(max(1, args.repeat)))
    if args.e2e:
        print_e2e(run_e2e(max(1, args.repeat)))
    print("\n说明：For 在“步长是常量、循环体不改计数器”时编译成 range，已和 Python 的 for 同速；"
          "剩下的差距来自 VBS 的 & 要走类型转换、函数调用要多一层参数处理，"
          "以及过程内（locals）与模块级（globals）的查找差异。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
