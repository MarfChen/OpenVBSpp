"""Semantic guards for the v0.9.15 performance work: the optimisations changed generated code, not language semantics.

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

v0.9.15 的性能优化的语义守卫。

优化改的是「怎么生成代码」，不是语言语义，所以这些用例专门钉住边界：

  1. For 快路径（编译成 range）后，循环变量的终值仍符合 VBScript；
  2. 只读的 ByRef 参数按值传，但读写语义与写回仍然正确；
  3. 没有 On Error / Err 的过程不再生成错误帧，出错仍要正常抛出；
  4. datetime / math / random / re / json 改成懒加载后不该在启动时被导入。
"""

import io
import os
import subprocess
import sys
import unittest
from contextlib import redirect_stdout

from OpenVBSPP.runtime import vbs_run


def out(code):
    buffer = io.StringIO()
    with redirect_stdout(buffer):
        vbs_run(code, filename="<test>")
    return [line for line in buffer.getvalue().splitlines()]


class TestForFastPath(unittest.TestCase):
    """Boundary semantics of compiling For into range.

    中文：For 编译成 range 之后的边界语义。"""

    def test_counter_after_normal_loop(self):
        self.assertEqual(out("for i = 1 to 3\nnext\necho i"), ["4"])

    def test_counter_after_negative_step(self):
        self.assertEqual(out("for i = 3 to 1 step -1\nnext\necho i"), ["0"])

    def test_counter_after_skipping_step(self):
        self.assertEqual(out("for i = 0 to 10 step 5\nnext\necho i"), ["15"])

    def test_empty_loop_keeps_start_value(self):
        self.assertEqual(out("i = 99\nfor i = 5 to 1\nnext\necho i"), ["5"])

    def test_empty_loop_with_negative_step(self):
        self.assertEqual(out("i = 99\nfor i = 1 to 5 step -1\nnext\necho i"), ["1"])

    def test_exit_for_does_not_add_step(self):
        code = ("for i = 1 to 10\n"
                "  if i = 3 then exit for\n"
                "next\necho i")
        self.assertEqual(out(code), ["3"])

    def test_body_assigns_counter_falls_back(self):
        # a loop body that assigns the counter must fall back to while or iterations shift
        # 中文：循环体改计数器时必须退回 while 形式，否则会漏掉/多算迭代
        code = ("for i = 1 to 5\n"
                "  echo i\n"
                "  i = i + 1\n"
                "next\necho i")
        self.assertEqual(out(code), ["1", "3", "5", "7"])

    def test_float_step(self):
        self.assertEqual(out("for i = 0 to 0.6 step 0.3\necho CStr(Round(i, 1))\nnext"),
                         ["0", "0.3", "0.6"])

    def test_float_bounds_without_step(self):
        self.assertEqual(out("for i = 1.5 to 3.5\nnext\necho i"), ["4.5"])

    def test_end_expression_evaluated_once(self):
        code = "n = 3\nfor i = 1 to n\nn = 1\necho i\nnext\necho i"
        self.assertEqual(out(code), ["1", "2", "3", "4"])

    def test_nested_loop_reusing_counter(self):
        # the inner loop pushes the counter to 2, the outer +1 lands on 3, so only one line prints
        # 中文：内层循环把计数器推到 2，外层再 +1 就到 3，所以只输出一次
        code = ("for i = 1 to 2\n"
                "  for i = 1 to 1\n"
                "    echo i\n"
                "  next\n"
                "next")
        self.assertEqual(out(code), ["1"])


class TestByRefReadOnly(unittest.TestCase):
    """Read-only ByRef parameters travel by value: readable, not writable by the callee; mutated ones are still written back.

    中文：只读 ByRef 参数按值传：读得到、写不到调用方；被改的仍然写回。"""

    def test_read_only_byref_does_not_change_caller(self):
        code = ("Function Twice(n)\n  Twice = n * 2\nEnd Function\n"
                "x = 5\necho Twice(x)\necho x")
        self.assertEqual(out(code), ["10", "5"])

    def test_mutated_byref_writes_back(self):
        code = ("Sub Bump(n)\n  n = n + 1\nEnd Sub\n"
                "x = 5\nBump x\necho x")
        self.assertEqual(out(code), ["6"])

    def test_byref_passed_down_still_propagates(self):
        code = ("Sub Inner(n)\n  n = n * 10\nEnd Sub\n"
                "Sub Outer(n)\n  Inner n\nEnd Sub\n"
                "x = 4\nOuter x\necho x")
        self.assertEqual(out(code), ["40"])

    def test_swap_still_works(self):
        code = ("Sub Swap(ByRef a, ByRef b)\n"
                "  t = a\n  a = b\n  b = t\nEnd Sub\n"
                "x = 1\ny = 2\nSwap x, y\necho CStr(x) & \",\" & CStr(y)")
        self.assertEqual(out(code), ["2,1"])

    def test_byval_argument_unaffected(self):
        code = ("Function Id(ByVal n)\n  n = n + 1\n  Id = n\nEnd Function\n"
                "x = 7\necho Id(x)\necho x")
        self.assertEqual(out(code), ["8", "7"])

    def test_read_only_byref_through_delegate(self):
        code = ("Function Double(n)\n  Double = n * 2\nEnd Function\n"
                "f = AddressOf Double\necho f(21)")
        self.assertEqual(out(code), ["42"])

    def test_read_only_byref_array(self):
        code = ("Function Total(items)\n"
                "  sum = 0\n"
                "  For Each item In items\n"
                "    sum = sum + item\n"
                "  Next\n"
                "  Total = sum\n"
                "End Function\n"
                "a = Array(1, 2, 3)\necho Total(a)")
        self.assertEqual(out(code), ["6"])


class TestErrorFrameOnDemand(unittest.TestCase):
    """Procedures without On Error / Err no longer build an error frame, yet propagation must not change.

    中文：没有 On Error / Err 的过程不再生成错误帧，但错误传播不能变。"""

    def test_error_propagates_without_frame(self):
        from OpenVBSPP.exceptions import OVPPError

        code = ("Sub Boom()\n  x = 1 / 0\nEnd Sub\nBoom")
        with self.assertRaises(OVPPError) as ctx:
            out(code)
        self.assertIn("Division", str(ctx.exception))

    def test_on_error_resume_next_still_works(self):
        code = ("Sub Attempt()\n"
                "  On Error Resume Next\n"
                "  x = 1 / 0\n"
                "  echo \"after \" & CStr(Err.Number)\n"
                "  On Error GoTo 0\n"
                "End Sub\n"
                "Attempt\necho \"done\"")
        lines = out(code)
        self.assertEqual(lines[-1], "done")
        self.assertTrue(lines[0].startswith("after "))
        self.assertNotEqual(lines[0], "after 0")

    def test_err_object_keeps_frame(self):
        code = ("Sub Report()\n"
                "  On Error Resume Next\n"
                "  x = 1 / 0\n"
                "  echo CStr(Err.Number <> 0)\n"
                "End Sub\n"
                "Report")
        self.assertEqual(out(code), ["True"])

    def test_on_error_goto_still_works(self):
        code = ("Sub Guarded()\n"
                "  On Error GoTo Failed\n"
                "  x = 1 / 0\n"
                "  echo \"unreachable\"\n"
                "  Exit Sub\n"
                "Failed:\n"
                "  echo \"caught\"\n"
                "End Sub\n"
                "Guarded\necho \"done\"")
        self.assertEqual(out(code), ["caught", "done"])


class TestLazyImports(unittest.TestCase):
    """懒加载：内置函数模块不该把 re / json / datetime 拉进启动路径。

    注意：本仓库的 venv 里 site-packages 的启动钩子会提前 import re，
    所以这里用 `-S`（不加载 site）跑子进程，只让源码目录进 sys.path ——
    OpenVBSPP 运行期没有第三方依赖，`-S` 下照样能 import。
    """

    def _modules_after(self, statement, isolated=True):
        code = ("import sys\n"
                f"{statement}\n"
                "print('re' in sys.modules, 'json' in sys.modules)\n")
        env = dict(os.environ)
        env["PYTHONPATH"] = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        command = [sys.executable, "-c", code]
        if isolated:
            command.insert(1, "-S")
        result = subprocess.run(command, capture_output=True, text=True, env=env)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.strip()

    def test_builtins_module_is_lazy(self):
        self.assertEqual(self._modules_after("import OpenVBSPP.ovpprt.builtins"),
                         "False False")

    def test_regexp_module_is_lazy(self):
        self.assertEqual(self._modules_after("import OpenVBSPP.ovpprt.regexp"),
                         "False False")

    def test_jsonobj_module_is_lazy(self):
        self.assertEqual(self._modules_after("import OpenVBSPP.ovpprt.jsonobj"),
                         "False False")

    def test_object_module_is_lazy(self):
        self.assertEqual(self._modules_after("import OpenVBSPP.ovpprt.object"),
                         "False False")

    def test_runtime_does_not_pull_json(self):
        # runtime itself imports re, so only pin that json is not dragged in
        # 中文：runtime 自己要用 re，所以只钉住 json 不被顺带拉起来
        self.assertEqual(self._modules_after("import OpenVBSPP.runtime").split()[1],
                         "False")

    def test_lazy_module_still_provides_attributes(self):
        # the proxy module must really work, not be traded away for speed
        # 中文：代理模块必须真的能用（避免「为了快而坏掉」）
        code = ("import OpenVBSPP.ovpprt.builtins as b\n"
                "print(b.math.floor(1.9), len(b.re.findall('a', 'aa')))\n")
        env = dict(os.environ)
        env["PYTHONPATH"] = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        result = subprocess.run([sys.executable, "-S", "-c", code],
                                capture_output=True, text=True, env=env)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), "1 2")


if __name__ == "__main__":
    unittest.main()
