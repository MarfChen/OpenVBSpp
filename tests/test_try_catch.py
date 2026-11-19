"""
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
Try / Catch / Finally 结构化异常（超集扩展）。"""

import contextlib
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import OVPPError        # noqa: E402
from OpenVBSPP.runtime import vbs_run             # noqa: E402


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<try>")
    return buf.getvalue().splitlines()


class TestTryCatch(unittest.TestCase):
    def test_catch_receives_error_object(self):
        code = ("Try\n x = 1 / 0\nCatch e\n"
                ' echo "num=" & e.Number\nEnd Try\necho "after"')
        self.assertEqual(out(code), ["num=11", "after"])

    def test_catch_variable_exposes_description(self):
        code = 'Try\n Err.Raise 53, "src", "自定义说明"\nCatch e\n echo e.Description\nEnd Try'
        self.assertEqual(out(code), ["自定义说明"])

    def test_when_filter_matches(self):
        code = ("Try\n Err.Raise 53\n"
                'Catch e When e.Number = 53\n echo "文件错误"\n'
                'Catch e\n echo "其他"\nEnd Try')
        self.assertEqual(out(code), ["文件错误"])

    def test_when_filter_falls_through_to_wildcard(self):
        code = ("Try\n Err.Raise 13\n"
                'Catch e When e.Number = 53\n echo "文件错误"\n'
                'Catch e\n echo "其他 " & e.Number\nEnd Try')
        self.assertEqual(out(code), ["其他 13"])

    def test_finally_always_runs(self):
        code = 'Try\n echo "t"\nFinally\n echo "f"\nEnd Try'
        self.assertEqual(out(code), ["t", "f"])

    def test_finally_runs_when_catching(self):
        code = ('Try\n Err.Raise 5\nCatch e\n echo "caught"\n'
                'Finally\n echo "cleanup"\nEnd Try')
        self.assertEqual(out(code), ["caught", "cleanup"])

    def test_exit_sub_runs_finally(self):
        code = ("Sub F()\n Try\n  echo \"body\"\n  Exit Sub\n"
                ' Finally\n  echo "cleanup"\n End Try\n echo "never"\nEnd Sub\nF')
        self.assertEqual(out(code), ["body", "cleanup"])

    def test_unmatched_error_propagates(self):
        code = "Try\n Err.Raise 42\nFinally\n echo \"cleanup\"\nEnd Try"
        with self.assertRaises(OVPPError):
            out(code)

    def test_nested_try_reraise(self):
        code = ("Sub Inner()\n Try\n  Err.Raise 5\n Catch e\n"
                '  echo "inner " & e.Number\n  Err.Raise 9\n End Try\nEnd Sub\n'
                'Try\n Inner\nCatch e\n echo "outer " & e.Number\nEnd Try')
        self.assertEqual(out(code), ["inner 5", "outer 9"])

    def test_err_line_points_at_failing_statement(self):
        code = 'Try\n y = 1\n z = 2 / 0\nCatch e\n echo e.Line\nEnd Try'
        self.assertEqual(out(code), ["3"])

    def test_try_overrides_resume_next_inside_block(self):
        code = ('On Error Resume Next\nTry\n x = 1 / 0\nCatch e\n'
                ' echo "caught"\nEnd Try\ny = 1 / 0\necho "still alive"')
        self.assertEqual(out(code), ["caught", "still alive"])

    def test_error_state_after_catch(self):
        code = ('Try\n Err.Raise 7\nCatch e\n echo e.Number\nEnd Try\n'
                "echo Err.Number")
        self.assertEqual(out(code), ["7", "7"])


class TestErrorDictionary(unittest.TestCase):
    def test_raise_without_description_uses_dictionary(self):
        code = "Try\n Err.Raise 53\nCatch e\n echo e.Description\nEnd Try"
        self.assertEqual(out(code), ["File not found"])

    def test_describe_number_falls_back(self):
        from OpenVBSPP.ovpprt import err
        self.assertEqual(err.describe_number(53), "File not found")
        self.assertEqual(err.describe_number(999999), "Error 999999")

    def test_unknown_error_number(self):
        code = 'Try\n Err.Raise 12345\nCatch e\n echo e.Number & " " & e.Description\nEnd Try'
        self.assertEqual(out(code), ["12345 Error 12345"])


class TestLabels(unittest.TestCase):
    def test_on_error_goto_label(self):
        code = ('Sub F()\n On Error GoTo handler\n echo "start"\n'
                ' x = 1 / 0\n echo "never"\n Exit Sub\n'
                'handler:\n echo "handled " & Err.Number\nEnd Sub\nF')
        self.assertEqual(out(code), ["start", "handled 11"])

    def test_goto_label(self):
        code = ('Sub F()\n echo "a"\n GoTo done\n echo "b"\n'
                'done:\n echo "c"\nEnd Sub\nF')
        self.assertEqual(out(code), ["a", "c"])

    def test_undefined_label_is_a_compile_error(self):
        with self.assertRaises(OVPPError):
            out("Sub F()\n On Error GoTo nowhere\nEnd Sub\nF")

    def test_label_is_not_a_variable(self):
        # 标签行不是赋值：done: 后面不能当变量用
        code = 'Sub F()\n GoTo done\ndone:\n echo "ok"\nEnd Sub\nF'
        self.assertEqual(out(code), ["ok"])


class TestTryAsIdentifier(unittest.TestCase):
    """`Try` 是上下文关键字：作为普通变量名必须仍然合法（VBS 超集要求）。"""

    def test_try_can_be_a_variable(self):
        self.assertEqual(out("Try = 5\necho Try"), ["5"])

    def test_catch_and_finally_can_be_variables(self):
        self.assertEqual(out("Catch = 1\nFinally = 2\necho Catch + Finally"), ["3"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
