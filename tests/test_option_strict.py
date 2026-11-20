"""`Option Strict` runtime coercion tests.

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
`Option Strict` 运行期类型转换测试。

运行：
    python -m unittest tests.test_option_strict -v
"""

import contextlib
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import OVPPSyntaxError, OVPPTypeError   # noqa: E402
from OpenVBSPP.runtime import vbs_run                             # noqa: E402


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<strict>")
    return buf.getvalue().splitlines()


def run(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<strict>")
    return buf.getvalue()


class TestStrictOff(unittest.TestCase):
    """Without the option VBScript coercion still applies.

    中文：不写这个 Option 时仍然是 VBS 的强制转换语义。
    """

    def test_text_plus_text_concatenates(self):
        self.assertEqual(out('echo "1" + "2"'), ["12"])

    def test_text_times_number(self):
        self.assertEqual(out('echo "2" * 3'), ["6"])

    def test_text_minus_text(self):
        self.assertEqual(out('echo "6" - "4"'), ["2"])

    def test_options_are_independent(self):
        # Option Compare alone must not switch arithmetic into strict mode.
        # 中文：只写 Option Compare 不会把算术切到严格模式。
        code = 'Option Compare Text\necho "1" + "2"'
        self.assertEqual(out(code), ["12"])


class TestStrictOn(unittest.TestCase):
    """With `Option Strict` a text operand is an error 13.

    中文：打开 `Option Strict` 后，文本操作数报错误 13。
    """

    def test_text_plus_text_is_type_mismatch(self):
        code = 'Option Strict\necho "1" + "2"'
        with self.assertRaises(OVPPTypeError) as ctx:
            out(code)
        self.assertEqual(ctx.exception.number, 13)

    def test_number_plus_text_is_type_mismatch(self):
        code = 'Option Strict\necho 1 + "2"'
        with self.assertRaises(OVPPTypeError):
            out(code)

    def test_text_times_number_is_type_mismatch(self):
        # the silently wrong `"222"` from the coercing path must not come back
        # 中文：强制转换路径里那个静默错误结果 `"222"` 不能再出现
        code = 'Option Strict\necho "2" * 3'
        with self.assertRaises(OVPPTypeError):
            out(code)

    def test_text_divided_by_number_is_type_mismatch(self):
        code = 'Option Strict\necho "6" / "3"'
        with self.assertRaises(OVPPTypeError):
            out(code)

    def test_integer_division_and_mod_reject_text(self):
        # 中文：整除与取模同样拒绝文本
        for expression in ('"6" \\ 2', '"6" Mod 4'):
            code = "Option Strict\necho " + expression
            with self.assertRaises(OVPPTypeError, msg=expression):
                out(code)

    def test_unary_minus_rejects_text(self):
        code = 'Option Strict\necho -"3"'
        with self.assertRaises(OVPPTypeError):
            out(code)

    def test_numbers_still_work(self):
        code = ("Option Strict\n"
                "echo 1 + 2\n"
                "echo 10 / 4\n"
                "echo 5 \\ 2\n"
                "echo 7 Mod 4\n"
                "echo 2 ^ 10")
        self.assertEqual(out(code), ["3", "2.5", "2", "3", "1024"])

    def test_empty_still_counts_as_zero(self):
        code = "Option Strict\necho Empty + 1"
        self.assertEqual(out(code), ["1"])

    def test_null_still_propagates(self):
        code = "Option Strict\necho IsNull(Null + 1)"
        self.assertEqual(out(code), ["True"])

    def test_concatenation_is_still_allowed(self):
        # `&` is an explicit text operation, so it is not a conversion.
        # 中文：`&` 是显式的文本运算，不算隐式转换。
        code = 'Option Strict\necho "a" & 1 & Null & "b"'
        self.assertEqual(out(code), ["a1b"])

    def test_comparison_is_unchanged(self):
        # Comparisons keep their existing behaviour; `Option Strict` only guards
        # arithmetic. 中文：比较行为不变，`Option Strict` 只管算术。
        code = 'Option Strict\necho 1 = 1\necho "a" < "b"'
        self.assertEqual(out(code), ["True", "True"])

    def test_error_13_is_catchable_by_on_error(self):
        code = ('Option Strict\n'
                'On Error Resume Next\n'
                'x = 1 + "2"\n'
                'echo Err.Number')
        self.assertEqual(out(code), ["13"])


class TestStrictAndTyping(unittest.TestCase):
    def test_option_typing_turns_strict_on(self):
        # `Option Typing` implies `Option Strict`.
        # 中文：`Option Typing` 隐含打开 `Option Strict`。
        code = 'Option Typing\necho "1" + "2"'
        with self.assertRaises(OVPPTypeError):
            out(code)

    def test_typed_code_still_runs_arithmetic(self):
        code = ("Option Typing\n"
                "Function Twice(n -> Integer) -> Integer\n"
                "    Twice = n * 2\n"
                "End Function\n"
                "echo Twice(21)")
        self.assertEqual(out(code), ["42"])

    def test_explicit_strict_on(self):
        code = 'Option Strict On\necho "1" + "2"'
        with self.assertRaises(OVPPTypeError):
            out(code)


class TestOptionPlacement(unittest.TestCase):
    def test_option_after_code_is_rejected(self):
        # A late `Option Strict Off` would otherwise rewrite the code above it.
        # 中文：迟到的 `Option Strict Off` 否则会改变它上面代码的编译方式。
        code = 'echo 1\nOption Strict Off\necho 2'
        with self.assertRaises(OVPPSyntaxError):
            out(code)

    def test_option_inside_procedure_is_rejected(self):
        code = 'Sub P()\n    Option Strict\nEnd Sub\nP'
        with self.assertRaises(OVPPSyntaxError):
            out(code)

    def test_options_may_follow_each_other(self):
        code = 'Option Explicit\nOption Strict\nDim a\na = 1 + 1\necho a'
        self.assertEqual(out(code), ["2"])


if __name__ == "__main__":
    unittest.main()
