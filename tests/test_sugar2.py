"""v0.9.1 sugar: string interpolation / compound assignment / ?. / ??.

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
v0.9.1 语法糖：字符串插值 / 复合赋值 / ?. / ??。
"""

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
        vbs_run(code, filename="<sugar2>")
    return buf.getvalue().splitlines()


class TestInterpolation(unittest.TestCase):
    def test_variable(self):
        self.assertEqual(out('name = "World"\necho $"Hello {name}!"'), ["Hello World!"])

    def test_expression(self):
        self.assertEqual(out('x = 10\necho $"x={x} double={x * 2}"'), ["x=10 double=20"])

    def test_plain_text_unchanged(self):
        self.assertEqual(out('echo $"no holes"'), ["no holes"])

    def test_multiple_holes(self):
        self.assertEqual(out('a = 1: b = 2\necho $"{a}-{b}-{a + b}"'), ["1-2-3"])


class TestCompoundAssignment(unittest.TestCase):
    def test_plus(self):
        self.assertEqual(out("a = 1\na += 2\necho a"), ["3"])

    def test_times(self):
        self.assertEqual(out("a = 3\na *= 4\necho a"), ["12"])

    def test_minus(self):
        self.assertEqual(out("a = 10\na -= 4\necho a"), ["6"])

    def test_concat(self):
        self.assertEqual(out('s = "a"\ns &= "b"\necho s'), ["ab"])

    def test_member_target(self):
        code = ("Class C\n Public n\nEnd Class\nSet c = New C\n"
                "c.n = 5\nc.n += 3\necho c.n")
        self.assertEqual(out(code), ["8"])


class TestOptionalChaining(unittest.TestCase):
    def test_reads_when_object_exists(self):
        code = ('Class P\n Public Name\nEnd Class\nSet p = New P\n'
                'p.Name = "ann"\necho p?.Name')
        self.assertEqual(out(code), ["ann"])

    def test_nothing_does_not_raise(self):
        # when n is Nothing, n?.X yields Nothing instead of raising
        # 中文：n 是 Nothing 时 n?.X 得到 Nothing，不抛异常
        self.assertEqual(out("Set n = Nothing\necho TypeName(n?.X)"), ["Nothing"])

    def test_method_call_on_nothing(self):
        code = 'Set n = Nothing\nx = n?.Missing(1)\necho TypeName(x)'
        self.assertEqual(out(code), ["Nothing"])


class TestNullCoalescing(unittest.TestCase):
    def test_empty_uses_right(self):
        self.assertEqual(out('v = Empty\necho v ?? "fb"'), ["fb"])

    def test_value_kept(self):
        self.assertEqual(out('v = 5\necho v ?? "fb"'), ["5"])

    def test_null_uses_right(self):
        self.assertEqual(out('v = Null\necho v ?? "fb"'), ["fb"])


class TestAddressOfNotYet(unittest.TestCase):
    def test_clear_error(self):
        code = "Function F(n)\n F = n\nEnd Function\nAddressOf F"
        with self.assertRaises(OVPPError):
            out(code)


if __name__ == "__main__":
    unittest.main(verbosity=2)
