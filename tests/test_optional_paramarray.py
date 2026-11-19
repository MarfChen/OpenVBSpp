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
Optional / ParamArray 参数测试（VBS 原生语法）。"""

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
        vbs_run(code, filename="<params>")
    return buf.getvalue().splitlines()


class TestOptional(unittest.TestCase):
    def test_default_used_when_omitted(self):
        code = ('Sub Greet(name, Optional greeting = "hi")\n'
                ' echo greeting & " " & name\nEnd Sub\n'
                'Greet "ann"\nGreet "bob", "yo"')
        self.assertEqual(out(code), ["hi ann", "yo bob"])

    def test_all_params_optional(self):
        code = ("Function F(Optional a = 1, Optional b = 2)\n F = a + b\nEnd Function\n"
                "echo F()\necho F(10)\necho F(10, 20)")
        self.assertEqual(out(code), ["3", "12", "30"])

    def test_optional_byref_writeback(self):
        code = ("Sub Bump(ByRef n, Optional step = 1)\n n = n + step\nEnd Sub\n"
                "x = 10\nBump x\nBump x, 5\necho x")
        self.assertEqual(out(code), ["16"])

    def test_optional_in_class_method(self):
        code = ("Class G\n Public Function Say(name, Optional greeting = \"hi\")\n"
                "  Say = greeting & \" \" & name\n End Function\nEnd Class\n"
                "Set g = New G\necho g.Say(\"ann\")\necho g.Say(\"bob\", \"yo\")")
        self.assertEqual(out(code), ["hi ann", "yo bob"])

    def test_optional_before_required_is_rejected(self):
        with self.assertRaises(OVPPError):
            out("Sub Bad(Optional a = 1, b)\nEnd Sub")

    def test_default_only_on_optional(self):
        with self.assertRaises(OVPPError):
            out("Sub Bad(a = 1)\nEnd Sub")


class TestParamArray(unittest.TestCase):
    def test_sum_of_variadic_args(self):
        code = ("Function Total(ParamArray nums())\n"
                " Dim i, s\n s = 0\n"
                " For i = 0 To UBound(nums)\n  s = s + nums(i)\n Next\n"
                " Total = s\nEnd Function\n"
                "echo Total()\necho Total(1, 2, 3)\necho Total(1, 2, 3, 4)")
        self.assertEqual(out(code), ["0", "6", "10"])

    def test_paramarray_after_required(self):
        code = ("Sub Show(label, ParamArray rest())\n"
                " echo label & \":\" & (UBound(rest) + 1)\nEnd Sub\n"
                "Show \"a\"\nShow \"b\", 1, 2")
        self.assertEqual(out(code), ["a:0", "b:2"])

    def test_paramarray_in_method(self):
        code = ("Class S\n Public Function Sum(ParamArray xs())\n"
                "  Dim i, t\n  t = 0\n"
                "  For i = 0 To UBound(xs)\n   t = t + xs(i)\n  Next\n"
                "  Sum = t\n End Function\nEnd Class\n"
                "Set s = New S\necho s.Sum(4, 5, 6)")
        self.assertEqual(out(code), ["15"])

    def test_paramarray_must_be_last(self):
        with self.assertRaises(OVPPError):
            out("Sub Bad(ParamArray xs(), b)\nEnd Sub")


if __name__ == "__main__":
    unittest.main(verbosity=2)
