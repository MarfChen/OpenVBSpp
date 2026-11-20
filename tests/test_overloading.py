"""Procedure/method overloading tests (v0.6.1: dispatch by argument count, resolved at compile time).

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
过程/方法重载测试（v0.6.1：按参数个数分派，编译期解析）。

运行：
    python -m unittest tests.test_overloading -v
"""

import contextlib
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import OVPPArgumentError, OVPPError   # noqa: E402
from OpenVBSPP.runtime import vbs_run                           # noqa: E402


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<overload>")
    return buf.getvalue().splitlines()


class TestFunctionOverloading(unittest.TestCase):
    def test_zero_one_two_arguments(self):
        code = ("Function Greet()\n Greet = \"hi\"\nEnd Function\n"
                "Function Greet(name)\n Greet = \"hi \" & name\nEnd Function\n"
                "Function Greet(greeting, name)\n Greet = greeting & \", \" & name\n"
                "End Function\n"
                "echo Greet()\necho Greet(\"ann\")\necho Greet(\"hello\", \"bob\")")
        self.assertEqual(out(code), ["hi", "hi ann", "hello, bob"])

    def test_sub_overloading(self):
        code = ("Sub Show(a)\n echo \"one:\" & a\nEnd Sub\n"
                "Sub Show(a, b)\n echo \"two:\" & a & b\nEnd Sub\n"
                "Show 1\nShow 1, 2")
        self.assertEqual(out(code), ["one:1", "two:12"])

    def test_wrong_arity_reports_available_overloads(self):
        code = ("Sub Show(a)\n echo a\nEnd Sub\n"
                "Sub Show(a, b)\n echo a & b\nEnd Sub\n"
                "Show 1, 2, 3")
        with self.assertRaises(OVPPArgumentError) as ctx:
            out(code)
        self.assertIn("available: 1, 2", ctx.exception.message)

    def test_same_arity_is_rejected(self):
        code = ("Sub Show(a)\n echo a\nEnd Sub\n"
                "Sub Show(b)\n echo b\nEnd Sub")
        with self.assertRaises(OVPPError):
            out(code)

    def test_overload_keeps_byref(self):
        code = ("Sub Bump(ByRef n)\n n = n + 1\nEnd Sub\n"
                "Sub Bump(ByRef n, ByVal k)\n n = n + k\nEnd Sub\n"
                "x = 1\nBump x\ny = 1\nBump y, 10\necho x & \",\" & y")
        self.assertEqual(out(code), ["2,11"])


class TestMethodOverloading(unittest.TestCase):
    def test_method_overload_by_arity(self):
        code = ("Class Calc\n Public Function Add(a)\n  Add = a\n End Function\n"
                " Public Function Add(a, b)\n  Add = a + b\n End Function\nEnd Class\n"
                "Set c = New Calc\necho c.Add(5)\necho c.Add(5, 6)")
        self.assertEqual(out(code), ["5", "11"])

    def test_method_overload_wrong_arity(self):
        code = ("Class Calc\n Public Function Add(a)\n  Add = a\n End Function\n"
                " Public Function Add(a, b)\n  Add = a + b\n End Function\nEnd Class\n"
                "Set c = New Calc\necho c.Add(1, 2, 3)")
        with self.assertRaises(OVPPArgumentError) as ctx:
            out(code)
        self.assertIn("available: 1, 2", ctx.exception.message)

    def test_overloaded_method_inside_other_method(self):
        code = ("Class Calc\n Public Function Add(a)\n  Add = a\n End Function\n"
                " Public Function Add(a, b)\n  Add = a + b\n End Function\n"
                " Public Function Sum3(a, b, c)\n  Sum3 = Add(Add(a, b), c)\n"
                " End Function\nEnd Class\n"
                "Set c = New Calc\necho c.Sum3(1, 2, 3)")
        self.assertEqual(out(code), ["6"])


class TestTypingWithOverloads(unittest.TestCase):
    def test_argument_types_checked_per_arity(self):
        code = ("Option Typing\n"
                "Function F(a -> Integer) -> Integer\n F = a\nEnd Function\n"
                "Function F(a -> Integer, b -> Integer) -> Integer\n F = a + b\nEnd Function\n"
                "echo F(1)\necho F(1, 2)")
        self.assertEqual(out(code), ["1", "3"])

    def test_wrong_argument_type_rejected(self):
        code = ("Option Typing\n"
                "Function F(a -> Integer) -> Integer\n F = a\nEnd Function\n"
                "Function F(a -> Integer, b -> Integer) -> Integer\n F = a + b\nEnd Function\n"
                'echo F("x")')
        with self.assertRaises(OVPPError):
            out(code)


if __name__ == "__main__":
    unittest.main(verbosity=2)
