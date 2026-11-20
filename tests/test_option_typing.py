"""`Option Typing` compile-time type checking tests.

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
`Option Typing` 编译期静态类型检查测试。

运行：
    python -m unittest tests.test_option_typing -v
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
        vbs_run(code, filename="<typing>")
    return buf.getvalue().splitlines()


class TestTypingOff(unittest.TestCase):
    def test_annotations_are_ignored_without_option(self):
        # without Option Typing annotations are not checked (equivalent to plain VBS)
        # 中文：没有 Option Typing 时注解不参与检查（等价于普通 VBS）
        code = 'Dim n -> String\nn = 42\necho n'
        self.assertEqual(out(code), ["42"])


class TestTypingOn(unittest.TestCase):
    def test_declared_types_ok(self):
        code = ("Option Typing\nDim n -> Integer\nn = 42\necho n")
        self.assertEqual(out(code), ["42"])

    def test_numeric_widening(self):
        code = "Option Typing\nDim d -> Double\nd = 3\necho d"
        self.assertEqual(out(code), ["3"])

    def test_variant_accepts_everything(self):
        code = 'Option Typing\nDim v\nv = "s"\nv = 1\necho v'
        self.assertEqual(out(code), ["1"])

    def test_integer_to_string_is_rejected(self):
        with self.assertRaises(OVPPTypeError) as ctx:
            out("Option Typing\nDim s -> String\ns = 42")
        self.assertIn("String", str(ctx.exception))

    def test_string_to_integer_is_rejected(self):
        with self.assertRaises(OVPPTypeError):
            out('Option Typing\nDim n -> Integer\nn = "abc"')

    def test_builtin_return_type_is_checked(self):
        # Len returns Integer, so it cannot be assigned to a String
        # 中文：Len 返回 Integer，不能赋给 String
        with self.assertRaises(OVPPTypeError):
            out('Option Typing\nDim s -> String\ns = Len("abc")')

    def test_builtin_return_type_ok(self):
        code = 'Option Typing\nDim n -> Integer\nn = Len("abc")\necho n'
        self.assertEqual(out(code), ["3"])

    def test_argument_type_checked(self):
        code = ("Option Typing\nFunction Twice(n -> Integer) -> Integer\n"
                " Twice = n * 2\nEnd Function\necho Twice(\"x\")")
        with self.assertRaises(OVPPTypeError) as ctx:
            out(code)
        self.assertIn("argument 1", str(ctx.exception))

    def test_argument_type_ok(self):
        code = ("Option Typing\nFunction Twice(n -> Integer) -> Integer\n"
                " Twice = n * 2\nEnd Function\necho Twice(4)")
        self.assertEqual(out(code), ["8"])

    def test_return_type_checked(self):
        code = ('Option Typing\nFunction Bad() -> Integer\n Bad = "str"\nEnd Function')
        with self.assertRaises(OVPPTypeError):
            out(code)

    def test_return_type_ok(self):
        code = ("Option Typing\nFunction Add(a -> Integer, b -> Integer) -> Integer\n"
                " Add = a + b\nEnd Function\necho Add(1, 2)")
        self.assertEqual(out(code), ["3"])

    def test_class_field_type_checked(self):
        code = ("Option Typing\nClass Foo\n Public x -> Integer\nEnd Class\n"
                "Dim f -> Foo\nSet f = New Foo\nf.x = \"str\"")
        with self.assertRaises(OVPPTypeError):
            out(code)

    def test_class_field_type_ok(self):
        code = ("Option Typing\nClass Foo\n Public x -> Integer\nEnd Class\n"
                "Dim f -> Foo\nSet f = New Foo\nf.x = 5\necho f.x")
        self.assertEqual(out(code), ["5"])

    def test_property_type_is_inferred(self):
        code = ("Option Typing\nClass Foo\n Private n\n"
                " Public Property Get Value() -> Integer\n  Value = n\n End Property\n"
                " Public Sub Set_(v -> Integer)\n  n = v\n End Sub\nEnd Class\n"
                "Dim f -> Foo\nSet f = New Foo\nf.Set_ 3\nDim n -> Integer\n"
                "n = f.Value\necho n")
        self.assertEqual(out(code), ["3"])

    def test_array_element_type_checked(self):
        code = ('Option Typing\nDim a(2) -> Integer\na(0) = "x"')
        with self.assertRaises(OVPPTypeError) as ctx:
            out(code)
        self.assertIn("array element", str(ctx.exception))

    def test_array_element_type_ok(self):
        code = ("Option Typing\nDim a(2) -> Integer\na(0) = 5\n"
                "Dim n -> Integer\nn = a(0)\necho n")
        self.assertEqual(out(code), ["5"])

    def test_untyped_array_is_lenient(self):
        code = 'Option Typing\nDim a(2)\na(0) = "x"\necho a(0)'
        self.assertEqual(out(code), ["x"])

    def test_option_typing_inside_procedure_rejected(self):
        with self.assertRaises(OVPPSyntaxError):
            out("Sub P()\n Option Typing\nEnd Sub")


if __name__ == "__main__":
    unittest.main(verbosity=2)
