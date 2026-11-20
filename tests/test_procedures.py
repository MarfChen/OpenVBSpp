"""Regression tests for procedural features (Sub / Function / Dim / ReDim / Const / Option Explicit / arrays / built-ins).

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
面向过程特性（Sub / Function / Dim / ReDim / Const / Option Explicit / 数组 / 内置函数）的回归测试。

运行：
    python -m unittest discover -s tests -v
"""

import contextlib
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import (                       # noqa: E402
    OVPPArgumentError, OVPPError, OVPPOutOfRangeError,
    OVPPProcedureNotDefinedError, OVPPSyntaxError, OVPPTypeError,
    OVPPVariableNotDefinedError,
)
from OpenVBSPP.ovpprt import builtins as vb_builtins      # noqa: E402
from OpenVBSPP.ovpprt.empty import EMPTY                  # noqa: E402
from OpenVBSPP.runtime import vbs_run                     # noqa: E402


def run(code, filename="<test>"):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename=filename)
    return buf.getvalue()


def out(code):
    return run(code).splitlines()


class TestSub(unittest.TestCase):
    def test_no_parameters(self):
        self.assertEqual(out('Sub Hi()\n echo "hi"\nEnd Sub\nHi'), ["hi"])

    def test_call_keyword(self):
        self.assertEqual(out('Sub Hi()\n echo "hi"\nEnd Sub\nCall Hi'), ["hi"])

    def test_arguments_with_parentheses(self):
        code = 'Sub Show(a, b)\n echo a & "-" & b\nEnd Sub\nCall Show("x", "y")'
        self.assertEqual(out(code), ["x-y"])

    def test_arguments_without_parentheses(self):
        code = 'Sub Show(a, b)\n echo a & "-" & b\nEnd Sub\nShow "x", "y"'
        self.assertEqual(out(code), ["x-y"])

    def test_sub_calls_sub(self):
        code = ('Sub A()\n echo "A"\n B\nEnd Sub\n'
                'Sub B()\n echo "B"\nEnd Sub\nA')
        self.assertEqual(out(code), ["A", "B"])

    def test_procedure_defined_after_use(self):
        self.assertEqual(out('Hi\nSub Hi()\n echo "late"\nEnd Sub'), ["late"])

    def test_exit_sub(self):
        code = ('Sub Guard(n)\n If n < 0 Then\n  echo "neg"\n  Exit Sub\n End If\n'
                ' echo "pos"\nEnd Sub\nGuard -1\nGuard 1')
        self.assertEqual(out(code), ["neg", "pos"])


class TestFunction(unittest.TestCase):
    def test_return_value(self):
        code = 'Function Add(a, b)\n Add = a + b\nEnd Function\necho Add(2, 3)'
        self.assertEqual(out(code), ["5"])

    def test_exit_function(self):
        code = ('Function Sign(n)\n Sign = 0\n If n > 0 Then\n  Sign = 1\n'
                '  Exit Function\n End If\n If n < 0 Then\n  Sign = -1\n End If\n'
                'End Function\necho Sign(7) & Sign(-2) & Sign(0)')
        self.assertEqual(out(code), ["1-10"])

    def test_recursion(self):
        code = ('Function Fact(n)\n If n <= 1 Then\n  Fact = 1\n Else\n'
                '  Fact = n * Fact(n - 1)\n End If\nEnd Function\necho Fact(5)')
        self.assertEqual(out(code), ["120"])

    def test_recursion_fibonacci(self):
        code = ('Function Fib(n)\n If n < 2 Then\n  Fib = n\n Else\n'
                '  Fib = Fib(n - 1) + Fib(n - 2)\n End If\nEnd Function\necho Fib(10)')
        self.assertEqual(out(code), ["55"])

    def test_function_default_return_is_empty(self):
        self.assertEqual(run("Function F()\nEnd Function\necho F()"), "\n")

    def test_nested_function_call_as_argument(self):
        code = ('Function Double(n)\n Double = n * 2\nEnd Function\n'
                'Function Add(a, b)\n Add = a + b\nEnd Function\n'
                'echo Add(Double(2), Double(3))')
        self.assertEqual(out(code), ["10"])


class TestByRefAndByVal(unittest.TestCase):
    def test_byref_is_default(self):
        code = ('Sub Swap(a, b)\n Dim t\n t = a\n a = b\n b = t\nEnd Sub\n'
                'x = 1\ny = 2\nSwap x, y\necho x & "," & y')
        self.assertEqual(out(code), ["2,1"])

    def test_explicit_byref(self):
        code = ('Sub Bump(ByRef n)\n n = n + 1\nEnd Sub\nx = 1\nBump x\necho x')
        self.assertEqual(out(code), ["2"])

    def test_byval_does_not_write_back(self):
        code = ('Sub Bump(ByVal n)\n n = n + 100\n echo n\nEnd Sub\n'
                'x = 1\nBump x\necho x')
        self.assertEqual(out(code), ["101", "1"])

    def test_byref_function_writeback_in_expression(self):
        code = ('Function Bump(ByRef a)\n a = a + 1\n Bump = a * 10\nEnd Function\n'
                'x = 1\ny = Bump(x)\necho x & "," & y')
        self.assertEqual(out(code), ["2,20"])

    def test_vbs_assignment_order_return_value_wins(self):
        # VBS: ByRef write-back happens first, then the return value is assigned to x
        # 中文：VBS：先完成 ByRef 写回，再把返回值赋给 x
        code = ('Function Bump(ByRef a)\n a = a + 1\n Bump = a * 10\nEnd Function\n'
                'x = 1\nx = Bump(x)\necho x')
        self.assertEqual(out(code), ["20"])

    def test_array_byref(self):
        code = ('Sub Fill(ByRef arr, ByVal n)\n Dim i\n For i = 0 To n - 1\n'
                '  arr(i) = i * i\n Next\nEnd Sub\n'
                'Dim a(4)\nFill a, 5\necho a(0) & a(4)\necho UBound(a)')
        self.assertEqual(out(code), ["016", "4"])

    def test_literal_argument_to_byref_parameter(self):
        # a literal has no writable target, so this must not raise
        # 中文：字面量没有可写回的目标，不应报错
        code = 'Sub Bump(a)\n a = a + 1\n echo a\nEnd Sub\nBump 1'
        self.assertEqual(out(code), ["2"])

    def test_redim_on_byref_array_parameter(self):
        code = ('Sub Grow(ByRef arr)\n ReDim arr(3)\n arr(0) = 7\nEnd Sub\n'
                'Dim a()\nGrow a\necho UBound(a) & a(0)')
        self.assertEqual(out(code), ["37"])

    def test_sub_to_sub_byref_passing(self):
        code = ('Sub Inner(ByRef n)\n n = n * 2\nEnd Sub\n'
                'Sub Outer(ByRef n)\n Inner n\n Inner n\nEnd Sub\n'
                'x = 3\nOuter x\necho x')
        self.assertEqual(out(code), ["12"])

    def test_recursive_byref_accumulator(self):
        code = ('Sub CountDown(ByRef total, ByVal n)\n If n <= 0 Then\n  Exit Sub\n End If\n'
                ' total = total + n\n CountDown total, n - 1\nEnd Sub\n'
                'Dim t\nt = 0\nCountDown t, 4\necho t')
        self.assertEqual(out(code), ["10"])


class TestProceduresInContext(unittest.TestCase):
    def test_function_with_select_case(self):
        code = ('Function Kind(n)\n Select Case n\n  Case 0\n   Kind = "zero"\n'
                '  Case 1, 2\n   Kind = "small"\n  Case Else\n   Kind = "big"\n'
                ' End Select\nEnd Function\n'
                'echo Kind(0) & "," & Kind(2) & "," & Kind(9)')
        self.assertEqual(out(code), ["zero,small,big"])

    def test_local_array_inside_function(self):
        code = ('Function MaxOf(n)\n Dim arr(4)\n Dim i\n For i = 0 To 4\n'
                '  arr(i) = (i + 1) * n\n Next\n MaxOf = arr(4)\nEnd Function\n'
                'echo MaxOf(3)')
        self.assertEqual(out(code), ["15"])

    def test_empty_sub_and_function(self):
        code = ('Sub Nop()\nEnd Sub\nFunction Nop2()\nEnd Function\n'
                'Nop\nx = Nop2()\necho "[" & x & "]"')
        self.assertEqual(out(code), ["[]"])

    def test_private_and_public_prefixes(self):
        code = ('Private Sub A()\n echo "a"\nEnd Sub\n'
                'Public Sub B()\n A\n echo "b"\nEnd Sub\nB')
        self.assertEqual(out(code), ["a", "b"])

    def test_procedure_names_are_case_insensitive(self):
        code = 'Sub MiXeD()\n echo "mixed"\nEnd Sub\nmixed\nMIXED'
        self.assertEqual(out(code), ["mixed", "mixed"])


class TestDimAndArrays(unittest.TestCase):
    def test_dim_scalar_defaults_to_empty(self):
        self.assertEqual(run("Dim x\necho x"), "\n")

    def test_empty_behaves_like_zero_and_blank(self):
        code = 'Dim x\necho x + 1\necho "[" & x & "]"\necho x == 0'
        self.assertEqual(out(code), ["1", "[]", "True"])

    def test_dim_multiple(self):
        self.assertEqual(out("Dim a, b\na = 1\nb = 2\necho a + b"), ["3"])

    def test_array_upper_bound_is_inclusive(self):
        code = "Dim a(3)\na(0) = 1\na(3) = 4\necho UBound(a)\necho a(0) & a(3)"
        self.assertEqual(out(code), ["3", "14"])

    def test_two_dimensional_array(self):
        code = "Dim g(1, 2)\ng(1, 2) = 99\ng(0, 0) = 7\necho g(1, 2) & g(0, 0)"
        self.assertEqual(out(code), ["997"])

    def test_dynamic_array_and_redim(self):
        code = ("Dim a()\nReDim a(2)\na(0) = 5\nReDim Preserve a(4)\n"
                "echo UBound(a) & a(0)")
        self.assertEqual(out(code), ["45"])

    def test_redim_discards_values(self):
        code = "Dim a(2)\na(0) = 5\nReDim a(2)\necho a(0) == 0"
        self.assertEqual(out(code), ["True"])

    def test_array_parameter_declaration(self):
        # the classic Sub P(a()) form is accepted too
        # 中文：经典写法 Sub P(a()) 也会被接受
        code = ('Sub Fill(a())\n a(0) = 42\nEnd Sub\nDim x(2)\nFill x\necho x(0)')
        self.assertEqual(out(code), ["42"])


class TestConstAndOptionExplicit(unittest.TestCase):
    def test_const(self):
        self.assertEqual(out('Const PI = 3\nConst NAME = "vbs"\necho PI & NAME'),
                         ["3vbs"])

    def test_option_explicit_accepts_declared(self):
        self.assertEqual(out("Option Explicit\nDim x\nx = 1\necho x"), ["1"])

    def test_option_explicit_rejects_undeclared_assignment(self):
        with self.assertRaises(OVPPVariableNotDefinedError) as ctx:
            run("Option Explicit\nx = 1")
        self.assertEqual(ctx.exception.line_number, 2)
        self.assertIn("Dim", ctx.exception.message)

    def test_option_explicit_rejects_undeclared_read(self):
        with self.assertRaises(OVPPVariableNotDefinedError):
            run("Option Explicit\nDim x\necho y")

    def test_option_explicit_inside_procedure(self):
        code = ("Option Explicit\nSub P()\n Dim local\n local = 1\n echo local\nEnd Sub\nP")
        self.assertEqual(out(code), ["1"])

    def test_option_explicit_rejects_undeclared_local(self):
        with self.assertRaises(OVPPVariableNotDefinedError):
            run("Option Explicit\nSub P()\n local = 1\nEnd Sub\nP")

    def test_option_explicit_allows_module_variable_in_procedure(self):
        code = ("Option Explicit\nDim total\nSub Add1(n)\n total = total + n\nEnd Sub\n"
                "Add1 5\nAdd1 2\necho total")
        self.assertEqual(out(code), ["7"])

    def test_option_explicit_must_be_top_level(self):
        with self.assertRaises(OVPPSyntaxError):
            run("Sub P()\n Option Explicit\nEnd Sub")


class TestScope(unittest.TestCase):
    def test_procedure_modifies_script_level_variable(self):
        code = ("counter = 0\nSub Inc()\n counter = counter + 1\nEnd Sub\n"
                "Inc\nInc\necho counter")
        self.assertEqual(out(code), ["2"])

    def test_procedure_local_does_not_leak(self):
        code = ('Sub P()\n Dim inner\n inner = 1\nEnd Sub\nP\necho inner')
        with self.assertRaises(OVPPVariableNotDefinedError.__bases__[0]):
            run(code)

    def test_loop_variable_inside_procedure_is_local(self):
        code = ('i = 99\nSub P()\n For i = 1 To 2\n  echo i\n Next\nEnd Sub\nP\necho i')
        # i already exists at script level; the procedure mutates the same variable (VBS late binding)
        # 中文：i 在脚本级已存在，过程里修改的是同一个变量（VBS 的后期绑定行为）
        self.assertEqual(out(code), ["1", "2", "3"])


class TestBuiltinFunctions(unittest.TestCase):
    def test_string_functions(self):
        code = ('echo Len("hello")\necho UCase("abc")\necho LCase("ABC")\n'
                'echo Left("abcdef", 2)\necho Right("abcdef", 2)\n'
                'echo Mid("abcdef", 2, 3)\necho Trim("  x  ")')
        self.assertEqual(out(code), ["5", "ABC", "abc", "ab", "ef", "bcd", "x"])

    def test_instr_replace_reverse(self):
        code = ('echo InStr("hello world", "world")\necho InStr(1, "abcabc", "c")\n'
                'echo Replace("a-b-c", "-", "+")\necho StrReverse("abc")\n'
                'echo InStr("abc", "z")')
        self.assertEqual(out(code), ["7", "3", "a+b+c", "cba", "0"])

    def test_math_functions(self):
        code = 'echo Abs(-3)\necho Int(3.9)\necho Fix(-3.9)\necho Sgn(-5)\necho Sqr(9)'
        self.assertEqual(out(code), ["3", "3", "-3", "-1", "3"])

    def test_conversion_functions(self):
        code = ('echo CStr(12) & "x"\necho CInt("42")\necho CDbl("1.5")\n'
                'echo IsNumeric("12") & IsNumeric("abc")')
        self.assertEqual(out(code), ["12x", "42", "1.5", "TrueFalse"])

    def test_array_functions(self):
        code = ('Dim parts\nparts = Split("a,b,c", ",")\n'
                'echo UBound(parts)\necho parts(1)\necho Join(parts, "|")\n'
                'echo LBound(parts)\necho UBound(Array(1, 2, 3))')
        self.assertEqual(out(code), ["2", "b", "a|b|c", "0", "2"])

    def test_isemptyand_typename(self):
        code = ('Dim x\necho IsEmpty(x)\necho IsArray(Array(1))\n'
                'echo TypeName("s")\necho TypeName(1)')
        self.assertEqual(out(code), ["True", "True", "String", "Integer"])

    def test_msgbox(self):
        self.assertEqual(out('x = MsgBox("hello")\necho x'), ["hello", "1"])

    def test_builtin_registry_is_lowercase(self):
        for name in vb_builtins.REGISTRY:
            self.assertEqual(name, name.lower())


class TestProceduralErrors(unittest.TestCase):
    def test_wrong_argument_count(self):
        with self.assertRaises(OVPPArgumentError):
            run("Sub S(a)\nEnd Sub\nS 1, 2")

    def test_undefined_procedure(self):
        with self.assertRaises(OVPPProcedureNotDefinedError):
            run("Nope 1")

    def test_sub_used_as_function(self):
        with self.assertRaises(OVPPTypeError):
            run("Sub S()\nEnd Sub\nx = S()")

    def test_nested_procedure_rejected(self):
        with self.assertRaises(OVPPSyntaxError):
            run("Sub A()\n Sub B()\n End Sub\nEnd Sub")

    def test_exit_sub_outside_sub(self):
        with self.assertRaises(OVPPSyntaxError):
            run("exit sub")

    def test_duplicate_procedure(self):
        with self.assertRaises(OVPPError):
            run("Sub A()\nEnd Sub\nSub A()\nEnd Sub")

    def test_array_index_out_of_range(self):
        with self.assertRaises(OVPPOutOfRangeError) as ctx:
            run("Dim a(1)\necho a(5)")
        self.assertEqual(ctx.exception.line_number, 2)

    def test_using_scalar_as_array(self):
        with self.assertRaises(OVPPTypeError.__bases__[0]):
            run("x = 1\necho x(0)")

    def test_procedure_name_in_error_points_at_line(self):
        with self.assertRaises(OVPPProcedureNotDefinedError) as ctx:
            run('echo 1\nNope\n')
        self.assertEqual(ctx.exception.line_number, 2)


class TestEmptyValue(unittest.TestCase):
    def test_str_and_repr(self):
        self.assertEqual(str(EMPTY), "")
        self.assertEqual(repr(EMPTY), "Empty")
        self.assertFalse(bool(EMPTY))

    def test_arithmetic(self):
        self.assertEqual(EMPTY + 1, 1)
        self.assertEqual(1 + EMPTY, 1)
        self.assertEqual(EMPTY - 2, -2)
        self.assertEqual(EMPTY * 5, 0)

    def test_comparisons(self):
        self.assertTrue(EMPTY == 0)
        self.assertTrue(EMPTY == "")
        self.assertTrue(EMPTY == EMPTY)
        self.assertFalse(EMPTY != 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
