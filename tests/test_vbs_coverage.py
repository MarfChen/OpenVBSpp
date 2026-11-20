"""v0.9.21 coverage: the VBScript built-ins and literals that were still missing.

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

v0.9.21 补齐的 VBScript 内置与字面量。

每个用例钉住一处曾经缺失或有偏差的行为：十六/八进制与日期字面量、数学函数、
类型转换、VarType、RGB、ScriptEngine 信息、区域设置、IsMissing、GetRef、
LoadPicture、ExecuteGlobal，以及只有时间的 TimeValue 与 Time()。
"""

import contextlib
import io
import math
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import (                 # noqa: E402
    OVPPError,
    OVPPObjectNotSupportedError,
    OVPPProcedureNotDefinedError,
)
from OpenVBSPP.runtime import vbs_run             # noqa: E402


def out(code):
    """Run a snippet and return its stdout lines.

    中文：运行一段代码，返回 stdout 的每一行。
    """
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<v0921>")
    return buf.getvalue().splitlines()


class TestNumericLiterals(unittest.TestCase):
    """&H / &O literals (missing until 0.9.21).

    中文：十六进制与八进制字面量。
    """

    def test_hex(self):
        self.assertEqual(out("""echo &HFF"""), ["255"])

    def test_hex_is_case_insensitive(self):
        self.assertEqual(out("""echo &hff"""), ["255"])

    def test_octal(self):
        self.assertEqual(out("""echo &O17"""), ["15"])

    def test_mixed_with_arithmetic(self):
        self.assertEqual(out("""echo &H10 + &O10"""), ["24"])


class TestDateLiterals(unittest.TestCase):
    """#...# date literals (missing until 0.9.21).

    中文：#...# 日期字面量。
    """

    def test_iso_date(self):
        self.assertEqual(out("""echo Year(#2020-05-06#)"""), ["2020"])

    def test_us_date(self):
        self.assertEqual(
            out("""echo Month(#1/2/2020#) & "-" & Day(#1/2/2020#)"""),
            ["1-2"])

    def test_date_and_time(self):
        self.assertEqual(out("""echo Hour(#1/2/2020 3:04:05 PM#)"""), ["15"])

    def test_date_literal_is_a_date(self):
        self.assertEqual(out("""echo VarType(#2020-05-06#)"""), ["7"])


class TestMathBuiltins(unittest.TestCase):
    """Sin / Cos / Tan / Atn / Exp / Log.

    中文：三角与指数、对数函数。
    """

    def test_trigonometry(self):
        self.assertEqual(out("""echo Sin(0)"""), ["0"])
        self.assertEqual(out("""echo Cos(0)"""), ["1"])
        self.assertEqual(out("""echo Tan(0)"""), ["0"])

    def test_arctangent(self):
        self.assertAlmostEqual(
            float(out("""echo Atn(1)""")[0]), math.pi / 4, places=12)

    def test_exponential_and_log(self):
        self.assertEqual(out("""echo Exp(0)"""), ["1"])
        self.assertAlmostEqual(
            float(out("""echo Log(Exp(1))""")[0]), 1.0, places=12)

    def test_log_of_non_positive_raises(self):
        with self.assertRaises(OVPPError):
            out("""echo Log(0)""")


class TestConversions(unittest.TestCase):
    """CCur / CSng / CDate.

    中文：货币、单精度与日期转换。
    """

    def test_ccur_rounds_to_four_places(self):
        self.assertEqual(out("""echo CCur(1.23456789)"""), ["1.2346"])

    def test_csng_keeps_single_precision(self):
        self.assertEqual(out("""echo CSng(1.23456789)"""), ["1.234568"])

    def test_cdate(self):
        self.assertEqual(
            out("""echo Year(CDate("2020-05-06"))"""), ["2020"])


class TestVarType(unittest.TestCase):
    """VarType constants follow the VBScript/VBA numbering.

    中文：VarType 沿用 VBScript/VBA 的编号。
    """

    def test_scalars(self):
        self.assertEqual(
            out("""echo VarType(1) & "," & VarType("a") & "," & VarType(True) """
                """& "," & VarType(1.5) & "," & VarType(Empty)"""),
            ["3,8,11,5,0"])

    def test_date(self):
        self.assertEqual(out("""echo VarType(Now())"""), ["7"])

    def test_array(self):
        self.assertEqual(
            out("""a = Array(1, 2)
echo VarType(a)"""), ["8204"])


class TestRgb(unittest.TestCase):
    """RGB packs into a BGR-long, matching VBScript.

    中文：RGB 按 VBScript 的方式打包成整数。
    """

    def test_channels(self):
        self.assertEqual(out("""echo RGB(255, 0, 0)"""), ["255"])
        self.assertEqual(out("""echo RGB(0, 255, 0)"""), ["65280"])
        self.assertEqual(out("""echo RGB(0, 0, 255)"""), ["16711680"])


class TestScriptEngineInfo(unittest.TestCase):
    """ScriptEngine and the three version functions.

    中文：ScriptEngine 与三个版本号函数。
    """

    def test_engine_name(self):
        self.assertEqual(out("""echo ScriptEngine()"""), ["VBScript"])

    def test_version_numbers(self):
        self.assertEqual(
            out("""echo ScriptEngineMajorVersion() & "." & """
                """ScriptEngineMinorVersion()"""),
            ["5.8"])

    def test_build_version(self):
        self.assertEqual(out("""echo ScriptEngineBuildVersion()"""), ["0"])


class TestLocale(unittest.TestCase):
    """SetLocale returns the previous LCID; GetLocale reports the current one.

    中文：SetLocale 返回旧的 LCID，GetLocale 报告当前的。
    """

    def test_default(self):
        self.assertEqual(out("""echo GetLocale()"""), ["1033"])

    def test_set_returns_previous(self):
        self.assertEqual(
            out("""echo SetLocale(1041) & " " & GetLocale()"""),
            ["1033 1041"])


class TestIsMissing(unittest.TestCase):
    """IsMissing sees whether an optional argument was supplied.

    中文：IsMissing 判断可选参数是否传了。
    """

    def test_omitted_and_supplied(self):
        code = """Sub S(Optional x)
    echo IsMissing(x)
End Sub
S()
S(1)"""
        self.assertEqual(out(code), ["True", "False"])


class TestGetRef(unittest.TestCase):
    """GetRef returns a reference usable with call syntax.

    中文：GetRef 返回可以按调用语法使用的过程引用。
    """

    def test_call_with_argument(self):
        code = """Function Dbl(n)
    Dbl = n * 2
End Function
Set f = GetRef("Dbl")
echo f(21)"""
        self.assertEqual(out(code), ["42"])

    def test_call_without_arguments(self):
        code = """Function Seven()
    Seven = 7
End Function
Set f = GetRef("Seven")
echo f()"""
        self.assertEqual(out(code), ["7"])

    def test_name_is_case_insensitive(self):
        code = """Function Dbl(n)
    Dbl = n * 2
End Function
Set f = GetRef("dbl")
echo f(4)"""
        self.assertEqual(out(code), ["8"])

    def test_unknown_name_raises(self):
        with self.assertRaises(OVPPProcedureNotDefinedError):
            out("""Set f = GetRef("nope")""")


class TestLoadPicture(unittest.TestCase):
    """LoadPicture exists so old scripts fail with a clear message.

    中文：LoadPicture 存在，只是没有图形库，报错信息要明确。
    """

    def test_raises_not_supported(self):
        with self.assertRaises(OVPPObjectNotSupportedError):
            out("""Set p = LoadPicture("a.bmp")""")


class TestExecuteGlobal(unittest.TestCase):
    """ExecuteGlobal runs in the global scope, like Execute.

    中文：ExecuteGlobal 在全局作用域执行，与 Execute 一致。
    """

    def test_defines_global(self):
        self.assertEqual(out("""ExecuteGlobal "y = 5"
echo y"""), ["5"])

    def test_sees_and_updates_globals(self):
        # Procedures added at run time cannot be called directly: names bind at
        # compile time, so GetRef or a module is the supported route.
        # 中文：运行期新定义的过程无法直接调用 —— 名字在编译期绑定，
        # 中文：需要间接调用请用 GetRef 或独立模块。
        code = """z = 1
ExecuteGlobal "z = z + 1"
echo z"""
        self.assertEqual(out(code), ["2"])


class TestTimeOnly(unittest.TestCase):
    """TimeValue / Time() carry a time only, not a broken date.

    中文：TimeValue / Time() 只含时间，不该被当成坏日期。
    """

    def test_time_value_24h(self):
        self.assertEqual(out("""echo Hour(TimeValue("13:45"))"""), ["13"])

    def test_time_value_ampm(self):
        self.assertEqual(out("""echo Hour(TimeValue("1:45:00 PM"))"""), ["13"])

    def test_time_is_convertible(self):
        value = out("""echo Hour(Time())""")[0]
        self.assertTrue(0 <= int(value) <= 23)


class TestScriptEndAndStop(unittest.TestCase):
    """A bare `End` stops the script; `Stop` is a debugger no-op.

    中文：单独的 `End` 结束整个脚本；`Stop` 是调试断点，这里接受但不做任何事。
    """

    def test_end_stops_the_script(self):
        self.assertEqual(out("""echo 1
End
echo 2"""), ["1"])

    def test_end_at_end_of_file(self):
        self.assertEqual(out("""echo 1
End"""), ["1"])

    def test_end_inside_a_block(self):
        code = """Dim x
x = 1
If x = 1 Then
echo 1
End
echo 2
End If
echo 3"""
        self.assertEqual(out(code), ["1"])

    def test_end_inside_a_procedure(self):
        code = """Sub Boom()
echo 1
End
echo 2
End Sub
Boom
echo 3"""
        self.assertEqual(out(code), ["1"])

    def test_try_cannot_catch_end(self):
        code = """Try
echo 1
End
echo 2
Catch e
echo "caught"
End Try
echo 3"""
        self.assertEqual(out(code), ["1"])

    def test_on_error_resume_next_cannot_swallow_end(self):
        code = """On Error Resume Next
echo 1
End
echo 2"""
        self.assertEqual(out(code), ["1"])

    def test_stop_is_a_no_op(self):
        self.assertEqual(out("""echo 1
Stop
echo 2"""), ["1", "2"])

    def test_end_and_stop_still_need_their_blocks_closed(self):
        # The parser still checks block balance; only the block-closing forms
        # (End If / End Sub / ...) take a keyword after them.
        # 中文：块结构仍然要闭合；只有 `End If` 这类块结束符才带关键字。
        with self.assertRaises(OVPPError):
            out("""If 1 = 1 Then
echo 1
""")


class TestPredefinedConstants(unittest.TestCase):
    """The VBScript constants scripts pass to MsgBox and compare with VarType.

    中文：脚本传给 MsgBox、与 VarType 比较的那些 VBScript 预定义常量。
    """

    def test_msgbox_return_values(self):
        self.assertEqual(out("""echo vbOK & "," & vbCancel & "," & vbYes & "," & vbNo"""),
                         ["1,2,6,7"])

    def test_msgbox_buttons_and_icons(self):
        # vbYesNo + vbQuestion == 4 + 32: the stand-in accepts and ignores it.
        # 中文：vbYesNo + vbQuestion == 4 + 32，控制台替身接受并忽略。
        self.assertEqual(out("""echo vbYesNo + vbQuestion"""), ["36"])
        self.assertEqual(out('MsgBox "save?", vbYesNo + vbQuestion, "t"'),
                         ["save?"])

    def test_vartype_constants(self):
        code = """If VarType("a") = vbString Then
echo "string"
End If
If VarType(1.5) = vbDouble Then
echo "double"
End If"""
        self.assertEqual(out(code), ["string", "double"])

    def test_array_constant(self):
        self.assertEqual(out("""echo vbArray"""), ["8192"])

    def test_source_constants(self):
        self.assertEqual(out("""echo vbObjectError"""), ["-2147221504"])


class TestOmittedArguments(unittest.TestCase):
    """VBScript lets a call leave an argument out: `GetObject(, "WScript.Shell")`.

    中文：VBScript 允许调用时省略实参：`GetObject(, "WScript.Shell")`。
    """

    def test_empty_slot_becomes_empty(self):
        code = ("Sub Q(a, b)\n"
                "  echo TypeName(a) & \"|\" & b\n"
                "End Sub\n"
                "Q , 5")
        self.assertEqual(out(code), ["Empty|5"])

    def test_trailing_comma_keeps_two_arguments(self):
        code = ("Sub R(a, b)\n"
                "  echo TypeName(a) & \"|\" & TypeName(b)\n"
                "End Sub\n"
                "R(1,)")
        self.assertEqual(out(code), ["Integer|Empty"])

    def test_parenthesised_leading_slot(self):
        code = ("Sub S(a, b)\n"
                "  echo TypeName(a) & \"|\" & b\n"
                "End Sub\n"
                "S(, \"x\")")
        self.assertEqual(out(code), ["Empty|x"])

    def test_no_arguments_still_means_none(self):
        code = ("Sub T(a, b)\n"
                "  echo a & b\n"
                "End Sub\n"
                "T")
        with self.assertRaises(Exception):
            out(code)
