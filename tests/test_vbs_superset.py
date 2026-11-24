"""VBScript superset conformance tests (project rule: any valid VBScript program must compile and run).

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
VBS 超集一致性测试。

项目约定（见 doc/VBS_SUPERSET.md）：**任何合法的 VBScript 程序都必须能编译运行**，
除非它落在 KNOWN_GAPS 里。这里把这条约定变成可执行的检查：

  * VBS_FEATURES —— 合法 VBS 片段，必须能跑出预期结果；
  * KNOWN_GAPS   —— 已知缺口，用 expectedFailure 标记；
    某个缺口被补上时，对应的 expectedFailure 会变成 "unexpected success"，
    提醒把用例移进 VBS_FEATURES 并更新文档。

运行：
    python -m unittest tests.test_vbs_superset -v
"""

import contextlib
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.runtime import vbs_run       # noqa: E402


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<vbs-superset>")
    return buf.getvalue().splitlines()


# --------------------------------------------------------------------- valid VBScript / 中文：合法 VBS
VBS_FEATURES = [
    # statements and separators
    # 中文：语句与分隔
    ("colon separators", "a = 1 : b = 2 : echo a + b", ["3"]),
    ("line continuation",
     "total = 1 + _\n 2 + _\n 3\necho total", ["6"]),
    ("rem comment", "rem hello\necho 1", ["1"]),
    ("apostrophe comment", "' hello\necho 1", ["1"]),
    ("blank lines", "\n\na = 1\n\necho a\n", ["1"]),

    # assignment
    # 中文：赋值
    ("plain assignment", "a = 1\necho a", ["1"]),
    ("Let statement", "Let a = 5\necho a", ["5"]),
    ("Set statement", "Set a = 3\necho a", ["3"]),
    ("multiple assignment separators", "a = 1 : a = a + 1 : echo a", ["2"]),

    # comparison: VBS uses = for equality
    # 中文：比较：VBS 用 = 表示相等
    ("equality with =", "echo 1 = 1", ["True"]),
    ("inequality <>", "echo 1 <> 2", ["True"]),
    ("If with =", "If 2 = 2 Then\n echo \"ok\"\nEnd If", ["ok"]),
    ("If with <>", "If 2 <> 3 Then\n echo \"ok\"\nEnd If", ["ok"]),

    # the If family
    # 中文：If 家族
    ("single line If", "If 1 = 1 Then echo \"yes\"", ["yes"]),
    ("single line If Else", "If 1 = 2 Then echo \"a\" Else echo \"b\"", ["b"]),
    ("single line If with colon",
     "If 1 = 1 Then echo \"a\" : echo \"b\"", ["a", "b"]),
    ("block If", "If 1 = 1 Then\n echo \"a\"\nEnd If", ["a"]),
    ("ElseIf chain",
     "x = 2\nIf x = 1 Then\n echo 1\nElseIf x = 2 Then\n echo 2\nElse\n echo 3\nEnd If",
     ["2"]),
    ("nested If", "If 1 = 1 Then\n If 2 = 2 Then\n  echo \"deep\"\n End If\nEnd If", ["deep"]),

    # loops
    # 中文：循环
    ("For Next", "For i = 1 To 3\n echo i\nNext", ["1", "2", "3"]),
    ("For Step", "For i = 1 To 5 Step 2\n echo i\nNext", ["1", "3", "5"]),
    ("For Each", "For Each x In Array(1, 2)\n echo x\nNext", ["1", "2"]),
    ("Exit For", "For i = 1 To 9\n If i = 2 Then Exit For\n echo i\nNext", ["1"]),
    ("Do While", "i = 0\nDo While i < 2\n echo i\n i = i + 1\nLoop", ["0", "1"]),
    ("Do Until", "i = 0\nDo Until i = 2\n echo i\n i = i + 1\nLoop", ["0", "1"]),
    ("Do Loop While", "i = 0\nDo\n i = i + 1\nLoop While i < 2\necho i", ["2"]),
    ("Do Loop Until", "i = 0\nDo\n i = i + 1\nLoop Until i = 2\necho i", ["2"]),
    ("While Wend", "i = 0\nWhile i < 2\n echo i\n i = i + 1\nWend", ["0", "1"]),
    ("Exit Do", "i = 0\nDo\n i = i + 1\n If i = 2 Then Exit Do\nLoop\necho i", ["2"]),

    # Select Case
    ("Select Case basic",
     "x = 2\nSelect Case x\nCase 1\n echo \"a\"\nCase 2\n echo \"b\"\nEnd Select", ["b"]),
    ("Select Case list",
     "x = 3\nSelect Case x\nCase 1, 3\n echo \"hit\"\nCase Else\n echo \"miss\"\nEnd Select",
     ["hit"]),
    ("Select Case range",
     "x = 4\nSelect Case x\nCase 1 To 5\n echo \"in\"\nEnd Select", ["in"]),
    ("Select Case Is =",
     "x = 2\nSelect Case x\nCase Is = 2\n echo \"two\"\nEnd Select", ["two"]),
    ("Select Case Is >=",
     "x = 9\nSelect Case x\nCase Is >= 5\n echo \"big\"\nEnd Select", ["big"]),

    # procedures
    # 中文：过程
    ("Sub", "Sub Hi()\n echo \"hi\"\nEnd Sub\nHi", ["hi"]),
    ("Sub with args", "Sub P(a, b)\n echo a + b\nEnd Sub\nP 1, 2", ["3"]),
    ("Call statement", "Sub P(a)\n echo a\nEnd Sub\nCall P(7)", ["7"]),
    ("chained member calls",
     "Class Child\n"
     " Public Sub Run()\n"
     "  echo 42\n"
     " End Sub\n"
     "End Class\n"
     "Class Parent\n"
     " Public Function GetChild(n)\n"
     "  Set GetChild = New Child\n"
     " End Function\n"
     "End Class\n"
     "Set p = New Parent\n"
     "p.GetChild(1).Run\n"
     "Call p.GetChild(1).Run()",
     ["42", "42"]),
    ("Function", "Function F(a)\n F = a * 2\nEnd Function\necho F(3)", ["6"]),
    ("ByRef", "Sub B(n)\n n = n + 1\nEnd Sub\nx = 1\nB x\necho x", ["2"]),
    ("ByVal", "Sub B(ByVal n)\n n = n + 1\nEnd Sub\nx = 1\nB x\necho x", ["1"]),
    ("Exit Sub", "Sub P()\n Exit Sub\n echo \"no\"\nEnd Sub\nP\necho \"ok\"", ["ok"]),
    ("Exit Function", "Function F()\n F = 1\n Exit Function\nEnd Function\necho F()", ["1"]),
    ("recursion",
     "Function Fact(n)\n If n <= 1 Then\n  Fact = 1\n Else\n  Fact = n * Fact(n - 1)\n End If\nEnd Function\necho Fact(4)",
     ["24"]),

    # declarations and arrays
    # 中文：声明与数组
    ("Dim", "Dim a\na = 1\necho a", ["1"]),
    ("Dim multiple", "Dim a, b\na = 1 : b = 2\necho a + b", ["3"]),
    ("Dim array", "Dim a(2)\na(0) = 1\necho UBound(a)", ["2"]),
    ("ReDim", "Dim a()\nReDim a(1)\na(0) = 3\necho a(0)", ["3"]),
    ("ReDim Preserve", "Dim a(1)\na(0) = 3\nReDim Preserve a(3)\necho a(0)", ["3"]),
    ("Const", "Const N = 5\necho N", ["5"]),
    ("Erase", "Dim a(1)\na(0) = 9\nErase a\necho IsEmpty(a(0))", ["True"]),
    ("Option Explicit", "Option Explicit\nDim x\nx = 1\necho x", ["1"]),

    # error handling
    # 中文：错误处理
    ("On Error Resume Next",
     "On Error Resume Next\nx = 1 / 0\necho \"after\"\necho Err.Number", ["after", "11"]),
    ("On Error GoTo 0",
     "On Error Resume Next\nx = 1 / 0\nOn Error GoTo 0\necho Err.Number", ["11"]),
    ("Err.Clear",
     "On Error Resume Next\nx = 1 / 0\nErr.Clear\necho Err.Number", ["0"]),
    ("Err.Raise",
     "On Error Resume Next\nErr.Raise 1001, \"s\", \"boom\"\necho Err.Number", ["1001"]),

    # With
    ("With block",
     "With Err\n .Number = 5\nEnd With\necho Err.Number", ["5"]),
    ("With + method",
     "With Err\n .Number = 9\n .Clear()\nEnd With\necho Err.Number", ["0"]),

    # values
    # 中文：值
    ("Empty literal", "echo IsEmpty(Empty)", ["True"]),
    ("Null literal", "echo IsNull(Null)", ["True"]),
    ("Nothing literal", "echo Nothing Is Nothing", ["True"]),
    ("Is operator", "Dim o\no = Nothing\necho o Is Nothing", ["True"]),
    ("boolean literals", "echo True\necho False", ["True", "False"]),
    ("float literal", "echo 3.5", ["3.5"]),
    ("string concat", "echo \"a\" & \"b\"", ["ab"]),
    ("Mod operator", "echo 7 Mod 2", ["1"]),
    ("integer division", "echo 7 \\ 2", ["3"]),
    ("exponent", "echo 2 ^ 3", ["8"]),
    ("Xor", "echo 5 Xor 3", ["6"]),
    ("Eqv", "echo 6 Eqv 3", ["-6"]),
    ("Imp", "echo 0 Imp 1", ["-1"]),
    ("vbCrLf constant", "echo \"a\" & vbCrLf & \"b\"", ["a", "b"]),
    ("vbTrue constant", "echo vbTrue + 0", ["-1"]),

    # arithmetic coercion (VBS: + concatenates only when both sides are text)
    # 中文：算术强制转换（VBS 的 + 只在两侧都是文本时拼接）
    ("string + string concatenates", "echo \"1\" + \"2\"", ["12"]),
    ("number + numeric string", "echo 1 + \"2\"", ["3"]),
    ("numeric string + number", "echo \"2\" + 1", ["3"]),
    ("numeric string subtraction", "echo \"6\" - \"4\"", ["2"]),
    ("numeric string multiplication", "echo \"2\" * 3", ["6"]),
    ("numeric string division", "echo \"6\" / \"3\"", ["2"]),
    ("division always yields a Double", "echo 4 / 2", ["2"]),
    ("exponent always yields a Double", "echo 2 ^ 10", ["1024"]),
    ("Empty + Empty is zero", "echo Empty + Empty", ["0"]),
    ("True + True is -2", "echo True + True", ["-2"]),
    ("non numeric string is a type mismatch",
     "On Error Resume Next\nx = \"abc\" + 1\necho Err.Number", ["13"]),

    # Null propagation through built-ins
    # 中文：Null 在内置函数中的传播
    ("Null propagates through +", "echo IsNull(Null + 1)", ["True"]),
    ("Len(Null) is Null", "echo IsNull(Len(Null))", ["True"]),
    ("UCase(Null) is Null", "echo IsNull(UCase(Null))", ["True"]),
    ("Trim(Null) is Null", "echo IsNull(Trim(Null))", ["True"]),
    ("Abs(Null) is Null", "echo IsNull(Abs(Null))", ["True"]),
    ("TypeName(Null)", "echo TypeName(Null)", ["Null"]),
    ("& treats Null as empty", "echo \"a\" & Null & \"b\"", ["ab"]),
    ("Null & Null is Null", "echo IsNull(Null & Null)", ["True"]),
    ("CStr(Null) raises 94",
     "On Error Resume Next\nx = CStr(Null)\necho Err.Number", ["94"]),

    # undeclared variables read as Empty (no Option Explicit)
    # 中文：未声明变量读作 Empty（没有 Option Explicit 时）
    ("undeclared variable is Empty", "echo TypeName(undeclaredvar)", ["Empty"]),
    ("undeclared variable concatenates", "echo undeclaredvar & 1", ["1"]),
    ("undeclared variable adds as zero", "echo undeclaredvar + 1", ["1"]),
    ("undeclared variable inside a Sub",
     "Sub P()\n echo TypeName(innerthing)\nEnd Sub\nP", ["Empty"]),

    # object model
    # 中文：对象模型
    ("Class + New",
     "Class Foo\n Public x\nEnd Class\nSet f = New Foo\nf.x = 5\necho f.x", ["5"]),
    ("Class method",
     "Class Foo\n Public Function Twice(n)\n  Twice = n * 2\n End Function\n"
     "End Class\nSet f = New Foo\necho f.Twice(4)", ["8"]),
    ("Class property",
     "Class Foo\n Private n\n Public Property Get Value()\n  Value = n\n End Property\n"
     " Public Property Let Value(v)\n  n = v\n End Property\nEnd Class\n"
     "Set f = New Foo\nf.Value = 7\necho f.Value", ["7"]),
    ("Class_Initialize",
     "Class Foo\n Public x\n Private Sub Class_Initialize()\n  x = 42\n End Sub\n"
     "End Class\nSet f = New Foo\necho f.x", ["42"]),
    ("CreateObject script class",
     "Class Foo\n Public x\nEnd Class\nSet f = CreateObject(\"Foo\")\nf.x = 3\necho f.x",
     ["3"]),
    ("Dictionary",
     "Set d = CreateObject(\"Scripting.Dictionary\")\nd.Add \"a\", 1\necho d(\"a\")\n"
     "echo d.Count", ["1", "1"]),
    ("For Each Dictionary",
     "Set d = CreateObject(\"Scripting.Dictionary\")\nd.Add \"k\", 1\n"
     "For Each k In d\n echo k\nNext", ["k"]),
    ("Is Nothing",
     "Set f = Nothing\necho f Is Nothing", ["True"]),
    ("TypeName object",
     "Class Foo\nEnd Class\nSet f = New Foo\necho TypeName(f)", ["Foo"]),
    ("GetObject",
     "Class Foo\n Public x\nEnd Class\nSet f = GetObject(\"Foo\")\nf.x = 4\necho f.x",
     ["4"]),
    ("private field via property",
     "Class Foo\n Private n\n Public Property Get Value()\n  Value = n\n"
     " End Property\n Public Sub Set_(v)\n  n = v\n End Sub\nEnd Class\n"
     "Set f = New Foo\nf.Set_ 6\necho f.Value", ["6"]),

    # dynamic execution and formatting
    # 中文：动态执行与格式化
    ("Execute", "x = 0\nExecute \"x = 41 + 1\"\necho x", ["42"]),
    ("Eval", "echo Eval(\"1 + 1\")", ["2"]),
    ("FormatNumber", "echo FormatNumber(1234.5, 2)", ["1,234.50"]),
    ("FormatCurrency", "echo FormatCurrency(3)", ["$3.00"]),
    ("FormatPercent", "echo FormatPercent(0.5)", ["50.00%"]),
    ("FormatDateTime", "echo FormatDateTime(\"2027-01-02 03:04:05\", vbShortDate)",
     ["2027-01-02"]),
    ("Option Compare Text", "Option Compare Text\necho \"A\" = \"a\"", ["True"]),
    ("Option Compare Binary", "echo \"A\" = \"a\"", ["False"]),

    # parameter shapes (native VBS)
    # 中文：参数形态（VBS 原生）
    ("Optional parameter",
     "Sub Greet(name, Optional greeting = \"hi\")\n echo greeting & \" \" & name\nEnd Sub\nGreet \"ann\"\nGreet \"bob\", \"yo\"",
     ["hi ann", "yo bob"]),
    ("ParamArray",
     "Function Total(ParamArray nums())\n Dim i, s\n s = 0\n For i = 0 To UBound(nums)\n  s = s + nums(i)\n Next\n Total = s\nEnd Function\necho Total(1, 2, 3)",
     ["6"]),

    # built-in functions
    # 中文：内置函数
    ("Len", "echo Len(\"abc\")", ["3"]),
    ("Mid", "echo Mid(\"abcdef\", 2, 3)", ["bcd"]),
    ("InStr", "echo InStr(\"abc\", \"b\")", ["2"]),
    ("UCase/LCase", "echo UCase(\"a\") & LCase(\"B\")", ["Ab"]),
    ("Trim", "echo Trim(\"  x  \")", ["x"]),
    ("Split/Join", "Dim p\np = Split(\"a,b\", \",\")\necho Join(p, \"-\")", ["a-b"]),
    ("Array/UBound", "echo UBound(Array(1, 2, 3))", ["2"]),
    ("CStr/CInt", "echo CStr(1) & CInt(\"2\")", ["12"]),
    ("Abs/Int", "echo Abs(-2) & Int(2.9)", ["22"]),
    ("IsNumeric", "echo IsNumeric(\"12\")", ["True"]),
    ("TypeName", "echo TypeName(1)", ["Integer"]),
]


class TestVBSSuperset(unittest.TestCase):
    """Valid VBScript must run unchanged.

    中文：合法 VBS 代码必须能原样运行。"""

    def test_features(self):
        for name, code, expected in VBS_FEATURES:
            with self.subTest(feature=name):
                self.assertEqual(out(code), expected, f"feature: {name}")


# --------------------------------------------------------------------- known gaps / 中文：已知缺口
# these are valid VBScript but not implemented yet (see doc/VBS_SUPERSET.md).
# 中文：这些是合法 VBS，但 OpenVBS++ 还没实现（见 doc/VBS_SUPERSET.md）。
# note: once a gap is fixed the expectedFailure turns into an unexpected success,
# 中文：注意：缺口补齐后 expectedFailure 会变成 unexpected success，
# then move the case into VBS_FEATURES and update the docs.
# 中文：那时应该把用例移到 VBS_FEATURES，并同步更新文档。
#
# the object model (Class / Property / CreateObject / Dictionary) landed in v0.4,
# 中文：对象模型（Class / Property / CreateObject / Dictionary）已在 v0.4 补齐，
# the remaining gaps concern dynamic execution and formatting functions.
# 中文：剩下的缺口是执行动态代码与格式化函数相关。

# --------------------------------------------------------------------- known gaps / 中文：已知缺口
# since v0.5 the gaps listed earlier (Class/Property/CreateObject/Dictionary/Execute/Eval/
# 中文：v0.5 起，之前列出的缺口（Class/Property/CreateObject/Dictionary/Execute/Eval/
# formatting functions / Option Compare Text) are in place.
# 中文：格式化函数/Option Compare Text）全部补齐。
#
# what is left: comparisons compare by type instead of coercing. VBScript turns
# `1 = "1"` into True (numeric comparison), makes `"2" > 10` False and raises 13
# for `"abc" = 5`; we use Python's rules, which give False, error 13 and False
# respectively. Closing this means routing every comparison through a runtime
# helper, which costs ~60 ns per comparison — see doc/VBS_SUPERSET.md §4.
# 中文：剩下的缺口是比较——我们按类型比较而不做强制转换。VBS 里 `1 = "1"` 为 True
# （按数值比）、`"2" > 10` 为 False、`"abc" = 5` 报 13；我们用 Python 规则，
# 分别得到 False、错误 13、False。补齐它需要让每次比较都走运行时函数，
# 每次约 60 ns——见 doc/VBS_SUPERSET.md §4。
class TestKnownGaps(unittest.TestCase):
    """Valid VBScript that OpenVBS++ does not reproduce yet (§4).

    中文：OpenVBS++ 暂时还跑不出同样结果的合法 VBS（见 §4）。"""

    @unittest.expectedFailure
    def test_number_equals_numeric_text(self):
        # VBS: 1 = "1" → True (convert the text, compare as numbers)
        # 中文：VBS 里 1 = "1" 为 True（把文本转成数值再比）
        self.assertEqual(out('echo 1 = "1"'), ["True"])

    @unittest.expectedFailure
    def test_numeric_text_is_greater_than_number(self):
        # VBS: "2" > 10 → False; we raise error 13 instead.
        # 中文：VBS 里 "2" > 10 为 False；我们报错误 13。
        self.assertEqual(out('echo "2" > 10'), ["False"])

    @unittest.expectedFailure
    def test_text_compared_with_number_is_type_mismatch(self):
        # VBS: "abc" = 5 → error 13; we return False.
        # 中文：VBS 里 "abc" = 5 报错误 13；我们返回 False。
        code = 'On Error Resume Next\nx = ("abc" = 5)\necho Err.Number'
        self.assertEqual(out(code), ["13"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
