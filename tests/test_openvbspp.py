"""OpenVBS++ regression tests.

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
OpenVBS++ 回归测试。

运行：
    python -m unittest discover -s tests -v
或：
    python tests/test_openvbspp.py
"""

import contextlib
import io
import marshal
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.compiler import vbs_compile            # noqa: E402
from OpenVBSPP.exceptions import (                     # noqa: E402
    OVPPSyntaxError, OVPPVariableNotDefinedError, OVPPDivisionByZeroError,
    OVPPTypeError,
)
from OpenVBSPP.ovpprt import boolean as vbool           # noqa: E402
from OpenVBSPP.runtime import vbs_run                   # noqa: E402


EXAMPLES = ROOT / "examples"


def run(code, filename="<test>"):
    """Run code and return what it wrote to stdout.

    中文：运行代码并返回 stdout 内容。"""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename=filename)
    return buf.getvalue()


def out(code, filename="<test>"):
    """Run code and return its output split into lines.

    中文：运行代码并返回按行切分的输出。"""
    return run(code, filename=filename).splitlines()


class TestExamples(unittest.TestCase):
    """Scripts under examples/ must keep running with stable output.

    中文：examples/ 下的脚本必须保持可运行且有稳定输出。"""

    EXPECTED = {
        "arrays.vbs": ["total=150", "5 -> 60", "ann / bob / cid", "ANN3", "42"],
        "classes.vbs": ["counter=7", "count=2", "one=1", "two=2"],
        "comments.vbs": ["3"],
        "control1.vbs": ["15"],
        "control2.vbs": ["45", "45"],
        "control3.vbs": ["45"],
        "control4.vbs": ["B"],
        "control5.vbs": ["B"],
        "control6.vbs": ["hit"],
        "do_post.vbs": ["3", "2"],
        "expression1.vbs": ["3"],
        "fizzbuzz.vbs": ["1", "2", "Fizz", "4", "Buzz", "Fizz", "7", "8", "Fizz",
                         "Buzz", "11", "Fizz", "13", "14", "FizzBuzz"],
        "for_loop.vbs": ["15", "10", "7", "4", "1", "stop at 4"],
        "inheritance.vbs": ["hits=3", "step=4"],
        "operations1.vbs": ["3", "2", "-1", "0.5"],
        "operations2.vbs": ["0", "-1", "0"],
        "on_error.vbs": ["err=11", "ok=5", "raised=1001", "inner=13", "after sub err=13"],
        "option_strict.vbs": ["3", "2.5", "a1", "True", "13", "13"],
        "procedures.vbs": ["-- start", "-- factorial", "120", "2,1",
                           "inside=102", "outside=2"],
        "try_catch.vbs": ["  [finally] 清理", "5", "  [业务错误] 除数不能为 0",
                          "  [finally] 清理", "0", "  内层捕获 53", "  外层捕获 99"],
        "regexp_json.vbs": ["True", "匹配数 2", "  111-2222 @2", "  333-4444 @13",
                            '[1, "two", true]', "2:two", "81"],
        "sugar.vbs": ["3", "a is one", "b is not one", "1", "2", "3"],
        "typing.vbs": ["20", "0.5", "openvbs"],
        "variable_and_print.vbs": ["hello world", "1234"],
    }

    def test_examples(self):
        for name, expected in self.EXPECTED.items():
            path = EXAMPLES / name
            with self.subTest(example=name):
                self.assertTrue(path.is_file(), f"missing example {name}")
                code = path.read_text(encoding="utf-8")
                self.assertEqual(out(code, filename=str(path)), expected)

    def test_every_example_is_covered(self):
        found = {p.name for p in EXAMPLES.glob("*.vbs")}
        self.assertEqual(found, set(self.EXPECTED),
                         "新增/删除示例后请同步更新 EXPECTED")


class TestOperators(unittest.TestCase):
    def test_arithmetic_and_precedence(self):
        self.assertEqual(out("echo 1 + 2 * 3"), ["7"])
        self.assertEqual(out("echo (1 + 2) * 3"), ["9"])
        self.assertEqual(out("echo 7 - 2 - 1"), ["4"])
        self.assertEqual(out("echo 7 / 2"), ["3.5"])

    def test_unary(self):
        self.assertEqual(out("a = -1\necho -a"), ["1"])
        self.assertEqual(out("echo +5"), ["5"])
        self.assertEqual(out("echo -(-5)"), ["5"])

    def test_string_concat(self):
        self.assertEqual(out('echo "a" & "b"'), ["ab"])
        self.assertEqual(out("echo 1 & 2"), ["12"])
        self.assertEqual(out('a = 1\necho a + 2 & "u"'), ["3u"])

    def test_comparisons_return_vbs_booleans(self):
        self.assertEqual(out("echo 1 < 2"), ["True"])
        self.assertEqual(out("t = (1 < 2)\necho t + 0"), ["-1"])
        self.assertEqual(out("t = (1 > 2)\necho t + 0"), ["0"])

    def test_boolean_literals(self):
        self.assertEqual(out("echo true"), ["True"])
        self.assertEqual(out("echo false"), ["False"])
        self.assertEqual(out("echo true + true"), ["-2"])

    def test_and_or_not(self):
        # VBS logical operators are bitwise on integers: -1 And 0 = 0, Not 0 = -1
        # 中文：VBS 的逻辑运算符对整数是按位的：-1 And 0 = 0，Not 0 = -1
        self.assertEqual(out("echo -1 AND 0"), ["0"])
        self.assertEqual(out("echo -1 OR 0"), ["-1"])
        self.assertEqual(out("echo NOT 0"), ["-1"])
        self.assertEqual(out("echo NOT (1 == 2)"), ["True"])

    def test_bitwise_logical_operators(self):
        self.assertEqual(out("echo 5 AND 3"), ["1"])
        self.assertEqual(out("echo 5 OR 3"), ["7"])
        self.assertEqual(out("echo 5 XOR 3"), ["6"])

    def test_vbs_arithmetic_operators(self):
        self.assertEqual(out("echo 7 MOD 2"), ["1"])
        self.assertEqual(out("echo -7 MOD 2"), ["-1"])
        self.assertEqual(out("echo 7 \\ 2"), ["3"])
        self.assertEqual(out("echo 2 ^ 10"), ["1024"])
        self.assertEqual(out("echo 2 ^ 3 ^ 2"), ["64"])
        self.assertEqual(out("echo -2 ^ 2"), ["-4"])

    def test_concat_binds_tighter_than_compare(self):
        # this used to be a syntax error (wrong precedence function for the right operand)
        # 中文：曾经这里会报语法错误（比较右值用错了优先级函数）
        self.assertEqual(out('echo "a" < "b" & "c"'), ["True"])

    def test_compare_number_with_string_raises_type_error(self):
        with self.assertRaises(OVPPTypeError):
            run('echo 1 < "b" & "c"')

    def test_string_escaped_quotes(self):
        self.assertEqual(run('echo "say ""hi"""'), 'say "hi"\n')

    def test_unicode_string(self):
        self.assertEqual(run('echo "中文测试"'), "中文测试\n")


class TestControlFlow(unittest.TestCase):
    def test_if_else(self):
        self.assertEqual(out('a = 2\nif a == 1 then\necho "A"\nelse\necho "B"\nend if'), ["B"])

    def test_if_without_else(self):
        # this used to raise UnboundLocalError
        # 中文：曾经这里是 UnboundLocalError
        self.assertEqual(out('a = 1\nif a == 1 then\necho "x"\nend if\necho "y"'), ["x", "y"])
        self.assertEqual(out('a = 9\nif a == 1 then\necho "x"\nend if\necho "y"'), ["y"])

    def test_elseif_chain_without_else(self):
        code = ('a = 3\nif a == 1 then\necho 1\n'
                'elseif a == 2 then\necho 2\n'
                'elseif a == 3 then\necho 3\nend if')
        self.assertEqual(out(code), ["3"])

    def test_multiple_elseif(self):
        for value, expected in ((1, "1"), (2, "2"), (3, "3"), (4, "9")):
            code = (f'a = {value}\nif a == 1 then\necho 1\n'
                    'elseif a == 2 then\necho 2\n'
                    'elseif a == 3 then\necho 3\n'
                    'else\necho 9\nend if')
            with self.subTest(value=value):
                self.assertEqual(out(code), [expected])

    def test_while(self):
        self.assertEqual(out("i = 0\nwhile i < 3\necho i\ni = i + 1\nwend"), ["0", "1", "2"])

    def test_do_while_and_until(self):
        self.assertEqual(out("i = 0\ndo while i < 3\necho i\ni = i + 1\nloop"), ["0", "1", "2"])
        self.assertEqual(out("i = 0\ndo until i == 3\necho i\ni = i + 1\nloop"), ["0", "1", "2"])

    def test_do_loop_post_condition(self):
        self.assertEqual(out("i = 0\ndo\ni = i + 1\nloop while i < 3\necho i"), ["3"])
        self.assertEqual(out("i = 0\ndo\ni = i + 1\nloop until i == 2\necho i"), ["2"])

    def test_do_loop_runs_at_least_once(self):
        self.assertEqual(out("i = 5\ndo\necho i\ni = i + 1\nloop while i < 3"), ["5"])

    def test_exit_do_inside_if(self):
        # this used to hang
        # 中文：曾经这里会卡死
        code = ("i = 0\ndo\ni = i + 1\nif i == 2 then\nexit do\nend if\nloop\necho i")
        self.assertEqual(out(code), ["2"])

    def test_for_loop(self):
        self.assertEqual(out("for i = 0 to 3\necho i\nnext"), ["0", "1", "2", "3"])

    def test_for_loop_with_step(self):
        self.assertEqual(out("for i = 0 to 10 step 5\necho i\nnext"), ["0", "5", "10"])
        self.assertEqual(out("for i = 3 to 1 step -1\necho i\nnext"), ["3", "2", "1"])

    def test_for_loop_does_not_run_when_empty(self):
        self.assertEqual(out("for i = 5 to 1\necho i\nnext"), [])

    def test_for_next_with_variable(self):
        self.assertEqual(out("for i = 1 to 2\necho i\nnext i"), ["1", "2"])

    def test_for_end_value_evaluated_once(self):
        code = "n = 3\nfor i = 1 to n\nn = 1\necho i\nnext"
        self.assertEqual(out(code), ["1", "2", "3"])

    def test_exit_for(self):
        code = "for i = 1 to 100\nif i == 4 then\nexit for\nend if\necho i\nnext"
        self.assertEqual(out(code), ["1", "2", "3"])

    def test_nested_loops(self):
        code = "for i = 0 to 1\nfor j = 0 to 1\necho i & j\nnext\nnext"
        self.assertEqual(out(code), ["00", "01", "10", "11"])

    def test_select_case(self):
        code = ('x = 4\nselect case x\ncase 1\necho "A"\n'
                'case 2, 4\necho "B"\ncase else\necho "D"\nend select')
        self.assertEqual(out(code), ["B"])

    def test_select_case_ranges_and_is(self):
        # Case 1, 3 To 5, Is <> 10  ->  everything except 10 matches
        # 中文：Case 1, 3 To 5, Is <> 10  →  除 10 以外全部命中
        for value, expected in ((1, "hit"), (3, "hit"), (5, "hit"),
                                (7, "hit"), (20, "hit"), (10, "other")):
            code = (f'x = {value}\nselect case x\n'
                    'case 1, 3 to 5, is <> 10\necho "hit"\n'
                    'case is >= 20\necho "big"\n'
                    'case else\necho "other"\nend select')
            with self.subTest(value=value):
                self.assertEqual(out(code), [expected])

    def test_select_case_without_cases(self):
        # this used to raise ValueError: None disallowed in statement list
        # 中文：曾经抛 ValueError: None disallowed in statement list
        self.assertEqual(out("x = 1\nselect case x\nend select\necho 9"), ["9"])

    def test_select_case_with_only_else(self):
        # this used to raise AttributeError
        # 中文：曾经抛 AttributeError
        self.assertEqual(out('x = 1\nselect case x\ncase else\necho "e"\nend select'), ["e"])

    def test_nested_select_case(self):
        code = ('x = 2\nselect case x\ncase 1\necho "a"\ncase else\n'
                'select case x\ncase 2\necho "b"\nend select\nend select')
        self.assertEqual(out(code), ["b"])

    def test_select_case_does_not_clobber_user_variable(self):
        code = ('_select_case_tmp = 99\nx = 5\nselect case x\n'
                'case 5\necho _select_case_tmp\nend select')
        self.assertEqual(out(code), ["99"])

    def test_for_bounds_are_expressions(self):
        self.assertEqual(out("n = 3\nfor i = 1 to n * 2\necho i\nnext"),
                         ["1", "2", "3", "4", "5", "6"])

    def test_deeply_nested_blocks(self):
        code = (
            "sum = 0\n"
            "for i = 1 to 2\n"                 # For
            "  j = 0\n"
            "  while j < 2\n"                  # While
            "    select case i\n"              # Select
            "    case 1\n"
            "      if j == 0 then\n"           # If
            "        sum = sum + 1\n"
            "      else\n"
            "        sum = sum + 10\n"
            "      end if\n"
            "    case else\n"
            "      sum = sum + 100\n"
            "    end select\n"
            "    j = j + 1\n"
            "  wend\n"
            "next\n"
            "echo sum"
        )
        # i=1: +1 +10 ; i=2: two case-else hits add +200
        # 中文：i=1: +1 +10 ；i=2: 两次 case else 共 +200
        self.assertEqual(out(code), ["211"])

    def test_exit_do_from_inside_select_case(self):
        code = ("i = 0\ndo\ni = i + 1\nselect case i\ncase 3\nexit do\n"
                "end select\nloop\necho i")
        self.assertEqual(out(code), ["3"])

    def test_empty_block_bodies_are_allowed(self):
        # empty If / loop bodies need pass in Python; check compile still succeeds
        # 中文：空的 If / 循环体在 Python 里必须补 pass，这里确认不会编译失败
        self.assertEqual(out('if 1 == 1 then\nend if\necho "ok"'), ["ok"])
        self.assertEqual(out('for i = 1 to 2\nnext\necho "ok"'), ["ok"])
        self.assertEqual(out('while false\nwend\necho "ok"'), ["ok"])
        self.assertEqual(out('select case 1\ncase 1\nend select\necho "ok"'), ["ok"])


class TestLexing(unittest.TestCase):
    def test_rem_comment(self):
        self.assertEqual(out("rem a comment\necho 1"), ["1"])

    def test_apostrophe_comment(self):
        self.assertEqual(out("' a comment\necho 1"), ["1"])

    def test_trailing_comments(self):
        self.assertEqual(out("a = 1 ' trailing\necho a"), ["1"])
        self.assertEqual(out("a = 1 rem trailing\necho a"), ["1"])

    def test_comment_only_program(self):
        self.assertEqual(out("rem nothing here"), [])

    def test_crlf_line_endings(self):
        self.assertEqual(out("a = 1\r\necho a\r\n"), ["1"])

    def test_cr_only_line_endings(self):
        self.assertEqual(out("a = 1\recho a\r"), ["1"])

    def test_blank_lines_and_tabs(self):
        self.assertEqual(out("\n\n\t a = 1\n\n\techo a\n\n"), ["1"])

    def test_case_insensitive_keywords(self):
        self.assertEqual(out('IF 1 == 1 THEN\n ECHO "yes"\nEND IF'), ["yes"])


class TestSyntaxErrors(unittest.TestCase):
    def assertSyntaxError(self, code, line=None):
        with self.assertRaises(OVPPSyntaxError) as ctx:
            run(code)
        err = ctx.exception
        if line is not None:
            self.assertEqual(err.line_number, line, str(err))
        return err

    def test_missing_end_if(self):
        # this used to loop forever
        # 中文：曾经会无限循环（卡死）
        err = self.assertSyntaxError('if 1 == 1 then\necho 1\n')
        self.assertIn("End If", err.message)

    def test_missing_end_select(self):
        self.assertSyntaxError("select case 1\ncase 1\necho 1\n")

    def test_missing_next(self):
        self.assertSyntaxError("for i = 1 to 3\necho i\n")

    def test_missing_wend(self):
        self.assertSyntaxError("while 1 == 1\necho 1\n")

    def test_missing_loop(self):
        self.assertSyntaxError("i = 0\ndo\ni = i + 1\n")

    def test_stray_end_if(self):
        self.assertSyntaxError("end if")

    def test_single_line_if_is_supported(self):
        # single-line If has been supported since v0.3
        # 中文：v0.3 起支持单行 If
        self.assertEqual(out('if 1 == 1 then echo "x"'), ["x"])

    def test_exit_sub_unsupported(self):
        self.assertSyntaxError("exit sub")

    def test_unknown_statement(self):
        # 'print' is neither a procedure nor a keyword: it now reports procedure-not-defined and suggests Echo
        # 中文：'print' 不是过程，也不是语法关键字：现在会给出"过程未定义"并提示用 Echo
        from OpenVBSPP.exceptions import OVPPProcedureNotDefinedError
        with self.assertRaises(OVPPProcedureNotDefinedError) as ctx:
            run("print 1")
        self.assertIn("Echo", ctx.exception.message)

    def test_error_carries_line_number_and_source(self):
        err = self.assertSyntaxError("a = 1\nb = 2\nx = = 3\n", line=3)
        self.assertEqual(err.source_line, "x = = 3")

    def test_unterminated_string(self):
        self.assertSyntaxError('echo "abc\n')

    def test_float_literal_is_supported(self):
        self.assertEqual(out("echo 3.5"), ["3.5"])
        self.assertEqual(out("echo 2.0"), ["2"])

    def test_colon_statement_separator(self):
        # colon-separated statements have been supported since v0.3
        # 中文：v0.3 起支持用冒号分隔多条语句
        self.assertEqual(out("a = 1 : b = 2"), [])
        self.assertEqual(out("a = 1 : echo a"), ["1"])

    def test_missing_then(self):
        self.assertSyntaxError("if 1 == 1\necho 1\nend if")


class TestRuntimeErrors(unittest.TestCase):
    def test_undeclared_variable_reads_as_empty(self):
        # Without Option Explicit VBScript creates a variable the first time it
        # is read, and its value is Empty.
        # 中文：未开 Option Explicit 时，VBS 读到未声明的名字会自动创建 Empty 变量。
        self.assertEqual(out("a = 1\nb = 2\necho TypeName(zzz)\n"), ["Empty"])
        self.assertEqual(out('a = 1\nb = 2\necho zzz & "!"\n'), ["!"])

    def test_undeclared_variable_is_an_error_under_option_explicit(self):
        # With Option Explicit the name is still refused, with its location
        # 中文：开了 Option Explicit 仍然报错，并且带位置信息。
        with self.assertRaises(OVPPVariableNotDefinedError) as ctx:
            run("Option Explicit\nDim a\nDim b\na = 1\nb = 2\necho zzz\n")
        err = ctx.exception
        self.assertEqual(err.line_number, 6)
        self.assertEqual(err.source_line, "echo zzz")
        self.assertIn("zzz", err.message)

    def test_division_by_zero(self):
        with self.assertRaises(OVPPDivisionByZeroError) as ctx:
            run("echo 1 / 0\n")
        self.assertEqual(ctx.exception.line_number, 1)

    def test_error_message_mentions_location(self):
        with self.assertRaises(OVPPDivisionByZeroError) as ctx:
            run("echo 1 / 0\n")
        text = str(ctx.exception)
        self.assertIn("Line 1", text)
        self.assertIn("echo 1 / 0", text)


class TestCompiler(unittest.TestCase):
    def test_no_debug_output_by_default(self):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            vbs_compile("echo 1\n")
        self.assertEqual(buf.getvalue(), "")

    def test_verbose_prints_pipeline(self):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            vbs_compile("echo 1\n", verbose=True)
        text = buf.getvalue()
        self.assertIn("TOKENS:", text)
        self.assertIn("VBS AST:", text)
        self.assertIn("Echo(1)", text)

    def test_compiled_filename_is_used(self):
        code_obj = vbs_compile("echo 1\n", filename="demo.vbs")
        self.assertEqual(code_obj.co_filename, "demo.vbs")

    def test_default_filename(self):
        self.assertEqual(vbs_compile("echo 1\n").co_filename, "<OpenVBS++>")

    def test_invalid_mode(self):
        with self.assertRaises(ValueError):
            vbs_compile("echo 1\n", mode="nonsense")

    def test_save_requires_filename(self):
        with self.assertRaises(ValueError):
            vbs_compile("echo 1\n", mode="save")

    def test_save_writes_pyc_and_ast(self):
        # this used to raise FileNotFoundError because ./__pycache__ was hard-coded
        # 中文：曾经因为写死 ./__pycache__ 而 FileNotFoundError
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "nested" / "cache"
            vbs_compile("echo 1\n", mode="save", filename="demo.vbs",
                        save_dir=str(out_dir))
            pyc = out_dir / "demo.pyc"
            ast_txt = out_dir / "demo.ast.txt"
            self.assertTrue(pyc.is_file())
            self.assertTrue(ast_txt.is_file())

            data = pyc.read_bytes()
            import importlib.util
            self.assertEqual(data[:4], importlib.util.MAGIC_NUMBER)
            code_obj = marshal.loads(data[16:])
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                exec(code_obj, {})
            self.assertEqual(buf.getvalue(), "1\n")

    def test_empty_program(self):
        self.assertEqual(run(""), "")


class TestBooleanRuntime(unittest.TestCase):
    def test_values(self):
        self.assertEqual(int(vbool.TRUE), -1)
        self.assertEqual(int(vbool.FALSE), 0)

    def test_display(self):
        self.assertEqual(str(vbool.TRUE), "True")
        self.assertEqual(str(vbool.FALSE), "False")
        self.assertEqual(repr(vbool.TRUE), "True")

    def test_arithmetic(self):
        self.assertEqual(vbool.TRUE + 0, -1)
        self.assertEqual(vbool.TRUE + vbool.TRUE, -2)

    def test_helpers(self):
        self.assertEqual(int(vbool.to_bool([])), 0)
        self.assertEqual(int(vbool.to_bool([1])), -1)
        self.assertEqual(int(vbool.ovpp_and(-1, 0)), 0)
        self.assertEqual(int(vbool.ovpp_or(-1, 0)), -1)
        self.assertEqual(int(vbool.ovpp_not(0)), -1)


class TestCli(unittest.TestCase):
    def _main(self, argv):
        from OpenVBSPP.cli import main
        out_buf, err_buf = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out_buf), contextlib.redirect_stderr(err_buf):
            code = main(argv)
        return code, out_buf.getvalue(), err_buf.getvalue()

    def test_runs_file(self):
        code, stdout, _ = self._main([str(EXAMPLES / "control1.vbs")])
        self.assertEqual(code, 0)
        self.assertEqual(stdout.split(), ["15"])

    def test_demo_without_arguments(self):
        code, stdout, _ = self._main([])
        self.assertEqual(code, 0)
        self.assertIn("Hello World", stdout)

    def test_missing_file(self):
        code, _, stderr = self._main(["does_not_exist.vbs"])
        self.assertEqual(code, 1)
        self.assertIn("File not found", stderr)

    def test_syntax_error_reports_and_exits_1(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.vbs"
            path.write_text("if 1 == 1 then\n", encoding="utf-8")
            code, _, stderr = self._main([str(path)])
        self.assertEqual(code, 1)
        self.assertIn("Syntax error", stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
