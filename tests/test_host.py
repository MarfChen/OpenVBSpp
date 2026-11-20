"""The `WScript` host object.

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

`WScript` 宿主对象。

老脚本习惯用 `WScript.Echo` 输出、用 `WScript.Arguments` 读命令行，这些用例把宿主对象的
表面钉在 Windows Script Host 的行为上：名称与版本、脚本文件名、零基参数、具名参数、
标准流、`Quit`、`Sleep`，以及 `WScript` 是只读的。
"""

import contextlib
import io
import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import OVPPSyntaxError      # noqa: E402
from OpenVBSPP.ovpprt import host                     # noqa: E402
from OpenVBSPP.runtime import vbs_run                 # noqa: E402


def out(code, filename="<host>", args=()):
    """Run a snippet and return its stdout lines.

    中文：运行一段代码，返回 stdout 的每一行。
    """
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename=filename, args=list(args))
    return buf.getvalue().splitlines()


class TestHostIdentity(unittest.TestCase):
    """Name, version and `TypeName`.

    中文：名称、版本与 `TypeName`。
    """

    def test_type_name_is_object(self):
        self.assertEqual(out("WScript.Echo TypeName(WScript)"), ["Object"])

    def test_name_and_version(self):
        self.assertEqual(
            out('WScript.Echo WScript.Name & " " & WScript.Version'),
            ["Windows Script Host 5.812"])

    def test_script_name_is_the_base_name(self):
        path = os.path.join("some", "dir", "legacy.vbs")
        self.assertEqual(
            out("WScript.Echo WScript.ScriptName", filename=path),
            ["legacy.vbs"])

    def test_script_full_name_is_absolute(self):
        lines = out('WScript.Echo WScript.ScriptFullName',
                    filename=os.path.join("some", "dir", "legacy.vbs"))
        self.assertTrue(os.path.isabs(lines[0]))
        self.assertTrue(lines[0].endswith("legacy.vbs"))

    def test_script_name_is_empty_without_a_file(self):
        self.assertEqual(out("WScript.Echo WScript.ScriptName", filename=None),
                         [""])

    def test_full_name_is_the_interpreter(self):
        self.assertEqual(out("WScript.Echo WScript.FullName"),
                         [sys.executable])

    def test_interactive_is_false_when_output_is_redirected(self):
        self.assertEqual(out("WScript.Echo WScript.Interactive"), ["False"])


class TestArguments(unittest.TestCase):
    """`WScript.Arguments`, including the zero-based index and named form.

    中文：`WScript.Arguments`，包括零基下标与具名参数。
    """

    def test_count_and_items(self):
        lines = out("WScript.Echo WScript.Arguments.Count\n"
                    "WScript.Echo WScript.Arguments(0)\n"
                    "WScript.Echo WScript.Arguments(1)",
                    args=["alpha", "beta"])
        self.assertEqual(lines, ["2", "alpha", "beta"])

    def test_item_method(self):
        self.assertEqual(
            out('WScript.Echo WScript.Arguments.Item(1)', args=["a", "b"]),
            ["b"])

    def test_for_each_walks_the_arguments(self):
        lines = out("For Each value In WScript.Arguments\n"
                    "    WScript.Echo value\n"
                    "Next", args=["one", "two", "three"])
        self.assertEqual(lines, ["one", "two", "three"])

    def test_missing_argument_is_subscript_out_of_range(self):
        # On Error Resume Next turns the failure into Err.Number 9.
        # 中文：On Error Resume Next 会把失败变成 Err.Number 9。
        self.assertEqual(
            out("On Error Resume Next\n"
                "value = WScript.Arguments(9)\n"
                "WScript.Echo Err.Number"), ["9"])

    def test_named_arguments(self):
        lines = out('WScript.Echo WScript.Arguments.Named.Exists("mode")\n'
                    'WScript.Echo WScript.Arguments.Named("mode")',
                    args=["plain", "/mode:fast"])
        self.assertEqual(lines, ["True", "fast"])

    def test_named_arguments_accept_equals(self):
        self.assertEqual(
            out('WScript.Echo WScript.Arguments.Named("level")',
                args=["/level=9"]), ["9"])

    def test_named_arguments_count(self):
        self.assertEqual(
            out("WScript.Echo WScript.Arguments.Named.Count",
                args=["plain", "/a:1", "/b:2"]), ["2"])


class TestStreams(unittest.TestCase):
    """`WScript.Echo` and the three `TextStream`s.

    中文：`WScript.Echo` 与三个 `TextStream`。
    """

    def test_echo_separates_values_with_spaces(self):
        self.assertEqual(out('WScript.Echo "a", "b", 3'), ["a b 3"])

    def test_stdout_write_line(self):
        self.assertEqual(out('WScript.StdOut.WriteLine "hello"'), ["hello"])

    def test_stdout_write_does_not_break_the_line(self):
        self.assertEqual(
            out('WScript.StdOut.Write "a"\nWScript.StdOut.Write "b"'), ["ab"])

    def test_stderr_stream_is_a_text_stream(self):
        err = io.StringIO()
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(err):
            vbs_run('WScript.StdErr.WriteLine "boom"')
        self.assertEqual(err.getvalue(), "boom\n")

    def test_type_name_of_streams(self):
        self.assertEqual(
            out("WScript.Echo TypeName(WScript.StdOut)"), ["TextStream"])

    def test_read_line_from_stdin(self):
        stream = host._TextStream(io.StringIO("first\nsecond\n"))
        self.assertEqual(stream.ReadLine(), "first")
        self.assertEqual(stream.ReadLine(), "second")
        with self.assertRaises(Exception):
            stream.ReadLine()

    def test_at_end_of_stream(self):
        stream = host._TextStream(io.StringIO(""))
        self.assertTrue(stream.AtEndOfStream)


class TestControl(unittest.TestCase):
    """`Quit`, `Sleep`, `Timeout` and read-only `WScript`.

    中文：`Quit`、`Sleep`、`Timeout` 与只读的 `WScript`。
    """

    def test_quit_raises_system_exit(self):
        with self.assertRaises(SystemExit) as caught:
            out("WScript.Quit 3")
        self.assertEqual(caught.exception.code, 3)

    def test_sleep_returns(self):
        self.assertEqual(out("WScript.Sleep 5\nWScript.Echo \"done\""),
                         ["done"])

    def test_timeout_is_writable(self):
        self.assertEqual(
            out("WScript.Timeout = 5\nWScript.Echo WScript.Timeout"), ["5"])

    def test_timeout_defaults_to_zero(self):
        self.assertEqual(out("WScript.Echo WScript.Timeout"), ["0"])

    def test_wscript_is_read_only(self):
        with self.assertRaises(OVPPSyntaxError):
            out("WScript = 1")

    def test_unknown_member_is_an_error(self):
        with self.assertRaises(Exception):
            out("WScript.NoSuchMember")


if __name__ == "__main__":
    unittest.main()
