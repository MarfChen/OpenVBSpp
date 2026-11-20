"""Multi-file: <script src> merging and <namespace src alias> namespaces.

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
多文件：<script src> 合并 与 <namespace src alias> 命名空间。
"""

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import OVPPError        # noqa: E402
from OpenVBSPP.runtime import vbs_run             # noqa: E402


LIB = '''Public Function Double(n)
    Double = n * 2
End Function

Public Function Greet(name)
    Greet = "hi " & name
End Function

Private Function Secret()
    Secret = 42
End Function
'''


class ModuleCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def write(self, name, text):
        path = self.dir / name
        path.write_text(text, encoding="utf-8")
        return path

    def run_file(self, path):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            vbs_run(path.read_text(encoding="utf-8"), filename=str(path))
        return buf.getvalue().splitlines()


class TestScriptInclude(ModuleCase):
    def test_merge_brings_procedures_into_scope(self):
        self.write("lib.vbs", LIB)
        main = self.write("main.vbs",
                          '<script src="lib.vbs">\n'
                          'echo Double(21)\necho Greet("ann")\n')
        self.assertEqual(self.run_file(main), ["42", "hi ann"])

    def test_missing_file_is_a_compile_error(self):
        main = self.write("main.vbs", '<script src="nope.vbs">\n')
        with self.assertRaises(OVPPError):
            self.run_file(main)

    def test_circular_dependency_is_reported(self):
        self.write("a.vbs", '<script src="b.vbs">\n')
        self.write("b.vbs", '<script src="a.vbs">\n')
        main = self.dir / "a.vbs"
        with self.assertRaises(OVPPError):
            self.run_file(main)

    def test_error_line_comes_from_included_file(self):
        lib = self.write("bad_lib.vbs", 'echo "ok"\necho 1 / 0\n')
        main = self.write("main.vbs", f'<script src="{lib.name}">\n')
        with self.assertRaises(OVPPError) as ctx:
            self.run_file(main)
        self.assertEqual(ctx.exception.line_number, 2)


class TestNamespace(ModuleCase):
    def test_namespace_prefix_call(self):
        self.write("lib.vbs", LIB)
        main = self.write("main.vbs",
                          '<namespace src="lib.vbs" alias="u">\n'
                          'echo u.Double(21)\necho u.Greet("bob")\n')
        self.assertEqual(self.run_file(main), ["42", "hi bob"])

    def test_alias_defaults_to_file_stem(self):
        self.write("utils.vbs", LIB)
        main = self.write("main.vbs", '<namespace src="utils.vbs">\n'
                                      'echo utils.Double(5)\n')
        self.assertEqual(self.run_file(main), ["10"])

    def test_nickyname_attribute_is_accepted(self):
        self.write("lib.vbs", LIB)
        main = self.write("main.vbs",
                          '<namespace src="lib.vbs" nickyname="tools">\n'
                          'echo tools.Double(6)\n')
        self.assertEqual(self.run_file(main), ["12"])

    def test_private_procedure_is_not_exported(self):
        self.write("lib.vbs", LIB)
        main = self.write("main.vbs",
                          '<namespace src="lib.vbs" alias="u">\necho u.Secret()\n')
        with self.assertRaises(OVPPError):
            self.run_file(main)

    def test_case_insensitive_member(self):
        self.write("lib.vbs", LIB)
        main = self.write("main.vbs",
                          '<namespace src="lib.vbs" alias="u">\necho U.DOUBLE(4)\n')
        self.assertEqual(self.run_file(main), ["8"])


class TestModuleIsolation(ModuleCase):
    def test_same_name_in_two_namespaces(self):
        self.write("a.vbs", 'Public Function Double(n)\n    Double = n * 2\nEnd Function\n')
        self.write("b.vbs", 'Public Function Double(n)\n    Double = n * 100\nEnd Function\n')
        main = self.write("main.vbs",
                          '<namespace src="a.vbs" alias="a">\n'
                          '<namespace src="b.vbs" alias="b">\n'
                          'echo a.Double(3)\necho b.Double(3)\n')
        self.assertEqual(self.run_file(main), ["6", "300"])

    def test_library_can_use_another_namespace(self):
        self.write("inner.vbs", 'Public Function Twice(n)\n    Twice = n * 2\nEnd Function\n')
        self.write("outer.vbs",
                   '<namespace src="inner.vbs" alias="deep">\n'
                   'Public Function Quad(n)\n'
                   '    Quad = deep.Twice(deep.Twice(n))\nEnd Function\n')
        main = self.write("main.vbs",
                          '<namespace src="outer.vbs" alias="o">\necho o.Quad(3)\n')
        self.assertEqual(self.run_file(main), ["12"])

    def test_script_merged_procedures_are_exported(self):
        self.write("helper.vbs", 'Public Function Helper(n)\n    Helper = n + 1\nEnd Function\n')
        self.write("lib.vbs",
                   '<script src="helper.vbs">\n'
                   'Public Function UseIt(n)\n    UseIt = Helper(n)\nEnd Function\n')
        main = self.write("main.vbs",
                          '<namespace src="lib.vbs" alias="L">\necho L.UseIt(1)\n')
        self.assertEqual(self.run_file(main), ["2"])


class TestPlainFileUnaffected(unittest.TestCase):
    def test_no_directives_still_single_unit(self):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            vbs_run("echo 1\n", filename="<plain>")
        self.assertEqual(buf.getvalue().strip(), "1")


if __name__ == "__main__":
    unittest.main(verbosity=2)
