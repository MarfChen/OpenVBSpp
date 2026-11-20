"""v0.9 additions: Format / date functions / FileSystemObject.

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
v0.9 新增能力：Format / 日期函数 / FileSystemObject。
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


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<v09>")
    return buf.getvalue().splitlines()


class TestFormat(unittest.TestCase):
    def test_number_patterns(self):
        self.assertEqual(out('echo Format(1234.5, "#,##0.00")'), ["1,234.50"])
        self.assertEqual(out('echo Format(0.256, "0.0%")'), ["25.6%"])
        self.assertEqual(out('echo Format(7, "0000")'), ["0007"])

    def test_date_patterns(self):
        code = 'd = DateSerial(2027, 1, 2)\necho Format(d, "yyyy-mm-dd")'
        self.assertEqual(out(code), ["2027-01-02"])

    def test_long_date_pattern(self):
        code = 'd = DateSerial(2027, 1, 2)\necho Format(d, "dddd, mmmm d, yyyy")'
        self.assertEqual(out(code), ["Saturday, January 2, 2027"])

    def test_no_format_returns_text(self):
        self.assertEqual(out("echo Format(3.0)"), ["3"])


class TestDateFunctions(unittest.TestCase):
    def test_date_serial_normalizes(self):
        self.assertEqual(out("echo Format(DateSerial(2027, 13, 1), \"yyyy-mm-dd\")"),
                         ["2028-01-01"])
        self.assertEqual(out("echo Format(DateSerial(2027, 1, 32), \"yyyy-mm-dd\")"),
                         ["2027-02-01"])

    def test_date_part(self):
        code = ('echo DatePart("yyyy", "2027-03-05")\n'
                'echo DatePart("q", "2027-03-05")\n'
                'echo DatePart("m", "2027-03-05")\n'
                'echo DatePart("d", "2027-03-05")')
        self.assertEqual(out(code), ["2027", "1", "3", "5"])

    def test_month_and_weekday_names(self):
        code = "echo MonthName(3)\necho MonthName(3, True)\necho WeekdayName(1)"
        self.assertEqual(out(code), ["March", "Mar", "Sunday"])

    def test_date_and_time_value(self):
        self.assertEqual(out('echo Format(DateValue("2027-05-06 07:08:09"), "yyyy-mm-dd")'),
                         ["2027-05-06"])
        self.assertEqual(out('echo Format(TimeValue("2027-05-06 07:08:09"), "hh:nn:ss")'),
                         ["07:08:09"])


class TestFileSystemObject(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = self.tmp.name.replace("\\", "/")

    def tearDown(self):
        self.tmp.cleanup()

    def test_write_then_read(self):
        code = (f'Set fso = CreateObject("Scripting.FileSystemObject")\n'
                f'p = "{self.dir}/a.txt"\n'
                f'Set f = fso.CreateTextFile(p, True)\n'
                f'f.WriteLine "one"\nf.WriteLine "two"\nf.Close\n'
                f'echo fso.FileExists(p)\n'
                f'Set r = fso.OpenTextFile(p, 1)\n'
                f'echo r.ReadLine\necho r.ReadLine\necho r.AtEndOfStream\nr.Close')
        self.assertEqual(out(code), ["True", "one", "two", "True"])

    def test_path_helpers(self):
        code = ('Set fso = CreateObject("Scripting.FileSystemObject")\n'
                'echo fso.GetFileName("C:\\data\\r.txt")\n'
                'echo fso.GetBaseName("C:\\data\\r.txt")\n'
                'echo fso.GetExtensionName("C:\\data\\r.txt")')
        self.assertEqual(out(code), ["r.txt", "r", "txt"])

    def test_folder_operations(self):
        target = f"{self.dir}/sub"
        code = (f'Set fso = CreateObject("Scripting.FileSystemObject")\n'
                f'fso.CreateFolder "{target}"\n'
                f'echo fso.FolderExists("{target}")\n'
                f'fso.DeleteFolder "{target}"\n'
                f'echo fso.FolderExists("{target}")')
        self.assertEqual(out(code), ["True", "False"])

    def test_read_missing_file_raises(self):
        code = ('Set fso = CreateObject("Scripting.FileSystemObject")\n'
                'Set r = fso.OpenTextFile("nope-does-not-exist.txt", 1)')
        with self.assertRaises(OVPPError):
            out(code)

    def test_method_names_are_case_insensitive(self):
        code = (f'Set fso = CreateObject("Scripting.FileSystemObject")\n'
                f'echo fso.fileexists("nope-does-not-exist.txt")')
        self.assertEqual(out(code), ["False"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
