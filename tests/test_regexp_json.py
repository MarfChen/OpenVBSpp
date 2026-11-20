"""v0.9.2: AddressOf (procedure values), the RegExp object, Json encode/decode.

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
v0.9.2：AddressOf（过程值）、RegExp 对象、Json 编解码。
"""

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
        vbs_run(code, filename="<v092>")
    return buf.getvalue().splitlines()


class TestAddressOf(unittest.TestCase):
    def test_call_through_variable(self):
        code = ("Function Double(n)\n Double = n * 2\nEnd Function\n"
                "f = AddressOf Double\necho f(21)")
        self.assertEqual(out(code), ["42"])

    def test_byval_arguments(self):
        code = ("Function Add(a, b)\n Add = a + b\nEnd Function\n"
                "g = AddressOf Add\necho g(3, 4)")
        self.assertEqual(out(code), ["7"])

    def test_delegate_passed_as_argument(self):
        code = ("Function Twice(fn, v)\n Twice = fn(v) * 2\nEnd Function\n"
                "Function Inc(n)\n Inc = n + 1\nEnd Function\n"
                "h = AddressOf Inc\necho Twice(h, 10)")
        self.assertEqual(out(code), ["22"])

    def test_unknown_procedure(self):
        with self.assertRaises(OVPPError):
            out("x = AddressOf NoSuchProc")


class TestRegExp(unittest.TestCase):
    PRELUDE = "Set re = New RegExp\n"

    def test_pattern_and_test(self):
        code = (self.PRELUDE + 're.Pattern = "\\d{3}-\\d{4}"\n'
                'echo re.Test("call 555-1234")\necho re.Test("nope")')
        self.assertEqual(out(code), ["True", "False"])

    def test_global_replace(self):
        code = self.PRELUDE + 're.Pattern = "a"\nre.Global = True\n' \
                              'echo re.Replace("banana", "-")'
        self.assertEqual(out(code), ["b-n-n-"])

    def test_non_global_replace_only_first(self):
        code = self.PRELUDE + 're.Pattern = "a"\necho re.Replace("banana", "-")'
        self.assertEqual(out(code), ["b-nana"])

    def test_execute_collection(self):
        code = (self.PRELUDE + 're.Pattern = "\\d+"\nre.Global = True\n'
                'Set ms = re.Execute("a1 b22")\n'
                'echo ms.Count\n'
                'For Each m In ms\n echo m.Value & "@" & m.FirstIndex\nNext')
        self.assertEqual(out(code), ["2", "1@1", "22@4"])

    def test_ignorecase(self):
        code = self.PRELUDE + 're.Pattern = "hello"\nre.IgnoreCase = True\n' \
                              'echo re.Test("HELLO")'
        self.assertEqual(out(code), ["True"])

    def test_createobject_progid(self):
        code = ('Set re = CreateObject("VBScript.RegExp")\n'
                're.Pattern = "x"\necho re.Test("axb")')
        self.assertEqual(out(code), ["True"])

    def test_invalid_pattern_reports_error(self):
        code = self.PRELUDE + 're.Pattern = "(" \necho re.Test("x")'
        with self.assertRaises(OVPPError):
            out(code)


class TestJson(unittest.TestCase):
    def test_parse_object(self):
        code = 'Set d = Json.Parse("{""name"": ""ann""}")\necho d("name")'
        self.assertEqual(out(code), ["ann"])

    def test_parse_array_and_numbers(self):
        code = 'a = Json.Parse("[1, 2, 3]")\necho UBound(a) & ":" & (a(0) + a(2))'
        self.assertEqual(out(code), ["2:4"])

    def test_parse_nested(self):
        code = ('Set d = Json.Parse("{""list"": [{""v"": 5}]}")\n'
                'items = d("list")\n'
                'Set first = items(0)\necho first("v")')
        self.assertEqual(out(code), ["5"])

    def test_stringify_roundtrip(self):
        code = 'text = Json.Stringify(Array(1, "two", True))\necho text'
        self.assertEqual(out(code), ['[1, "two", true]'])

    def test_stringify_dictionary(self):
        code = ('Set d = CreateObject("Scripting.Dictionary")\n'
                'd.Add "a", 1\necho Json.Stringify(d)')
        self.assertEqual(out(code), ['{"a": 1}'])

    def test_invalid_json_reports_error(self):
        with self.assertRaises(OVPPError):
            out('Set d = Json.Parse("{bad}")')


if __name__ == "__main__":
    unittest.main(verbosity=2)
