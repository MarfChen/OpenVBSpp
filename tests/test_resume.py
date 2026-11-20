"""Resume / Resume Next, implemented as a state machine.

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
Resume / Resume Next（状态机实现）。
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
        vbs_run(code, filename="<resume>")
    return buf.getvalue().splitlines()


class TestResumeNext(unittest.TestCase):
    def test_continues_after_failing_statement(self):
        code = ('Sub F()\n On Error GoTo h\n echo "a"\n x = 1 / 0\n'
                ' echo "b"\n Exit Sub\nh:\n echo "err " & Err.Number\n'
                ' Resume Next\nEnd Sub\nF')
        self.assertEqual(out(code), ["a", "err 11", "b"])

    def test_error_state_survives_resume(self):
        code = ('Sub F()\n On Error GoTo h\n Err.Raise 53\n echo "next"\n'
                ' Exit Sub\nh:\n Resume Next\nEnd Sub\nF\n'
                "echo Err.Number")
        self.assertEqual(out(code), ["next", "53"])


class TestResume(unittest.TestCase):
    def test_retry_reexecutes_the_statement(self):
        code = ('Sub F()\n Dim d\n d = 0\n On Error GoTo h\n'
                ' echo 10 / d\n echo "after"\n Exit Sub\n'
                'h:\n d = 2\n If d = 2 Then Resume\nEnd Sub\nF')
        self.assertEqual(out(code), ["5", "after"])

    def test_retry_inside_loop(self):
        code = ('Sub F()\n Dim d, n\n d = 0\n On Error GoTo h\n'
                ' echo 6 / d\n echo "done"\n Exit Sub\n'
                'h:\n n = n + 1\n d = n\n If n < 3 Then Resume\n'
                ' Resume Next\nEnd Sub\nF')
        self.assertEqual(out(code), ["6", "done"])


class TestResumeErrors(unittest.TestCase):
    def test_resume_outside_handler_is_a_compile_error(self):
        with self.assertRaises(OVPPError):
            out("Sub F()\n Resume Next\nEnd Sub\nF")

    def test_resume_with_label_is_not_supported_yet(self):
        code = ('Sub F()\n On Error GoTo h\n Err.Raise 5\nh:\n'
                ' Resume cleanup\ncleanup:\nEnd Sub\nF')
        with self.assertRaises(OVPPError):
            out(code)


class TestLabelHandlerWithoutResume(unittest.TestCase):
    def test_handler_still_works(self):
        code = ('Sub F()\n On Error GoTo h\n x = 1 / 0\n Exit Sub\n'
                'h:\n echo "handled " & Err.Number\nEnd Sub\nF')
        self.assertEqual(out(code), ["handled 11"])

    def test_goto_label(self):
        code = ('Sub F()\n echo "a"\n GoTo done\n echo "b"\n'
                'done:\n echo "c"\nEnd Sub\nF')
        self.assertEqual(out(code), ["a", "c"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
