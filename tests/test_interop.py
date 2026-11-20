"""Python interop: `#python ... #endpython` and `COP`. Since 0.9.20 both need an explicit `<script src="python-embed" />`.

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

中文：Python 互操作：`#python ... #endpython` 与 `COP`。

0.9.20 起这些能力必须显式导入：脚本里写 `<script src="python-embed" />`
（以前靠 `Option COP`，该语句已移除）。

运行：
    python -m unittest tests.test_interop -v
"""

import contextlib
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import OVPPSyntaxError        # noqa: E402
from OpenVBSPP.ovpprt.cop import COP, reset             # noqa: E402
from OpenVBSPP.runtime import vbs_run                   # noqa: E402


PY_IMPORT = '<script src="python-embed" />\n'


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<interop>")
    return buf.getvalue().splitlines()


class TestPythonEmbed(unittest.TestCase):
    def test_python_block_runs(self):
        code = PY_IMPORT + 'a = 10\n#python\nb = a * 2\n#endpython\necho b'
        self.assertEqual(out(code), ["20"])

    def test_python_block_can_define_function(self):
        code = (PY_IMPORT +
                '#python\n'
                'def triple(x):\n'
                '    return x * 3\n'
                'tripled = triple(4)\n'
                '#endpython\n'
                'echo tripled')
        self.assertEqual(out(code), ["12"])

    def test_line_numbers_stay_correct(self):
        # line numbers after a #python block must still point at the real lines
        # 中文：#python 块之后的行号必须仍然指向真实行
        code = PY_IMPORT + 'a = 1\n#python\nx = 2\n#endpython\necho 1 / 0'
        with self.assertRaises(Exception) as ctx:
            out(code)
        self.assertEqual(getattr(ctx.exception, "line_number", None), 6)

    def test_unterminated_block_is_ignored(self):
        # without #endpython the block is plain VBS, so a stray # is a syntax error
        # 中文：没有 #endpython 时，按普通 VBS 处理（# 会报语法错）
        with self.assertRaises(OVPPSyntaxError):
            out(PY_IMPORT + 'a = 1\n#python\nx = 2')

    def test_python_block_needs_explicit_import(self):
        with self.assertRaises(OVPPSyntaxError) as ctx:
            out('#python\nx = 2\n#endpython')
        self.assertIn("python-embed", str(ctx.exception))

    def test_option_cop_reports_removal(self):
        with self.assertRaises(OVPPSyntaxError) as ctx:
            out('Option COP\necho 1')
        self.assertIn("0.9.20", str(ctx.exception))


class TestCop(unittest.TestCase):
    def setUp(self):
        reset()

    def tearDown(self):
        reset()

    def test_cop_is_shared_with_python(self):
        code = (PY_IMPORT +
                'COP("seed") = 5\n'
                '#python\n'
                'from OpenVBSPP.ovpprt.cop import COP as PYCOP\n'
                'PYCOP["frompython"] = PYCOP["seed"] * 3\n'
                '#endpython\n'
                'echo COP("frompython")')
        self.assertEqual(out(code), ["15"])

    def test_python_object_visible_in_vbs(self):
        COP["obj"] = {"k": 1}
        code = (PY_IMPORT +
                'echo TypeName(COP("obj"))')
        self.assertEqual(out(code), ["dict"])

    def test_cop_requires_import(self):
        with self.assertRaises(Exception) as ctx:
            out('COP("x") = 1')
        self.assertIn("python-embed", str(ctx.exception))


if __name__ == "__main__":
    unittest.main(verbosity=2)
