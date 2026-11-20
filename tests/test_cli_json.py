"""`openvbs --json check|run`: machine-readable output.

中文：`openvbs --json check|run`：机器可读输出。"""

import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.cli import main as cli_main       # noqa: E402


def run_cli(argv):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = cli_main(argv)
    return code, out.getvalue().strip(), err.getvalue().strip()


class JsonCliCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def write(self, name, text):
        path = self.dir / name
        path.write_text(text, encoding="utf-8")
        return str(path)


class TestJsonCheck(JsonCliCase):
    def test_valid_script(self):
        path = self.write("ok.vbs", "echo 1\n")
        code, out, _ = run_cli(["--json", "check", path])
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out), {"ok": True, "diagnostics": []})

    def test_type_error_is_reported(self):
        path = self.write("bad.vbs", 'Option Typing\nDim n -> Integer\nn = "x"\n')
        code, out, _ = run_cli(["--json", "check", path])
        self.assertEqual(code, 1)
        payload = json.loads(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["diagnostics"][0]["line"], 3)
        self.assertIn("Type mismatch", payload["diagnostics"][0]["message"])

    def test_syntax_error_is_reported(self):
        path = self.write("bad2.vbs", "If True Then\n")
        code, out, _ = run_cli(["--json", "check", path])
        self.assertEqual(code, 1)
        self.assertTrue(json.loads(out)["diagnostics"])

    def test_missing_file(self):
        code, out, _ = run_cli(["--json", "check", str(self.dir / "nope.vbs")])
        self.assertEqual(code, 1)
        self.assertEqual(json.loads(out)["error"]["kind"], "FileNotFound")

    def test_flag_after_subcommand(self):
        path = self.write("ok.vbs", "echo 1\n")
        code, out, _ = run_cli(["check", "--json", path])
        self.assertEqual(code, 0)
        self.assertTrue(json.loads(out)["ok"])


class TestJsonRun(JsonCliCase):
    def test_output_is_captured(self):
        path = self.write("run.vbs", "echo 6 * 7\n")
        code, out, _ = run_cli(["--json", "run", path])
        self.assertEqual(code, 0)
        payload = json.loads(out)
        self.assertEqual(payload["output"], "42\n")
        self.assertIsNone(payload["error"])
        self.assertIn("duration_ms", payload)

    def test_runtime_error(self):
        path = self.write("boom.vbs", "echo 1 / 0\n")
        code, out, _ = run_cli(["--json", "run", path])
        self.assertEqual(code, 1)
        payload = json.loads(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["error"]["line"], 1)


class TestHumanMode(JsonCliCase):
    def test_check_ok_prints_ok(self):
        path = self.write("ok.vbs", "echo 1\n")
        code, out, _ = run_cli(["check", path])
        self.assertEqual(code, 0)
        self.assertIn("OK", out)

    def test_check_error_goes_to_stderr(self):
        path = self.write("bad.vbs", "If True Then\n")
        code, out, err = run_cli(["check", path])
        self.assertEqual(code, 1)
        self.assertTrue(err)
        self.assertEqual(out, "")

    def test_run_prints_output(self):
        path = self.write("run.vbs", "echo 5\n")
        code, out, _ = run_cli(["run", path])
        self.assertEqual(code, 0)
        self.assertEqual(out, "5")


if __name__ == "__main__":
    unittest.main(verbosity=2)
