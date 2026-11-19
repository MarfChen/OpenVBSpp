"""`openvbs -c CODE` / `-`(stdin)：命令行内联执行与诊断。"""

import contextlib
import io
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.cli import main as cli_main       # noqa: E402


def run_cli(argv, stdin=None):
    out, err = io.StringIO(), io.StringIO()
    saved_stdin = sys.stdin
    sys.stdin = io.StringIO("" if stdin is None else stdin)
    try:
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = cli_main(argv)
    finally:
        sys.stdin = saved_stdin
    return code, out.getvalue(), err.getvalue()


class TestInlineCode(unittest.TestCase):
    def test_dash_c_runs_source(self):
        code, out, _ = run_cli(["-c", "echo 1 + 1"])
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), "2")

    def test_long_form_code(self):
        code, out, _ = run_cli(["--code", "echo 21 * 2"])
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), "42")

    def test_code_and_script_are_exclusive(self):
        script = ROOT / "examples" / "fizzbuzz.vbs"
        code, out, err = run_cli(["-c", "echo 1", str(script)])
        self.assertEqual(code, 2)
        self.assertEqual(out.strip(), "")
        self.assertIn("二选一", err)

    def test_dash_c_stdin(self):
        code, out, _ = run_cli(["-c", "-"], stdin="echo 6 * 7\n")
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), "42")

    def test_script_dash_reads_stdin(self):
        code, out, _ = run_cli(["-"], stdin="echo 5 + 5\n")
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), "10")

    def test_runtime_error_returns_one(self):
        code, _, err = run_cli(["-c", "echo 1 / 0"])
        self.assertEqual(code, 1)
        self.assertIn("Division", err)


class TestInlineSubcommands(unittest.TestCase):
    def test_run_inline_json(self):
        code, out, _ = run_cli(["--json", "run", "-c", "echo 42"])
        self.assertEqual(code, 0)
        payload = json.loads(out)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["output"], "42\n")

    def test_check_inline_json_reports_line(self):
        code, out, _ = run_cli(["--json", "check", "-c", "If True Then"])
        self.assertEqual(code, 1)
        payload = json.loads(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["diagnostics"][0]["line"], 1)

    def test_check_inline_ok(self):
        code, out, _ = run_cli(["--json", "check", "-c", "echo 1"])
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out), {"ok": True, "diagnostics": []})

    def test_run_dash_reads_stdin(self):
        code, out, _ = run_cli(["--json", "run", "-"], stdin="echo 5 * 5\n")
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out)["output"], "25\n")

    def test_missing_code_after_flag(self):
        code, _, err = run_cli(["run", "-c"])
        self.assertEqual(code, 2)
        self.assertIn("usage", err)

    def test_subcommand_without_target(self):
        code, _, err = run_cli(["run"])
        self.assertEqual(code, 2)
        self.assertIn("usage", err)


if __name__ == "__main__":
    unittest.main()
