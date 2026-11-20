"""Service-layer tests for the built-in web IDE: no browser, pure logic plus one HTTP smoke.

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
内置网页 IDE 的服务层测试（不启浏览器，纯逻辑 + 一次 HTTP 冒烟）。
"""

import json
import sys
import threading
import unittest
import urllib.request
from pathlib import Path
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.ide import service                     # noqa: E402
from OpenVBSPP.ide.server import create_server        # noqa: E402


class TestService(unittest.TestCase):
    def test_run_produces_output(self):
        result = service.run_source('echo "hi"', "<t>")
        self.assertTrue(result["ok"])
        self.assertEqual(result["output"], "hi\n")
        self.assertIsNone(result["error"])

    def test_runtime_error_is_structured(self):
        result = service.run_source("echo missing", "<t>")
        self.assertFalse(result["ok"])
        self.assertEqual(result["error"]["line"], 1)
        self.assertIn("missing", result["error"]["message"])

    def test_check_reports_syntax_error_without_running(self):
        result = service.check_source("If True Then\n", "<t>")
        self.assertFalse(result["ok"])
        self.assertTrue(result["diagnostics"])

    def test_check_passes_for_valid_code(self):
        result = service.check_source("echo 1", "<t>")
        self.assertTrue(result["ok"])
        self.assertEqual(result["diagnostics"], [])

    def test_option_typing_errors_surface_in_check(self):
        code = 'Option Typing\nDim n -> Integer\nn = "x"'
        result = service.check_source(code, "<t>")
        self.assertFalse(result["ok"])

    def test_empty_source_is_ok(self):
        result = service.run_source("   ", "<t>")
        self.assertTrue(result["ok"])
        self.assertEqual(result["output"], "")

    def test_examples_are_runnable(self):
        for example in service.examples():
            result = service.run_source(example["source"], "<demo>")
            self.assertTrue(result["ok"], f"example {example['name']} failed")

    def test_capabilities_marks_packages_reserved(self):
        caps = service.capabilities()
        self.assertFalse(caps["features"]["packages"])
        self.assertFalse(caps["packages"]["enabled"])
        self.assertEqual(caps["packages"]["provider"], "LocalPackageStore")


class TestFileHelpers(unittest.TestCase):
    def test_read_normalizes_crlf_and_bom(self):
        from OpenVBSPP.ide import dialogs
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "crlf.vbs"
            path.write_bytes(b"\xef\xbb\xbfecho 1\r\necho 2\r\n")
            self.assertEqual(dialogs.read_script(path), "echo 1\necho 2\n")

    def test_write_uses_lf_regardless_of_platform(self):
        from OpenVBSPP.ide import dialogs
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "lf.vbs"
            dialogs.write_script(path, "echo 1\necho 2\n")
            self.assertEqual(path.read_bytes(), b"echo 1\necho 2\n")

    def test_platform_hint(self):
        from OpenVBSPP.ide import dialogs
        hint = dialogs.platform_hint()
        self.assertIn("native_dialogs", hint)
        self.assertEqual(hint["newline"], "LF (normalized on read/write)")


class TestWorkspace(unittest.TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.ws = service.Workspace(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_save_then_load(self):
        self.ws.save("demo", "echo 42")
        self.assertEqual(self.ws.list(), ["demo"])
        loaded = self.ws.load("demo")
        self.assertEqual(loaded["source"], "echo 42")

    def test_load_missing(self):
        self.assertFalse(self.ws.load("nope")["ok"])

    def test_path_escape_rejected(self):
        with self.assertRaises(ValueError):
            self.ws.save("../evil", "x")


class TestServer(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = TemporaryDirectory()
        workspace = service.Workspace(cls.tmp.name)
        cls.httpd = create_server("127.0.0.1", 0, workspace)
        cls.port = cls.httpd.server_address[1]
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        cls.tmp.cleanup()

    def url(self, path):
        return f"http://127.0.0.1:{self.port}{path}"

    def test_index_page_is_served(self):
        with urllib.request.urlopen(self.url("/")) as response:
            body = response.read().decode("utf-8")
        self.assertIn("OpenVBS++ Studio", body)

    def test_capabilities_endpoint(self):
        with urllib.request.urlopen(self.url("/api/capabilities")) as response:
            data = json.loads(response.read().decode("utf-8"))
        self.assertEqual(data["project"], "OpenVBS++")

    def test_run_endpoint(self):
        request = urllib.request.Request(
            self.url("/api/run"),
            data=json.dumps({"source": "echo 6*7"}).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(request) as response:
            data = json.loads(response.read().decode("utf-8"))
        self.assertEqual(data["output"], "42\n")

    def test_dialogs_endpoint_reports_platform(self):
        with urllib.request.urlopen(self.url("/api/dialogs")) as response:
            data = json.loads(response.read().decode("utf-8"))
        self.assertIn("native_dialogs", data)
        self.assertIsInstance(data["native_dialogs"], bool)

    def test_file_save_and_open_roundtrip(self):
        """Absolute-path read/write, which the native file dialog uses afterwards.

        中文：绝对路径读写（原生对话框选完文件之后用的就是这两个接口）。"""
        target = Path(self.tmp.name) / "picked.vbs"
        request = urllib.request.Request(
            self.url("/api/file/save"),
            data=json.dumps({"path": str(target), "source": "echo 7"}).encode(),
            headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(request) as response:
            saved = json.loads(response.read().decode("utf-8"))
        self.assertTrue(saved["ok"])
        self.assertTrue(target.is_file())
        self.assertEqual(target.read_text(encoding="utf-8"), "echo 7")

        request = urllib.request.Request(
            self.url("/api/file/open"),
            data=json.dumps({"path": str(target)}).encode(),
            headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(request) as response:
            loaded = json.loads(response.read().decode("utf-8"))
        self.assertEqual(loaded["source"], "echo 7")

    def test_packages_endpoint_is_reserved(self):
        request = urllib.request.Request(
            self.url("/api/packages/search"),
            data=b"{}", headers={"Content-Type": "application/json"})
        with self.assertRaises(urllib.error.HTTPError) as ctx:
            urllib.request.urlopen(request)
        self.assertEqual(ctx.exception.code, 501)


if __name__ == "__main__":
    unittest.main(verbosity=2)
