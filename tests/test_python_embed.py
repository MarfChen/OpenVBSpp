"""
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

Python 互操作能力包（python-embed）：随包分发、显式导入、可安装/可卸载。
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

from OpenVBSPP import over                                 # noqa: E402
from OpenVBSPP.exceptions import OVPPSyntaxError, OVPPError  # noqa: E402
from OpenVBSPP.ovpprt.cop import reset as reset_cop        # noqa: E402
from OpenVBSPP.runtime import vbs_run                      # noqa: E402

VENDOR = ROOT / "OpenVBSPP" / "vendor" / "python-embed"
EXTRAS = ROOT / "extras" / "python-embed"
IMPORT = '<script src="python-embed" />\n'


def out(code, base_dir=None):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<python-embed>")
    return buf.getvalue().splitlines()


class ManifestCapabilityTest(unittest.TestCase):
    def test_vendor_manifest_declares_capability(self):
        manifest = over.Manifest.load(VENDOR)
        self.assertEqual(manifest.name, "python-embed")
        self.assertEqual(manifest.capabilities, ["python-embed"])

    def test_capabilities_accepts_string_and_list(self):
        single = over.Manifest({"capabilities": "python-embed"})
        self.assertEqual(single.capabilities, ["python-embed"])
        many = over.Manifest({"capabilities": [" COM ", "python-embed"]})
        self.assertEqual(many.capabilities, ["com", "python-embed"])
        self.assertEqual(over.Manifest({}).capabilities, [])

    def test_bundled_packages_lists_feature_and_com(self):
        names = [name for name, _version in over.bundled_packages()]
        self.assertIn("python-embed", names)
        self.assertIn("com-support", names)


class VendorFallbackTest(unittest.TestCase):
    def test_project_package_wins_then_vendor(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            resolved = over.resolve_source("python-embed", base)
            self.assertIsNotNone(resolved)
            self.assertEqual(resolved.parent, VENDOR)

            local = base / over.PACKAGES_DIR / "python-embed"
            local.mkdir(parents=True)
            (local / over.MANIFEST_NAME).write_text(
                '{"name": "python-embed", "version": "9.9.9", '
                '"entry": "main.vbs", "capabilities": ["python-embed"]}',
                encoding="utf-8")
            (local / "main.vbs").write_text("' local override\n", encoding="utf-8")
            resolved = over.resolve_source("python-embed", base)
            self.assertEqual(resolved.parent, local)

    def test_unknown_package_still_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertIsNone(over.resolve_source("no-such-package", Path(tmp)))


class GatingTest(unittest.TestCase):
    def tearDown(self):
        reset_cop()

    def test_python_block_without_import_is_rejected(self):
        with self.assertRaises(OVPPSyntaxError) as ctx:
            out('#python\nx = 1\n#endpython')
        self.assertIn("python-embed", str(ctx.exception))

    def test_python_block_with_import_runs(self):
        self.assertEqual(
            out(IMPORT + '#python\ny = 6 * 7\n#endpython\necho y'), ["42"])

    def test_cop_needs_import(self):
        with self.assertRaises(OVPPError) as ctx:
            out('COP("k") = 1')
        self.assertIn("python-embed", str(ctx.exception))

    def test_cop_works_with_import(self):
        self.assertEqual(
            out(IMPORT + 'COP("k") = 7\necho COP("k")'), ["7"])

    def test_option_cop_is_gone(self):
        with self.assertRaises(OVPPSyntaxError) as ctx:
            out('Option COP\necho 1')
        message = str(ctx.exception)
        self.assertIn("Option COP", message)
        self.assertIn("python-embed", message)


class InstallLifecycleTest(unittest.TestCase):
    def test_over_add_and_remove_bundled_feature(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = Path(tmp)
            over.init(project, "app")
            over.add(project, "python-embed")
            target = project / over.PACKAGES_DIR / "python-embed"
            self.assertTrue((target / "main.vbs").is_file())
            lock = over.read_lock(project)
            extras = [item["name"] for item in lock.get("extras", [])]
            self.assertIn("python-embed", extras)

            names = [item["name"] for item in over.list_installed_details(project)]
            self.assertIn("python-embed", names)

            over.remove(project, "python-embed")
            self.assertFalse(target.exists())
            # 卸载之后仍能用：退回随包分发的副本
            self.assertEqual(over.resolve_source("python-embed", project),
                             VENDOR / "main.vbs")

    def test_over_add_unknown_package_without_source_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            project = Path(tmp)
            over.init(project, "app")
            with self.assertRaises(over.OverError):
                over.add(project, "no-such-package")


class PackagingTest(unittest.TestCase):
    def test_extras_copy_is_identical(self):
        for name in ("over.json", "main.vbs"):
            self.assertEqual((VENDOR / name).read_bytes(),
                             (EXTRAS / name).read_bytes(),
                             f"{name} drifted between vendor and extras")

    def test_vendor_vbs_ships_in_package_data(self):
        pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
        self.assertIn("vendor/python-embed/*.vbs", pyproject)
        manifest_in = (ROOT / "MANIFEST.in").read_text(encoding="utf-8")
        self.assertIn("recursive-include OpenVBSPP/vendor *.py *.json *.vbs *.md",
                      manifest_in)

    def test_pyinstaller_spec_bundles_feature(self):
        spec = (ROOT / "packaging" / "openvbs.spec").read_text(encoding="utf-8")
        self.assertIn('"vendor", "python-embed"', spec)
        self.assertIn("BUNDLED_EMBED", spec)


if __name__ == "__main__":
    unittest.main(verbosity=2)
