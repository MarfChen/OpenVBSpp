"""构造参数 + over 包管理器（M1）。"""

import contextlib
import io
import sys
import json
import subprocess
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP import over                       # noqa: E402
from OpenVBSPP.exceptions import OVPPError       # noqa: E402
from OpenVBSPP.runtime import vbs_run            # noqa: E402


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<t>")
    return buf.getvalue().splitlines()


class TestConstructorArgs(unittest.TestCase):
    def test_new_passes_arguments(self):
        code = ("Class P\n Public name\n"
                " Private Sub Class_Initialize(n)\n  name = n\n End Sub\n"
                "End Class\nSet p = New P(\"ann\")\necho p.name")
        self.assertEqual(out(code), ["ann"])

    def test_multiple_arguments(self):
        code = ("Class P\n Public a, b\n"
                " Private Sub Class_Initialize(x, y)\n  a = x\n  b = y\n End Sub\n"
                "End Class\nSet p = New P(1, 2)\necho p.a + p.b")
        self.assertEqual(out(code), ["3"])

    def test_createobject_rejects_servername(self):
        # VBS 的第二个参数是 servername（远程激活），本地不支持
        code = ("Class P\n Public name\n"
                " Private Sub Class_Initialize(n)\n  name = n\n End Sub\n"
                "End Class\nSet p = CreateObject(\"P\", \"bo\")")
        with self.assertRaises(OVPPError):
            out(code)

    def test_no_arguments_still_works(self):
        code = "Class P\n Public x\nEnd Class\nSet p = New P\necho IsObject(p)"
        self.assertEqual(out(code), ["True"])

    def test_unknown_class_with_arguments(self):
        with self.assertRaises(OVPPError):
            out("Set p = New Nope(1)")


class TestOverVersion(unittest.TestCase):
    def test_satisfies_ranges(self):
        self.assertTrue(over.satisfies("1.2.3", "^1.2"))
        self.assertFalse(over.satisfies("2.0.0", "^1.2"))
        self.assertTrue(over.satisfies("1.5.0", ">=1.0"))
        self.assertFalse(over.satisfies("0.9.0", ">=1.0"))
        self.assertTrue(over.satisfies("1.0.0", ""))

    def test_parse_version_pads(self):
        self.assertEqual(over.parse_version("1.2"), (1, 2, 0))
        self.assertEqual(over.parse_version("2"), (2, 0, 0))


class TestOverManifest(unittest.TestCase):
    def test_load_and_save(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "over.json").write_text(
                '{"name": "app", "version": "1.0.0", '
                '"dependencies": {"lib": "file:../lib"}}', encoding="utf-8")
            manifest = over.Manifest.load(root)
            self.assertEqual(manifest.name, "app")
            self.assertEqual(manifest.dependencies, {"lib": "file:../lib"})
            manifest.save()
            self.assertTrue((root / "over.json").is_file())

    def test_missing_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(over.OverError):
                over.Manifest.load(tmp)


class TestOverInstall(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def _make_lib(self, name, version, source):
        lib = self.base / ("src_" + name)
        lib.mkdir(parents=True, exist_ok=True)
        (lib / "over.json").write_text(
            '{"name": "%s", "version": "%s", "entry": "main.vbs"}' % (name, version),
            encoding="utf-8")
        (lib / "main.vbs").write_text(source, encoding="utf-8")
        return lib

    def test_install_local_dependency(self):
        lib = self._make_lib("text-utils", "1.2.0",
                             'Public Function Shout(s)\n    Shout = UCase(s)\nEnd Function\n')
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"text-utils": "file:../src_text-utils"}}',
            encoding="utf-8")

        installed = over.install(app)
        self.assertEqual(installed[0]["version"], "1.2.0")
        self.assertTrue((app / "packages" / "text-utils" / "main.vbs").is_file())
        self.assertTrue((app / "over.lock").is_file())
        self.assertEqual(over.list_installed(app), [("text-utils", "1.2.0")])
        self.assertIsNotNone(over.resolve_source("text-utils", app))
        self.assertIsNone(over.resolve_source("missing", app))

    def test_namespace_resolves_to_installed_package(self):
        self._make_lib("mathx", "1.0.0",
                       'Public Function Double(n)\n    Double = n * 2\nEnd Function\n')
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"mathx": "file:../src_mathx"}}', encoding="utf-8")
        over.install(app)

        main = app / "main.vbs"
        main.write_text('<namespace src="mathx" alias="m">\necho m.Double(21)\n',
                        encoding="utf-8")
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            vbs_run(main.read_text(encoding="utf-8"), filename=str(main))
        self.assertEqual(buf.getvalue().strip(), "42")

    def test_pack_then_install(self):
        lib = self._make_lib("packed", "1.0.0", "echo 1\n")
        archive = over.pack(lib)
        self.assertTrue(archive.is_file())
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"packed": "file:../%s"}}' % archive.name,
            encoding="utf-8")
        installed = over.install(app)
        self.assertEqual(installed[0]["version"], "1.0.0")
        self.assertTrue((app / "packages" / "packed" / "main.vbs").is_file())

    def test_conflict_reports_chain(self):
        shared = self._make_lib("shared", "1.5.0", "echo 1\n")
        a = self.base / "src_a"
        a.mkdir()
        (a / "over.json").write_text(
            '{"name": "a", "version": "1.0.0", '
            '"dependencies": {"shared": {"source": "../src_shared", '
            '"version": "^1.0"}}}', encoding="utf-8")
        b = self.base / "src_b"
        b.mkdir()
        (b / "over.json").write_text(
            '{"name": "b", "version": "1.0.0", '
            '"dependencies": {"shared": {"source": "../src_shared", '
            '"version": "^2.0"}}}', encoding="utf-8")

        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", "dependencies": {'
            '"a": "file:../src_a", "b": "file:../src_b"}}', encoding="utf-8")
        with self.assertRaises(over.OverError) as ctx:
            over.install(app)
        self.assertIn("Conflict for 'shared'", str(ctx.exception))
        self.assertIn("app -> b", str(ctx.exception))

    def test_circular_dependency_reports_chain(self):
        a = self.base / "src_a"
        a.mkdir()
        (a / "over.json").write_text(
            '{"name": "a", "version": "1.0.0", '
            '"dependencies": {"b": "file:../src_b"}}', encoding="utf-8")
        b = self.base / "src_b"
        b.mkdir()
        (b / "over.json").write_text(
            '{"name": "b", "version": "1.0.0", '
            '"dependencies": {"a": "file:../src_a"}}', encoding="utf-8")
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"a": "file:../src_a"}}', encoding="utf-8")
        with self.assertRaises(over.OverError) as ctx:
            over.install(app)
        self.assertIn("Circular", str(ctx.exception))

    def test_init_creates_manifest(self):
        app = self.base / "fresh"
        path = over.init(app, "fresh")
        self.assertTrue(path.is_file())
        manifest = over.Manifest.load(app)
        self.assertEqual(manifest.name, "fresh")
        with self.assertRaises(over.OverError):
            over.init(app)

    def test_frozen_install_uses_lock(self):
        self._make_lib("lib", "1.0.0", "echo 1\n")
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"lib": "file:../src_lib"}}', encoding="utf-8")
        over.install(app)
        lock_before = (app / "over.lock").read_text(encoding="utf-8")

        # 删掉安装结果，frozen 安装应能按锁恢复
        import shutil as _sh
        _sh.rmtree(app / "packages")
        installed = over.install(app, frozen=True)
        self.assertEqual(installed[0]["version"], "1.0.0")
        self.assertEqual((app / "over.lock").read_text(encoding="utf-8"),
                         lock_before)

    def test_frozen_install_rejects_drifted_source(self):
        lib = self._make_lib("lib", "1.0.0", "echo 1\n")
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"lib": "file:../src_lib"}}', encoding="utf-8")
        over.install(app)
        # 源升级到 2.0.0，frozen 安装必须报错而不是悄悄升级
        (lib / "over.json").write_text(
            '{"name": "lib", "version": "2.0.0", "entry": "main.vbs"}',
            encoding="utf-8")
        with self.assertRaises(over.OverError):
            over.install(app, frozen=True)

    def test_unresolvable_source_reports_error(self):
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"ghost": "file:../nowhere"}}', encoding="utf-8")
        with self.assertRaises(over.OverError):
            over.install(app)

    def test_version_mismatch_reports_error(self):
        self._make_lib("lib", "2.0.0", "echo 1\n")
        app = self.base / "app"
        app.mkdir()
        (app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"lib": {"source": "file:../src_lib", '
            '"version": "^1.0"}}}', encoding="utf-8")
        with self.assertRaises(over.OverError):
            over.install(app)


class TestOverDryRun(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        lib = self.base / "src_lib"
        lib.mkdir()
        (lib / "over.json").write_text(
            '{"name": "lib", "version": "1.0.0", "entry": "main.vbs"}',
            encoding="utf-8")
        (lib / "main.vbs").write_text("Public Function F()\n    F = 1\nEnd Function\n",
                                      encoding="utf-8")
        self.app = self.base / "app"
        self.app.mkdir()
        (self.app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"lib": "file:../src_lib"}}', encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def test_dry_run_does_not_write_lock(self):
        plan = over.install(self.app, dry_run=True)
        self.assertEqual(plan[0]["action"], "install")
        self.assertEqual(plan[0]["name"], "lib")
        self.assertEqual(plan[0]["version"], "1.0.0")
        self.assertFalse((self.app / "over.lock").exists())
        self.assertFalse((self.app / "packages").exists())

    def test_second_dry_run_reports_unchanged(self):
        over.install(self.app)
        plan = over.install(self.app, dry_run=True)
        self.assertEqual(plan[0]["action"], "unchanged")
        self.assertEqual(plan[0]["current"], "1.0.0")


class TestOverSearch(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        lib = self.base / "src_lib"
        lib.mkdir()
        (lib / "over.json").write_text(
            '{"name": "lib", "version": "2.0.0", "entry": "main.vbs"}',
            encoding="utf-8")
        (lib / "main.vbs").write_text("echo 1\n", encoding="utf-8")
        self.app = self.base / "app"
        self.app.mkdir()
        (self.app / "over.json").write_text(
            '{"name": "app", "version": "0.1.0", '
            '"dependencies": {"lib": "file:../src_lib"}}', encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def test_search_finds_installed(self):
        over.install(self.app)
        found = over.search("lib", self.app)
        self.assertTrue(any(item["name"] == "lib" and item["source"] == "installed"
                            for item in found))

    def test_search_uses_registry_entries(self):
        registry = {"text-utils": {"version": "1.4.0", "description": "string helpers"}}
        found = over.search("text", None, registry=registry)
        self.assertEqual(found[0]["name"], "text-utils")
        self.assertEqual(found[0]["version"], "1.4.0")
        self.assertEqual(found[0]["source"], "registry")

    def test_search_no_match(self):
        self.assertEqual(over.search("nothing-here", None, registry={}), [])


class TestOverGitTags(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.repo = Path(self.tmp.name) / "repo"
        self.repo.mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def _git(self, *args):
        return subprocess.run(["git", "-C", str(self.repo)] + list(args),
                              check=True, capture_output=True, text=True)

    def test_tags_are_enumerated_in_order(self):
        try:
            self._git("init", "-q")
            self._git("-c", "user.email=a@b", "-c", "user.name=t",
                      "commit", "-q", "--allow-empty", "-m", "c1")
        except (OSError, subprocess.CalledProcessError):
            self.skipTest("git not available")
        self._git("tag", "v1.0.0")
        self._git("-c", "user.email=a@b", "-c", "user.name=t",
                  "commit", "-q", "--allow-empty", "-m", "c2")
        self._git("tag", "v1.2.0")

        tags = over.git_tags("git+" + self.repo.as_uri())
        self.assertEqual(tags, ["v1.0.0", "v1.2.0"])

        chosen = over.select_git_ref("git+" + self.repo.as_uri(), "^1.0")
        self.assertEqual(chosen, "v1.2.0")


class TestOverLockCommit(unittest.TestCase):
    def test_lock_records_commit_for_git_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            repo = base / "repo"
            repo.mkdir()
            try:
                subprocess.run(["git", "-C", str(repo), "init", "-q"],
                               check=True, capture_output=True)
                (repo / "over.json").write_text(
                    '{"name": "glib", "version": "1.0.0", "entry": "m.vbs"}',
                    encoding="utf-8")
                (repo / "m.vbs").write_text("echo 1\n", encoding="utf-8")
                subprocess.run(["git", "-C", str(repo),
                                "-c", "user.email=a@b", "-c", "user.name=t",
                                "add", "-A"], check=True, capture_output=True)
                subprocess.run(["git", "-C", str(repo),
                                "-c", "user.email=a@b", "-c", "user.name=t",
                                "commit", "-q", "-m", "init"],
                               check=True, capture_output=True)
            except (OSError, subprocess.CalledProcessError):
                self.skipTest("git not available")

            app = base / "app"
            app.mkdir()
            (app / "over.json").write_text(
                '{"name": "app", "version": "0.1.0", "dependencies": '
                '{"glib": "git+' + repo.as_uri() + '"}}', encoding="utf-8")

            over.install(app)
            lock = json.loads((app / "over.lock").read_text(encoding="utf-8"))
            entry = lock["packages"][0]
            self.assertIn("commit", entry)
            self.assertEqual(len(entry["commit"]), 40)


if __name__ == "__main__":
    unittest.main(verbosity=2)
