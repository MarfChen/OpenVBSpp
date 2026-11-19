'''
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

over 包的 python 依赖解析测试：pip 由注入的 runner 顶替，完全离线。

运行：
    python -m unittest tests.test_over_python -v
'''

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP import over                      # noqa: E402


def write_manifest(directory, name, version="1.0.0", dependencies=None,
                   python=None, platforms=None):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    data = {"name": name, "version": version, "entry": "",
            "dependencies": dict(dependencies or {})}
    if python:
        data["python"] = dict(python)
    if platforms:
        data["platforms"] = list(platforms)
    (directory / "over.json").write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8")
    return directory


class Recorder:
    """顶替 pip：记录调用参数，返回假装装好的包。"""

    def __init__(self, installed=None, error=None):
        self.calls = []
        self.installed = installed if installed is not None else [
            {"name": "Six", "version": "1.17.0", "requested": True}]
        self.error = error

    def __call__(self, requirements, python_executable=None, verbose=False):
        self.calls.append(list(requirements))
        if self.error is not None:
            raise self.error
        return list(self.installed)


class OverPythonBase(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="ovpp-over-"))
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.root = self.tmp / "app"
        self.root.mkdir()
        self.pkgs = self.tmp / "pkgs"
        (self.root / "over.json").write_text(
            json.dumps({"name": "app", "version": "0.1.0", "entry": "main.vbs",
                        "dependencies": {}}) + "\n", encoding="utf-8")

    def add_dependency(self, name, package_dir=None):
        """在 over.json 里登记一条依赖（默认指向 ../pkgs/<name>）。"""
        spec = "file:../pkgs/" + (package_dir or name)
        manifest = json.loads((self.root / "over.json").read_text("utf-8"))
        manifest["dependencies"][name] = spec
        (self.root / "over.json").write_text(
            json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8")

    def read_lock(self):
        return json.loads((self.root / "over.lock").read_text("utf-8"))


class TestPythonDependencies(OverPythonBase):
    def test_install_runs_pip_once_with_all_requirements(self):
        write_manifest(self.pkgs / "text-utils", "text-utils",
                       python={"six": ">=1.16"})
        self.add_dependency("text-utils")
        recorder = Recorder()
        over.install(self.root, pip_runner=recorder)

        self.assertEqual(recorder.calls, [["six>=1.16"]])
        lock = self.read_lock()
        self.assertEqual(lock["version"], 2)
        self.assertEqual([item["name"] for item in lock["python"]], ["six"])
        self.assertEqual(lock["python"][0]["version"], "1.17.0")
        self.assertEqual(lock["python"][0]["required_by"], ["text-utils"])

    def test_transitive_package_requirements_are_merged(self):
        write_manifest(self.pkgs / "leaf", "leaf", python={"six": "<2"})
        write_manifest(self.pkgs / "middle", "middle",
                       dependencies={"leaf": "file:../pkgs/leaf"},
                       python={"six": ">=1.16"})
        self.add_dependency("middle")
        recorder = Recorder()
        over.install(self.root, pip_runner=recorder)

        self.assertEqual(len(recorder.calls), 1, "所有 python 依赖应一次性交给 pip")
        self.assertEqual(recorder.calls[0], ["six>=1.16,<2"])

    def test_python_deps_of_incompatible_platform_are_skipped(self):
        write_manifest(self.pkgs / "linux-only", "linux-only",
                       python={"uvloop": ">=0.19"}, platforms=["linux"])
        self.add_dependency("linux-only")
        recorder = Recorder()
        over.install(self.root, pip_runner=recorder)
        self.assertEqual(recorder.calls, [])
        self.assertNotIn("python", self.read_lock())

    def test_no_python_flag_skips_pip(self):
        write_manifest(self.pkgs / "text-utils", "text-utils",
                       python={"six": ">=1.16"})
        self.add_dependency("text-utils")
        recorder = Recorder()
        over.install(self.root, pip_runner=recorder, python_deps=False)
        self.assertEqual(recorder.calls, [])
        self.assertNotIn("python", self.read_lock())

    def test_project_without_python_deps_never_calls_pip(self):
        write_manifest(self.pkgs / "plain", "plain")
        self.add_dependency("plain")
        recorder = Recorder()
        over.install(self.root, pip_runner=recorder)
        self.assertEqual(recorder.calls, [])

    def test_pip_failure_is_reported(self):
        write_manifest(self.pkgs / "text-utils", "text-utils",
                       python={"six": ">=1.16"})
        self.add_dependency("text-utils")
        recorder = Recorder(error=over.OverError("pip exploded"))
        with self.assertRaises(over.OverError) as ctx:
            over.install(self.root, pip_runner=recorder)
        self.assertIn("pip exploded", str(ctx.exception))

    def test_frozen_install_pins_python_versions(self):
        write_manifest(self.pkgs / "text-utils", "text-utils",
                       python={"six": ">=1.16"})
        self.add_dependency("text-utils")
        over.install(self.root, pip_runner=Recorder())
        lock_before = (self.root / "over.lock").read_bytes()

        shutil.rmtree(self.root / "packages")
        recorder = Recorder(installed=[
            {"name": "six", "version": "1.17.0", "requested": True}])
        over.install(self.root, pip_runner=recorder, frozen=True)

        self.assertEqual(recorder.calls, [["six==1.17.0"]])
        self.assertEqual((self.root / "over.lock").read_bytes(), lock_before)

    def test_remove_python_dependency(self):
        write_manifest(self.pkgs / "text-utils", "text-utils",
                       python={"six": ">=1.16"})
        self.add_dependency("text-utils")
        over.install(self.root, pip_runner=Recorder())

        uninstaller = Recorder(installed=[])
        over.remove(self.root, "six", python=True, pip_runner=uninstaller)
        self.assertEqual(uninstaller.calls, [["six"]])

        self.assertNotIn("python", self.read_lock())
        self.assertEqual(over.list_installed(self.root),
                         [("text-utils", "1.0.0")])

    def test_list_installed_details_shows_kinds(self):
        write_manifest(self.pkgs / "text-utils", "text-utils",
                       python={"six": ">=1.16"})
        self.add_dependency("text-utils")
        over.install(self.root, pip_runner=Recorder())
        details = {item["name"]: item
                   for item in over.list_installed_details(self.root)}
        self.assertEqual(details["text-utils"]["kind"], "over")
        self.assertEqual(details["six"]["kind"], "python")
        self.assertEqual(details["six"]["required_by"], ["text-utils"])

    def test_info_reports_python_dependencies(self):
        write_manifest(self.pkgs / "text-utils", "text-utils",
                       python={"six": ">=1.16"})
        self.add_dependency("text-utils")
        over.install(self.root, pip_runner=Recorder())
        entry = over.info(self.root, "text-utils")
        self.assertIsNotNone(entry)
        self.assertEqual(entry["python_dependencies"], {"six": ">=1.16"})

    def test_info_missing_package_returns_none(self):
        self.assertIsNone(over.info(self.root, "nope"))


class TestComSupportPackage(OverPythonBase):
    def test_provision_bundled_package(self):
        entry = over.provision_bundled(self.root)
        self.assertEqual(entry["name"], "com-support")
        self.assertEqual(entry["kind"], "bundled")
        target = self.root / "packages" / "com-support"
        self.assertTrue((target / "com_support.py").is_file())
        self.assertEqual(over.list_installed(self.root),
                         [("com-support", "1.0.0")])

    def test_install_with_com_records_extra_and_python_dep(self):
        recorder = Recorder(installed=[
            {"name": "pywin32", "version": "312", "requested": True}])
        over.install(self.root, with_com=True, pip_runner=recorder)
        lock = self.read_lock()
        self.assertEqual(lock["extras"][0]["name"], "com-support")
        self.assertEqual(lock["python"][0]["name"], "pywin32")
        self.assertEqual(lock["python"][0]["required_by"], ["com-support"])
        self.assertEqual(recorder.calls, [["pywin32>=306"]])

    def test_remove_com_support_package(self):
        over.install(self.root, with_com=True,
                     pip_runner=Recorder(installed=[
                         {"name": "pywin32", "version": "312",
                          "requested": True}]))
        over.remove(self.root, "com-support")
        self.assertFalse((self.root / "packages" / "com-support").exists())
        self.assertNotIn("extras", self.read_lock())

    def test_remove_missing_package_raises(self):
        with self.assertRaises(over.OverError):
            over.remove(self.root, "ghost")


class TestRequirementHelpers(unittest.TestCase):
    def test_normalize_name(self):
        self.assertEqual(over._normalize_name("PyWin32"), "pywin32")
        self.assertEqual(over._normalize_name("zope_interface"), "zope-interface")
        self.assertEqual(over._normalize_name("a.b.C"), "a-b-c")

    def test_split_pip_requirement(self):
        self.assertEqual(over.split_pip_requirement("pywin32>=306"),
                         ("pywin32", ">=306"))
        self.assertEqual(over.split_pip_requirement("six"),
                         ("six", ""))
        self.assertEqual(over.split_pip_requirement("six >= 1.16"),
                         ("six", ">= 1.16"))

    def test_requirement_strings_accepts_dict_and_list(self):
        self.assertEqual(over.requirement_strings({"six": ">=1.16"}),
                         ["six>=1.16"])
        self.assertEqual(over.requirement_strings(["six>=1.16", " "]),
                         ["six>=1.16"])

    def test_manifest_supports_platform_aliases(self):
        manifest = over.Manifest.__new__(over.Manifest)
        manifest.data = {"platforms": ["windows"]}
        self.assertTrue(manifest.supports_platform("win32"))
        manifest.data = {"platforms": ["macos"]}
        self.assertTrue(manifest.supports_platform("darwin"))
        manifest.data = {}
        self.assertTrue(manifest.supports_platform("linux"),
                        "没声明 platforms 应当算全平台")


if __name__ == "__main__":
    unittest.main()
