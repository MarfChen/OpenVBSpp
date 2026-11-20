"""com-support tests: bundled distribution, runtime resolution, CreateObject without any import.

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

COM 支持包（com-support）测试：随包分发、运行期解析、无需 import 的 CreateObject。

运行：
    python -m unittest tests.test_com -v
"""

import contextlib
import importlib.util
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP import over                                  # noqa: E402
from OpenVBSPP.exceptions import (OVPPAutomationError,        # noqa: E402
                                 OVPPObjectNotSupportedError)
from OpenVBSPP.ovpprt import com                            # noqa: E402
from OpenVBSPP.runtime import vbs_run                       # noqa: E402

VENDOR = ROOT / "OpenVBSPP" / "vendor" / "com-support"
EXTRAS = ROOT / "extras" / "com-support"

# COM itself is Windows-only, so the tests that need a live provider only run
# there; the packaging and manifest checks below stay platform independent.
# 中文：COM 本身只在 Windows 上可用，因此需要真实 provider 的用例只在 Windows
# 跑；下面的打包与清单检查仍然与平台无关。
WINDOWS_ONLY = unittest.skipUnless(sys.platform == "win32",
                                  "com-support targets Windows")


def out(code):
    """Run a piece of VBS and return its output.

    中文：跑一段 VBS 并返回它的输出。"""
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        vbs_run(code, filename="<com>")
    return buffer.getvalue()


FAKE_SUPPORT = '''\
VERSION = "9.9.9"
VBS_TYPE_NAME = "Object"


class ComSupportError(Exception):
    pass


def is_available():
    return True


class FakeObject:
    _vbs_name = "Object"

    def __init__(self, progid=""):
        self.progid = progid
        self.value = 0


def create(progid, *args):
    return FakeObject(progid)


def get(progid, *args):
    return FakeObject(progid)
'''

BROKEN_SUPPORT = '''\
VERSION = "0.0.1"


class ComSupportError(Exception):
    pass


def is_available():
    return True


def create(progid, *args):
    raise ComSupportError("the fake class is not registered")
'''

NAMED_ERROR_SUPPORT = '''\
VERSION = "0.0.3"
VBS_TYPE_NAME = "Object"


class ComSupportError(Exception):
    pass


class ComCreateError(ComSupportError):
    number = 429
    vbs_name = "ActiveX component can't create object"


class ComAttachError(ComSupportError):
    number = 432
    vbs_name = "File name or class name not found"


def is_available():
    return True


class FakeObject:
    _vbs_name = "Object"

    def __init__(self, progid=""):
        self.progid = progid


def create(progid, *args):
    raise ComCreateError("ActiveX component can't create object: '%s'" % progid)


def get(pathname="", classname=None):
    return FakeObject(classname or pathname)
'''

UNOBSERVABLE_SUPPORT = '''\
VERSION = "0.0.2"


class ComSupportError(Exception):
    pass


def is_available():
    return False


def create(progid, *args):
    raise ComSupportError("should not be reachable")
'''


def write_package(directory, source=FAKE_SUPPORT, version="9.9.9"):
    """Build an over package shaped like com-support.

    中文：造一个 com-support 形状的 over 包。"""
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "com_support.py").write_text(source, encoding="utf-8")
    manifest = {"name": "com-support", "version": version,
                "entry": "", "dependencies": {},
                "platforms": ["win32"], "python": {"pywin32": ">=306"}}
    (directory / "over.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8")
    return directory


class ComTestCase(unittest.TestCase):
    """Move cwd into a temporary project and clear env vars and module caches.

    中文：把 cwd 挪到临时项目里，清掉环境变量与模块缓存。"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="ovpp-com-"))
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self._old_cwd = Path.cwd()
        os.chdir(self.tmp)
        self.addCleanup(os.chdir, self._old_cwd)
        # temporary project root: com.py finds the project through over.json
        # 中文：临时项目根：com.py 靠 over.json 认项目目录
        (self.tmp / "over.json").write_text(
            json.dumps({"name": "demo", "version": "0.1.0", "entry": "main.vbs",
                        "dependencies": {}}) + "\n", encoding="utf-8")
        com.reset_cache()
        self.addCleanup(com.reset_cache)
        patcher = mock.patch.dict(os.environ, {}, clear=False)
        patcher.start()
        self.addCleanup(patcher.stop)
        os.environ.pop(com.ENV_DISABLE, None)
        os.environ.pop(com.ENV_PATH, None)


class TestPackageLayout(unittest.TestCase):
    def test_vendor_package_files(self):
        self.assertTrue((VENDOR / "com_support.py").is_file())
        self.assertTrue((VENDOR / "over.json").is_file())

    def test_manifest_declares_windows_and_pywin32(self):
        manifest = over.Manifest.load(VENDOR)
        self.assertEqual(manifest.name, "com-support")
        self.assertEqual(manifest.platforms, ["win32"])
        self.assertEqual(manifest.python_dependencies, {"pywin32": ">=306"})
        self.assertTrue(manifest.supports_platform("win32"))
        self.assertFalse(manifest.supports_platform("linux"))

    def test_extras_mirror_matches_vendor(self):
        self.assertTrue(EXTRAS.is_dir(), "extras/com-support 应当存在")
        for name in ("com_support.py", "over.json"):
            with self.subTest(name=name):
                self.assertEqual((VENDOR / name).read_bytes(),
                                 (EXTRAS / name).read_bytes(),
                                 f"{name} 在 extras 与 vendor 里必须一致")

    def test_adapter_is_self_contained(self):
        source = (VENDOR / "com_support.py").read_text(encoding="utf-8")
        self.assertNotIn("import OpenVBSPP", source)
        self.assertNotIn("from OpenVBSPP", source)
        self.assertNotIn("from .", source)

    def test_over_status_reports_bundled_copy(self):
        status = over.com_support_status()
        self.assertEqual(Path(status["directory"]).resolve(), VENDOR.resolve())
        self.assertEqual(status["origin"], "bundled")
        self.assertEqual(status["available"], sys.platform == "win32")


class TestBridgeResolution(ComTestCase):
    @WINDOWS_ONLY
    def test_falls_back_to_bundled_copy(self):
        directory, origin = com.resolve()
        self.assertEqual(origin, "bundled")
        self.assertEqual(Path(directory).resolve(), VENDOR.resolve())

    @WINDOWS_ONLY
    def test_project_package_wins(self):
        write_package(Path("packages") / "com-support")
        directory, origin = com.resolve()
        self.assertEqual(origin, "project")
        self.assertEqual(Path(directory).resolve(),
                         (Path.cwd() / "packages" / "com-support").resolve())
        self.assertEqual(com.provider().version, "9.9.9")

    @WINDOWS_ONLY
    def test_env_path_wins(self):
        custom = write_package(Path("elsewhere"), version="8.8.8")
        os.environ[com.ENV_PATH] = str(custom)
        directory, origin = com.resolve()
        self.assertEqual(origin, "env")
        self.assertEqual(Path(directory).resolve(), custom.resolve())

    def test_disabled_by_env(self):
        os.environ[com.ENV_DISABLE] = "0"
        directory, reason = com.resolve()
        self.assertIsNone(directory)
        self.assertIn("disabled", reason)
        self.assertIsNone(com.provider())

    def test_disabled_values_are_flexible(self):
        for value in ("off", "FALSE", "no", "disable", "disabled"):
            with self.subTest(value=value):
                os.environ[com.ENV_DISABLE] = value
                self.assertIsNone(com.provider())
        os.environ[com.ENV_DISABLE] = "1"
        # "1" is not a disable value: only assert it is not treated as "off"; no pywin32 required
        # 中文："1" 不是关闭值：只断言它没被当成"关闭"，不要求本机装了 pywin32
        # (otherwise this reports a false failure where pywin32 is missing).
        # 中文：（否则没装 pywin32 的解释器上这条会误报失败）。
        self.assertNotIn("disabled by", com.describe())

    def test_non_windows_platform_is_gated(self):
        with mock.patch("OpenVBSPP.ovpprt.com.sys.platform", "linux"):
            directory, reason = com.resolve()
            self.assertIsNone(directory)
            self.assertIn("Windows", reason)
            self.assertIsNone(com.provider())

    def test_missing_pywin32_is_not_available(self):
        write_package(Path("packages") / "com-support",
                      source=UNOBSERVABLE_SUPPORT, version="0.0.2")
        self.assertIsNone(com.provider(), "is_available() 为假时应当没有 provider")

    def test_describe_mentions_state(self):
        text = com.describe()
        self.assertIsInstance(text, str)
        self.assertTrue(text)


class TestCreateObject(ComTestCase):
    def test_unknown_progid_without_com_mentions_com(self):
        os.environ[com.ENV_DISABLE] = "0"
        with self.assertRaises(OVPPObjectNotSupportedError) as ctx:
            out('Set o = CreateObject("Excel.Application")')
        message = str(ctx.exception)
        self.assertIn("COM", message)
        self.assertIn(com.ENV_DISABLE, message)

    @WINDOWS_ONLY
    def test_script_gets_provider_object_without_import(self):
        write_package(Path("packages") / "com-support")
        code = ('Set o = CreateObject("Fake.Thing")\n'
                'echo o.progid\n'
                'o.value = 41\n'
                'echo o.value + 1\n'
                'echo TypeName(o)')
        self.assertEqual(out(code), "Fake.Thing\n42\nObject\n")

    def test_builtin_classes_still_win(self):
        write_package(Path("packages") / "com-support")
        code = ('Set d = CreateObject("Scripting.Dictionary")\n'
                'd("a") = 1\n'
                'echo d("a")')
        self.assertEqual(out(code), "1\n")

    @WINDOWS_ONLY
    def test_provider_errors_are_translated(self):
        write_package(Path("packages") / "com-support",
                      source=BROKEN_SUPPORT, version="0.0.1")
        provider = com.provider()
        self.assertIsNotNone(provider)
        with self.assertRaises(OVPPObjectNotSupportedError) as ctx:
            provider.create("Fake.Thing")
        message = str(ctx.exception)
        self.assertIn("COM", message)
        self.assertIn("not registered", message)

    @WINDOWS_ONLY
    def test_get_object_uses_provider(self):
        write_package(Path("packages") / "com-support")
        provider = com.provider()
        self.assertEqual(provider.get("Fake.Thing").progid, "Fake.Thing")


@WINDOWS_ONLY
class TestErrorNumbers(ComTestCase):
    """`Err.Number` must match the codes a real VBScript host reports.

    中文：`Err.Number` 必须与真实 VBScript 宿主报的编号一致。
    """

    def setUp(self):
        super().setUp()
        write_package(Path("packages") / "com-support",
                      source=NAMED_ERROR_SUPPORT, version="0.0.3")
        self.provider = com.provider()

    def test_create_failure_keeps_429(self):
        with self.assertRaises(OVPPAutomationError) as ctx:
            self.provider.create("No.Such.Class.Anyway")
        self.assertEqual(ctx.exception.number, 429)
        self.assertEqual(ctx.exception.name,
                         "ActiveX component can't create object")

    def test_get_passes_pathname_and_class(self):
        obj = self.provider.get("", "WScript.Shell")
        self.assertEqual(obj.progid, "WScript.Shell")

    def test_get_passes_pathname_and_classname(self):
        obj = self.provider.get("winmgmts:", "SWbemLocator")
        self.assertEqual(obj.progid, "SWbemLocator")


class TestAdapter(unittest.TestCase):
    """Load the vendor adapter directly to check proxy behaviour (pywin32 not needed).

    中文：直接加载 vendor 里的适配层，验证代理行为（不需要 pywin32）。"""

    def setUp(self):
        spec = importlib.util.spec_from_file_location(
            "ovpp_com_support_adapter_test", VENDOR / "com_support.py")
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        self.addCleanup(sys.modules.pop, spec.name, None)
        spec.loader.exec_module(module)
        self.adapter = module

    def test_proxy_property_read(self):
        proxy = self.adapter.COMProxy(FakeDispatch(), "Fake.Thing")
        self.assertEqual(proxy.Name(), "demo")

    def test_proxy_method_call(self):
        dispatch = FakeDispatch()
        proxy = self.adapter.COMProxy(dispatch, "Fake.Thing")
        self.assertEqual(proxy.Do(1, 2), "Do(1, 2)")
        self.assertEqual(dispatch.calls, [("Do", (1, 2))])

    def test_property_used_as_method_is_rejected(self):
        proxy = self.adapter.COMProxy(FakeDispatch(), "Fake.Thing")
        with self.assertRaises(self.adapter.ComSupportError) as ctx:
            proxy.Name(1)
        self.assertIn("property", str(ctx.exception))

    def test_object_valued_property_is_wrapped_not_called(self):
        # `folder.Files` returns an object; calling it with no arguments used to
        # invoke the property's default member and raise a type mismatch.
        # 中文：`folder.Files` 返回的是对象；过去无参调用会误触默认成员并报类型不匹配。
        inner = FakeDispatch()
        inner.values = {"Count": 3}
        dispatch = FakeDispatch()
        dispatch.values["Files"] = inner
        proxy = self.adapter.COMProxy(dispatch, "Fake.Thing")
        files = proxy.Files()
        self.assertIsInstance(files, self.adapter.COMProxy)
        self.assertEqual(files.Count(), 3)
        self.assertEqual(dispatch.calls, [])

    def test_proxy_write_and_collection_behaviour(self):
        dispatch = FakeDispatch()
        proxy = self.adapter.COMProxy(dispatch, "Fake.Thing")
        proxy.Text = "hi"
        self.assertEqual(dispatch.Text, "hi")
        self.assertEqual(len(proxy), 2)
        self.assertEqual([item for item in proxy], [0, 1])
        self.assertEqual(proxy[1], 1)
        self.assertEqual(repr(proxy), "<COM Fake.Thing>")
        self.assertEqual(proxy._vbs_name, "Object")

    def test_create_requires_windows(self):
        with mock.patch("OpenVBSPP.ovpprt.com.sys.platform", "linux"):
            with self.assertRaises(self.adapter.ComSupportError) as ctx:
                self.adapter.create("Fake.Thing")
        self.assertIn("Windows", str(ctx.exception))

    def test_create_rejects_remote_activation(self):
        with self.assertRaises(self.adapter.ComSupportError):
            self.adapter.create("Fake.Thing", "server")


class RealComTest(unittest.TestCase):
    """End-to-end checks against the real COM runtime (Windows + pywin32 only).

    中文：对真实 COM 运行时做端到端检查（仅 Windows + pywin32）。
    """

    @classmethod
    def setUpClass(cls):
        try:
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
        except Exception:  # pragma: no cover - platform dependent
            raise unittest.SkipTest("pywin32 is not installed")

    def test_missing_progid_reports_429(self):
        self.assertEqual(
            out('On Error Resume Next\n'
                'Set o = CreateObject("No.Such.Class.Anyway")\n'
                'echo Err.Number'),
            "429\n")

    def test_get_object_class_only_without_instance_reports_429(self):
        self.assertEqual(
            out('On Error Resume Next\n'
                'Set o = GetObject(, "No.Such.Class.Anyway")\n'
                'echo Err.Number'),
            "429\n")

    def test_get_object_moniker_returns_services(self):
        code = ('Set svc = GetObject("winmgmts:")\n'
                'echo TypeName(svc)')
        self.assertIn("Services", out(code))

    def test_native_folder_files_count(self):
        code = ('Set fso = CreateObject("Scripting.FileSystemObject")\n'
                'Set f = fso.GetFolder(".")\n'
                'echo f.Files.Count > 0')
        self.assertEqual(out(code), "True\n")


class PackagingSpecTest(unittest.TestCase):
    """exe 打包配置的守卫：漏一个运行时模块，单文件 exe 里就会 ModuleNotFoundError。

    `com-support` 是数据文件、`com_support.py` 里的 win32com 是延迟 import，
    两者都躲得过 PyInstaller 的静态分析，因此只能在 spec 里显式声明。
    """

    def setUp(self):
        self.spec_path = ROOT / "packaging" / "openvbs.spec"
        self.spec = self.spec_path.read_text("utf-8")

    def test_all_runtime_modules_are_declared(self):
        ovpprt = ROOT / "OpenVBSPP" / "ovpprt"
        modules = sorted(p.stem for p in ovpprt.glob("*.py")
                         if p.stem != "__init__")
        missing = [name for name in modules
                   if f'"OpenVBSPP.ovpprt.{name}"' not in self.spec]
        self.assertEqual(missing, [],
                         f"packaging/openvbs.spec 的 hiddenimports 少了: {missing}")
        self.assertIn('"OpenVBSPP.over"', self.spec)

    def test_com_support_ships_as_data(self):
        self.assertIn("vendor", self.spec)
        self.assertIn("com-support", self.spec)
        self.assertIn("datas=[", self.spec)

    def test_pywin32_is_bundled_when_available(self):
        self.assertIn("pythoncom", self.spec)
        self.assertIn("pywintypes", self.spec)
        self.assertIn("pywin32_system32", self.spec)


class FakeDispatch:
    """A fake IDispatch dynamic object standing in for win32com.

    中文：冒充 win32com 的 IDispatch 动态对象。"""

    def __init__(self):
        self.calls = []
        self.values = {"Name": "demo"}

    def __getattr__(self, name):
        if name in self.values:
            return self.values[name]

        def method(*args):
            self.calls.append((name, args))
            return f"{name}({', '.join(str(arg) for arg in args)})"

        return method

    def __iter__(self):
        return iter([0, 1])

    def __len__(self):
        return 2

    def __getitem__(self, key):
        return key


if __name__ == "__main__":
    unittest.main()
