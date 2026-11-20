'''
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

COM bridge: connects the `com-support` over package to the object layer.

Scripts **never import anything**: `CreateObject` asks this module for an object
when it meets an unknown ProgID. Order of preference:

  1. the directory named by `OVPP_COM_PATH` (tests and advanced setups);
  2. `./packages/com-support/` inside the project (a project install wins, so it
     can pin a version or ship a different implementation);
  3. the copy bundled with the distribution (Windows wheels and the exe carry it,
     so it works with no setup).

That makes COM usable out of the box on Windows. `openvbs over remove com-support`
falls back to the bundled copy, and `OPENVBS_COM=0` turns COM support off.

中文：COM 桥接——把 over 包 `com-support` 接到对象层上。脚本里不需要 import：
CreateObject asks this module for a COM object when a ProgID is unknown. Lookup order:
中文：`CreateObject` 遇到未知 ProgID 时会自动问这里要一个 COM 对象。查找顺序是
the `OVPP_COM_PATH` environment variable, then `./packages/com-support/` in the project (an explicit install wins,
中文：`OVPP_COM_PATH` 环境变量 → 项目里的 `./packages/com-support/`（显式装过就以它为准，
so you can pin a version), then the vendor copy shipped with the package. It therefore works out of the box on Windows;
中文：可钉版本）→ 随包分发的 vendor 副本。因此 Windows 上开箱即用；
`openvbs over remove com-support` falls back to the bundled copy and `OPENVBS_COM=0` turns the bridge off entirely.
中文：`openvbs over remove com-support` 会退回随包副本；`OPENVBS_COM=0` 可以彻底关闭。
'''

import importlib.util
import os
import sys
from pathlib import Path

from OpenVBSPP.exceptions import OVPPAutomationError, OVPPObjectNotSupportedError

ENV_DISABLE = "OPENVBS_COM"
ENV_PATH = "OVPP_COM_PATH"
PACKAGE_NAME = "com-support"
MANIFEST_NAME = "over.json"
PACKAGES_DIR = "packages"

_DISABLED_VALUES = {"0", "off", "false", "no", "disable", "disabled"}

#: Loaded com_support modules: {resolved path: (mtime, module)}
#: 中文：已加载的 com_support 模块缓存，键是解析后的路径。
_MODULES = {}


class ComBridgeError(OVPPObjectNotSupportedError):
    """COM support is unavailable or failed to load.

    中文：COM 支持不可用或加载失败。
    """


def _bundled_dir():
    """The bundled copy: `OpenVBSPP/vendor/com-support/`.

    中文：随包分发的副本 `OpenVBSPP/vendor/com-support/`。
    """
    return Path(__file__).resolve().parent.parent / "vendor" / PACKAGE_NAME


def _project_dir(start=None):
    """Walk up from `start` (default: cwd) to the nearest project root with `over.json`.

    中文：从 `start`（默认当前目录）向上找一个带 `over.json` 的项目根。
    """
    current = Path(start or Path.cwd()).resolve()
    for candidate in (current, *current.parents):
        if (candidate / MANIFEST_NAME).is_file():
            return candidate
    return None


def _disabled_by_env():
    value = os.environ.get(ENV_DISABLE)
    return value is not None and str(value).strip().lower() in _DISABLED_VALUES


def candidates():
    """Return `[(directory, origin)]` in priority order.

    中文：按优先级返回 [(目录, 来源标签)]。
    """
    found = []
    explicit = os.environ.get(ENV_PATH)
    if explicit:
        found.append((Path(explicit), "env"))

    project = _project_dir()
    if project is not None:
        found.append((project / PACKAGES_DIR / PACKAGE_NAME, "project"))
    found.append((_bundled_dir(), "bundled"))
    return found


def resolve():
    """Return the effective com-support directory: `(directory, origin)` or
    `(None, reason)`.

    中文：返回实际生效的 com-support 目录 —— (目录, 来源) 或 (None, 原因)。
    """
    if _disabled_by_env():
        return None, f"disabled by {ENV_DISABLE}"
    if sys.platform != "win32":
        return None, f"COM is only available on Windows (platform: {sys.platform})"

    tried = []
    for directory, origin in candidates():
        script = Path(directory) / "com_support.py"
        if script.is_file():
            return Path(directory), origin
        tried.append(f"{directory} ({origin})")

    hint = ("the bundled copy is missing from this installation; "
            "run 'openvbs over add com-support file:<path>'")
    return None, f"com-support package not found; {hint}. Tried: " + ", ".join(tried)


def _load(directory):
    """Load `com_support.py` from a directory, cached by modification time.

    中文：从目录加载 `com_support.py`（按 mtime 缓存）。
    """
    script = Path(directory) / "com_support.py"
    key = str(script.resolve())
    try:
        stamp = script.stat().st_mtime
    except OSError:
        stamp = None

    cached = _MODULES.get(key)
    if cached is not None and cached[0] == stamp:
        return cached[1]

    module_name = "openvbspp_com_support_" + str(abs(hash(key)))
    spec = importlib.util.spec_from_file_location(module_name, key)
    if spec is None or spec.loader is None:
        raise ComBridgeError(f"Cannot load COM support from {key}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        sys.modules.pop(module_name, None)
        raise ComBridgeError(
            f"COM support failed to load from {key}: {exc}") from exc

    _MODULES[key] = (stamp, module)
    return module


def reset_cache():
    """Clear the loaded-module cache (used by tests).

    中文：清空已加载的模块缓存（测试用）。
    """
    for name in [item for item in sys.modules
                 if item.startswith("openvbspp_com_support_")]:
        sys.modules.pop(name, None)
    _MODULES.clear()


class Provider:
    """Wraps one loaded com_support module.

    中文：包着一个已加载的 com_support 模块。
    """

    def __init__(self, directory, origin, module):
        self.directory = Path(directory)
        self.origin = origin
        self.module = module

    @property
    def version(self):
        return getattr(self.module, "VERSION", "")

    def create(self, progid, args=()):
        """The implementation behind `CreateObject(progid)`.

        中文：`CreateObject(progid)` 的真实实现。
        """
        try:
            return self.module.create(str(progid), *args)
        except OVPPObjectNotSupportedError:
            raise
        except Exception as exc:
            raise self._translate(progid, exc, "create") from exc

    def get(self, pathname="", classname=None):
        """The implementation behind `GetObject([pathname], [class])`.

        中文：`GetObject([pathname], [class])` 的真实实现。
        """
        getter = getattr(self.module, "get", None)
        if getter is None:
            raise ComBridgeError(
                f"GetObject('{pathname}') is not supported by com-support "
                f"{self.version or '(unknown version)'}")
        try:
            return getter(str(pathname or ""), classname)
        except OVPPObjectNotSupportedError:
            raise
        except Exception as exc:
            raise self._translate(pathname or classname, exc, "get") from exc

    def _translate(self, progid, exc, action):
        message = str(exc).strip() or exc.__class__.__name__
        number = getattr(exc, "number", None)
        vbs_name = getattr(exc, "vbs_name", None)
        if number:
            error = OVPPAutomationError(message)
            error.number = int(number)
            if vbs_name:
                error.name = str(vbs_name)
            return error
        text = f"Cannot {action} COM object '{progid}': {message}"
        if isinstance(exc, getattr(self.module, "ComSupportError", ())):
            return ComBridgeError(text)
        return OVPPAutomationError(text)

    def __repr__(self):
        return f"<COM provider {self.origin}: {self.directory}>"


def provider():
    """Return the COM provider, or None when COM is unavailable.

    中文：拿到 COM provider；不可用时返回 None。
    """
    if _disabled_by_env() or sys.platform != "win32":
        return None
    directory, origin = resolve()
    if directory is None:
        return None
    try:
        module = _load(directory)
    except ComBridgeError:
        return None
    if hasattr(module, "is_available") and not module.is_available():
        return None
    return Provider(directory, origin, module)


def describe():
    """A one-line status for diagnostics and the CLI.

    中文：给诊断与 CLI 用的一句话状态说明。
    """
    directory, reason = resolve()
    if directory is None:
        return f"unavailable: {reason}"
    try:
        module = _load(directory)
    except ComBridgeError as exc:
        return f"broken: {exc}"
    available = getattr(module, "is_available", lambda: True)()
    state = "ready" if available else "pywin32 missing"
    return f"{state} ({reason}: {directory})"


def create(progid, args=()):
    """Called by the object layer: create a COM object, raising an error that names COM
    when that is impossible.

    中文：对象层调用——创建一个 COM 对象，失败时抛出带 COM 字样的错误。
    """
    active = provider()
    if active is None:
        raise OVPPObjectNotSupportedError(_unavailable_message(progid))
    return active.create(progid, args)


def get(pathname="", classname=None):
    """Called by the object layer for `GetObject([pathname], [class])`.

    中文：对象层调用——实现 `GetObject([pathname], [class])`。
    """
    active = provider()
    if active is None:
        raise OVPPObjectNotSupportedError(
            _unavailable_message(pathname or classname or "GetObject"))
    return active.get(pathname, classname)


def _unavailable_message(progid):
    _directory, reason = resolve()
    return (f"Cannot create object '{progid}': COM objects are not available "
            f"in this environment ({reason})")
