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

COM 桥接：把 over 包 `com-support` 接到对象层上。

脚本里 **不需要** import —— `CreateObject` 在遇到未知 ProgID 时会自动
问这里要一个 COM 对象。查找顺序：

  1. 环境变量 `OVPP_COM_PATH` 指向的目录（测试 / 高级用法）；
  2. 项目里的 `./packages/com-support/`（显式装过就以它为准，可用来钉版本）；
  3. 随包分发的 vendor 副本（Windows 版自带，零配置可用）。

因此 Windows 上**开箱即用**；想换实现或钉版本，就用 over 装一份到项目里；
`openvbs over remove com-support` 会退回随包副本；
设置 `OPENVBS_COM=0` 可以彻底关闭 COM 支持。
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

#: 已加载的 com_support 模块：{解析后的路径: (mtime, module)}
_MODULES = {}


class ComBridgeError(OVPPObjectNotSupportedError):
    """COM 支持不可用 / 加载失败。"""


def _bundled_dir():
    """随包分发的副本：OpenVBSPP/vendor/com-support/。"""
    return Path(__file__).resolve().parent.parent / "vendor" / PACKAGE_NAME


def _project_dir(start=None):
    """从 start（默认 cwd）向上找一个带 over.json 的项目根。"""
    current = Path(start or Path.cwd()).resolve()
    for candidate in (current, *current.parents):
        if (candidate / MANIFEST_NAME).is_file():
            return candidate
    return None


def _disabled_by_env():
    value = os.environ.get(ENV_DISABLE)
    return value is not None and str(value).strip().lower() in _DISABLED_VALUES


def candidates():
    """返回 [(目录, 来源标签)]，按优先级排列。"""
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
    """返回实际生效的 com-support 目录：(目录, 来源) 或 (None, 原因)。"""
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
    """从目录加载 com_support.py（按 mtime 缓存）。"""
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
    """清空已加载模块缓存（测试用）。"""
    for name in [item for item in sys.modules
                 if item.startswith("openvbspp_com_support_")]:
        sys.modules.pop(name, None)
    _MODULES.clear()


class Provider:
    """包着一个已加载的 com_support 模块。"""

    def __init__(self, directory, origin, module):
        self.directory = Path(directory)
        self.origin = origin
        self.module = module

    @property
    def version(self):
        return getattr(self.module, "VERSION", "")

    def create(self, progid, args=()):
        """CreateObject(progid) 的真实实现。"""
        try:
            return self.module.create(str(progid), *args)
        except OVPPObjectNotSupportedError:
            raise
        except Exception as exc:
            raise self._translate(progid, exc, "create") from exc

    def get(self, progid, args=()):
        """GetObject(progid)：拿已运行的实例。"""
        getter = getattr(self.module, "get", None)
        if getter is None:
            raise ComBridgeError(
                f"GetObject('{progid}') is not supported by com-support "
                f"{self.version or '(unknown version)'}")
        try:
            return getter(str(progid), *args)
        except OVPPObjectNotSupportedError:
            raise
        except Exception as exc:
            raise self._translate(progid, exc, "get") from exc

    def _translate(self, progid, exc, action):
        message = str(exc).strip() or exc.__class__.__name__
        text = f"Cannot {action} COM object '{progid}': {message}"
        if isinstance(exc, getattr(self.module, "ComSupportError", ())):
            return ComBridgeError(text)
        return OVPPAutomationError(text)

    def __repr__(self):
        return f"<COM provider {self.origin}: {self.directory}>"


def provider():
    """拿到 COM provider；不可用时返回 None。"""
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
    """给诊断/CLI 用的一句话说明。"""
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
    """对象层调用：创建一个 COM 对象，失败时抛出带 COM 字样的错误。"""
    active = provider()
    if active is None:
        raise OVPPObjectNotSupportedError(_unavailable_message(progid))
    return active.create(progid, args)


def _unavailable_message(progid):
    _directory, reason = resolve()
    return (f"Cannot create object '{progid}': COM objects are not available "
            f"in this environment ({reason})")
