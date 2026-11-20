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

The package manager interface. Nothing is installed here yet: it only defines the
contract that the IDE, the CLI and a future language server query.

Design goals:
  * the IDE can already show a packages entry point, marked as not enabled;
  * once a provider is registered, frontend code needs no change;
  * no network access and no third-party libraries.

The real work is delegated: `LocalPackageStore` plus `resolve()` are the two swap points.

中文：包管理器接口（预留）。现在只定义契约、不实现安装逻辑：IDE / CLI / LSP 都通过
The IDE asks for capabilities here; a real implementation only has to replace `LocalPackageStore` and `resolve()`. The goals are to give the IDE
中文：这里查询能力，将来落地时只需要替换 `LocalPackageStore` 与 `resolve()`。目标是让 IDE
a packages entry today (shown as disabled), to need no front-end change once a package manager registers a `Provider`,
中文：现在就能摆出「包」入口（显示未启用）、包管理器实现后只要注册一个 `Provider` 前端不用改，
and to stay offline with no third-party dependency.
中文：且不依赖网络、不引入第三方库。
'''

from dataclasses import dataclass, field


@dataclass
class PackageInfo:
    """One installable package (a wheel or a .ovppkg later on).

    中文：一个可安装的扩展包（将来对应 PyPI 上的 wheel 或 .ovppkg）。
    """
    name: str
    version: str = "0.0.0"
    description: str = ""
    # Procedures, classes and commands it registers.
    # 中文：注册的过程 / 类 / 命令。
    provides: list = field(default_factory=list)
    # Packages it depends on.
    # 中文：依赖的包。
    requires: list = field(default_factory=list)
    homepage: str = ""


class PackageProvider:
    """The package source protocol, implemented by local directories, indexes and remotes.

    中文：包源接口：本地目录 / 索引 / 远程仓库都实现这个协议。
    """

    def search(self, keyword=""):
        raise NotImplementedError

    def install(self, name, version=None):
        raise NotImplementedError

    def uninstall(self, name):
        raise NotImplementedError

    def installed(self):
        return []


class LocalPackageStore(PackageProvider):
    """A placeholder: read manifests from a local directory, installing nothing.

    中文：预留实现：从本地目录读包清单，不做安装。
    """

    def __init__(self, root=None):
        self.root = root

    def search(self, keyword=""):
        return []

    def install(self, name, version=None):
        raise NotImplementedError(
            "package installation is reserved for a future release; "
            "see doc/WEB_IDE.md §5")

    def uninstall(self, name):
        raise NotImplementedError("package removal is reserved for a future release")

    def installed(self):
        return []


_PROVIDER = LocalPackageStore()


def set_provider(provider):
    """The future package manager plugs in here, with no change to the IDE or CLI.

    中文：以后的包管理器通过这里接管（IDE/CLI 不需要改代码）。
    """
    global _PROVIDER
    _PROVIDER = provider


def get_provider():
    return _PROVIDER


def capabilities():
    """Tell a frontend the package manager exists but is not enabled, and show its contract.

    中文：告诉前端：包管理器存在但未启用，以及它的契约长什么样。
    """
    return {
        "enabled": False,
        "reason": "reserved for v0.9 (over package manager)",
        "provider": type(_PROVIDER).__name__,
        "operations": ["search", "install", "uninstall", "installed"],
        "package_fields": ["name", "version", "description",
                           "provides", "requires", "homepage"],
    }
