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

包管理器接口（**预留**）。

现在只定义契约，不实现安装逻辑：IDE、CLI、LSP 都通过这里查询能力，
以后 `over` 包管理器落地时，只需要替换 `LocalPackageStore` 与 `resolve()`。

设计目标：
  * IDE 现在就能把"包"入口摆出来，但显示"未启用"；
  * 包管理器实现后，只要注册一个 `Provider`，前端代码不用改；
  * 不依赖网络、不引入第三方库。
'''

from dataclasses import dataclass, field


@dataclass
class PackageInfo:
    """一个可安装的扩展包（将来对应 PyPI 上的 wheel 或 .ovppkg）。"""
    name: str
    version: str = "0.0.0"
    description: str = ""
    provides: list = field(default_factory=list)   # 注册的过程/类/命令
    requires: list = field(default_factory=list)   # 依赖的包
    homepage: str = ""


class PackageProvider:
    """包源接口：本地目录 / 索引 / 远程仓库都实现这个协议。"""

    def search(self, keyword=""):
        raise NotImplementedError

    def install(self, name, version=None):
        raise NotImplementedError

    def uninstall(self, name):
        raise NotImplementedError

    def installed(self):
        return []


class LocalPackageStore(PackageProvider):
    """预留实现：从本地目录读包清单，不做安装。"""

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
    """以后的包管理器通过这里接管（IDE/CLI 不需要改代码）。"""
    global _PROVIDER
    _PROVIDER = provider


def get_provider():
    return _PROVIDER


def capabilities():
    """告诉前端：包管理器存在但未启用，以及它的契约长什么样。"""
    return {
        "enabled": False,
        "reason": "reserved for v0.9 (over package manager)",
        "provider": type(_PROVIDER).__name__,
        "operations": ["search", "install", "uninstall", "installed"],
        "package_fields": ["name", "version", "description",
                           "provides", "requires", "homepage"],
    }
