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

VBS 内置集合：Scripting.Dictionary。

VBS 里 `d("a")` 走的是默认成员 Item，这里直接用映射协议实现，
所以 `d("a")` / `d("a") = 1` / `For Each k In d` 都能自然工作。
'''

from OpenVBSPP.exceptions import OVPPError


class Dictionary:
    """Scripting.Dictionary 的最小实现（键区分大小写，与 VBS 默认一致）。"""

    _vbs_name = "Dictionary"

    def __init__(self, *args, **kwargs):
        self._data = {}

    # ------------------------------------------------------------ VBS 方法
    @property
    def count(self):
        return len(self._data)

    def add(self, key, value):
        if key in self._data:
            raise OVPPError(f"Dictionary: key '{key}' already exists")
        self._data[key] = value

    def exists(self, key):
        return key in self._data

    def item(self, key):
        if key not in self._data:
            raise KeyError(key)
        return self._data[key]

    def remove(self, key):
        if key in self._data:
            del self._data[key]

    def removeall(self):
        self._data.clear()

    def keys(self):
        return list(self._data.keys())

    def items(self):
        return list(self._data.values())

    # ------------------------------------------------------------ 映射协议
    def __getitem__(self, key):
        return self._data[key]

    def __setitem__(self, key, value):
        self._data[key] = value

    def __delitem__(self, key):
        del self._data[key]

    def __contains__(self, key):
        return key in self._data

    def __iter__(self):
        # VBS 的 For Each 遍历键
        return iter(list(self._data.keys()))

    def __len__(self):
        return len(self._data)

    def __repr__(self):
        return f"Dictionary({self._data!r})"


def make_dictionary(*args):
    return Dictionary()
