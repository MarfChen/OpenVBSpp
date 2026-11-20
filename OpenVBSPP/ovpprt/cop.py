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

The common object pool (COP).

Once Python interop is imported (`<script src="python-embed" />`), a script can use
`COP` as a global dictionary; Python code gets **the same** object with
`from OpenVBSPP.ovpprt.cop import COP`, so both sides can exchange arbitrary Python
objects.

中文：公共对象池（COP）。用 `<script src="python-embed" />` 导入 Python 互操作后，
A script reaches one global dictionary through the name `COP`; Python code obtains the
中文：VBS 侧可以用 `COP` 这个名字访问一个全局字典；Python 侧用
same object with `from OpenVBSPP.ovpprt.cop import COP`, so both sides can exchange any
中文：`from OpenVBSPP.ovpprt.cop import COP` 拿到同一个对象，于是两边可以交换任意 Python
so both sides exchange any Python object.
中文：于是两边可以交换任意 Python 对象。
'''

from OpenVBSPP.ovpprt.collections import Dictionary


COP = Dictionary()


def reset():
    """Empty the pool (used by tests).

    中文：清空对象池（测试用）。
    """
    COP.removeall()
