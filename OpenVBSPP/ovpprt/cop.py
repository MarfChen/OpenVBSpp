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

公共对象池（COP）。

`Option COP` 打开后，VBS 侧可以用 `COP` 这个名字访问一个全局字典；
Python 侧用 `from OpenVBSPP.ovpprt.cop import COP` 拿到**同一个**对象，
于是两边可以交换任意 Python 对象。
'''

from OpenVBSPP.ovpprt.collections import Dictionary


COP = Dictionary()


def reset():
    """清空对象池（测试用）。"""
    COP.removeall()
