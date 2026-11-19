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

ByRef 参数的支持。

VBS 的过程参数默认是 ByRef（按引用传递）。Python 只有按值传递，
所以编译器把 ByRef 参数包进 Ref 盒子：

    Ref 创建：_r = Ref(x)
    过程内部：对参数 x 的读写都变成 x.v        （Attribute 访问）
    调用之后：x = _r.v                         （写回）

这样 `Sub Swap(ByRef a, ByRef b)` 这类代码才能生效。
'''

from OpenVBSPP.ovpprt.empty import EMPTY


class Ref:
    """一个可写的盒子，用来模拟 ByRef 参数。"""

    __slots__ = ("v",)

    def __init__(self, value=EMPTY):
        self.v = value

    def __repr__(self):
        return f"Ref({self.v!r})"
