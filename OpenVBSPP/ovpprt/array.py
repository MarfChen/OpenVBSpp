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

VBS 数组支持。

VBS 的 `Dim a(5)` 是**下界 0、上界 5**，共 6 个元素，
所以编译器传给 make() 的是"元素个数"（5 + 1）。
多维数组用嵌套 list 表示，`a(i, j)` 编译成 `a[i][j]`。
'''

from OpenVBSPP.ovpprt.empty import EMPTY


def make(dims):
    """按元素个数创建数组（dims[0] 是第 1 维的元素个数）。"""
    dims = [int(d) for d in dims]
    if not dims:
        return []
    if len(dims) == 1:
        return [EMPTY] * dims[0]
    return [make(dims[1:]) for _ in range(dims[0])]


def resize(array, dims):
    """ReDim [Preserve]：保留已有元素，按新的元素个数重建。"""
    dims = [int(d) for d in dims]
    if not isinstance(array, list):
        return make(dims)

    if not dims:
        return []

    if len(dims) == 1:
        new = [EMPTY] * dims[0]
        for i in range(min(len(array), dims[0])):
            new[i] = array[i]
        return new

    new = []
    for i in range(dims[0]):
        if i < len(array):
            new.append(resize(array[i], dims[1:]))
        else:
            new.append(make(dims[1:]))
    return new


def ubound(array, dimension=1):
    """UBound(arr[, dim])，1-based 维度参数。"""
    dim = int(dimension)
    node = array
    for _ in range(max(1, dim) - 1):
        node = node[0]
    return len(node) - 1


def lbound(array, dimension=1):
    """LBound 恒为 0（VBS 默认下界）。"""
    return 0


def erase(array):
    """Erase：把数组元素重置为 Empty（保持形状）。"""
    if not isinstance(array, list):
        return []
    if not array:
        return []
    if isinstance(array[0], list):
        return [erase(row) for row in array]
    return [EMPTY] * len(array)


def for_seq(start, end, step):
    """`For i = start To end Step step` 的可迭代序列（跳过循环变量赋值）。

    整数情形直接给 range，Python 的 for 比"while + 比较 + 自增"快得多；
    VBS 允许浮点步长，那种情况退化成列表。
    """
    try:
        return range(start, end + (1 if step > 0 else -1), step)
    except TypeError:
        values = []
        current = start
        if step >= 0:
            while current <= end:
                values.append(current)
                current += step
        else:
            while current >= end:
                values.append(current)
                current += step
        return values
