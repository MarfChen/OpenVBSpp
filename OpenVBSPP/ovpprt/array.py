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

VBScript arrays.

`Dim a(5)` has lower bound 0 and upper bound 5, six elements in total, so the
compiler passes the element count (5 + 1) to `make()`. Multi-dimensional arrays
are nested lists, and `a(i, j)` compiles to `a[i][j]`.

中文：VBS 数组支持。`Dim a(5)` 的下界是 0、上界是 5，共 6 个元素，所以编译器传给
`make()` takes the element count (5 + 1). A multi-dimensional array is a nested list, so `a(i, j)` compiles to
中文：`make()` 的是元素个数（5 + 1）。多维数组用嵌套 list 表示，`a(i, j)` 编译成
`a[i][j]`。
'''

from OpenVBSPP.ovpprt.empty import EMPTY


def check_index(index):
    """Reject negative array indices.

    中文：拒绝负数数组下标。
    """
    if isinstance(index, int) and index < 0:
        raise IndexError("Subscript out of range")
    return index


def get_item(value, index):
    """Read an item, checking array bounds.

    中文：读取元素，并检查数组下标。
    """
    if isinstance(value, list):
        index = check_index(index)
    return value[index]


def checked_target(value):
    """Guard array subscripts while preserving other objects.

    中文：检查数组下标，同时保留其他对象的行为。
    """
    if isinstance(value, list):
        return _CheckedArrayTarget(value)
    return value


class _CheckedArrayTarget:
    def __init__(self, value):
        self.value = value

    def __setitem__(self, index, value):
        self.value[check_index(index)] = value


def make(dims):
    """Create an array from element counts (`dims[0]` is the length of dimension 1).

    中文：按元素个数创建数组（`dims[0]` 是第 1 维的元素个数）。
    """
    dims = [int(d) for d in dims]
    if not dims:
        return []
    if len(dims) == 1:
        return [EMPTY] * dims[0]
    return [make(dims[1:]) for _ in range(dims[0])]


def resize(array, dims):
    """`ReDim [Preserve]`: rebuild the array with new lengths, keeping existing values.

    中文：`ReDim [Preserve]`——保留已有元素，按新的元素个数重建数组。
    """
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
    """`UBound(arr[, dim])`; the dimension argument is 1-based.

    中文：`UBound(arr[, dim])`，维度参数从 1 开始。
    """
    dim = int(dimension)
    node = array
    for _ in range(max(1, dim) - 1):
        node = node[0]
    return len(node) - 1


def lbound(array, dimension=1):
    """`LBound` is always 0 (the VBScript default lower bound).

    中文：`LBound` 恒为 0（VBS 的默认下界）。
    """
    return 0


def erase(array):
    """`Erase`: reset every element to Empty while keeping the shape.

    中文：`Erase`——把数组元素重置为 Empty（保持形状）。
    """
    if not isinstance(array, list):
        return []
    if not array:
        return []
    if isinstance(array[0], list):
        return [erase(row) for row in array]
    return [EMPTY] * len(array)


def for_seq(start, end, step):
    """The iterable behind `For i = start To end Step step`.

    Integer bounds return a `range`, because a Python `for` beats "while + compare
    + increment" by a wide margin; VBScript also allows a floating-point step, in
    which case this falls back to a list.

    中文：`For i = start To end Step step` 的可迭代序列（省掉循环变量赋值）。
    For integers it returns a `range`, since a Python for loop beats while + compare + increment by a wide margin;
    中文：整数情形直接给 `range`——Python 的 for 比「while + 比较 + 自增」快得多；
    VBScript also allows a floating-point step, and that case degrades to a list.
    中文：VBS 允许浮点步长，那种情况退化成列表。
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
