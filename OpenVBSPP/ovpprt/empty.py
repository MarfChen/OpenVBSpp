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

VBScript `Empty`.

After `Dim x`, x is Empty: it prints as an empty line, counts as 0 in arithmetic
and as "" when concatenated, exactly like VBScript.

中文：VBS 的 Empty 值。`Dim x` 之后 x 就是 Empty——打印为空行、参与算术时当 0、
It joins a string concatenation as an empty string, matching VBScript.
中文：参与字符串连接时当空串，与 VBScript 行为一致。
'''


class EmptyType:
    """The empty-value sentinel: 0 in arithmetic, "" as text.

    中文：空值哨兵——数值上等价于 0，字符串上等价于空串。
    """

    __slots__ = ()

    # Display / 中文：显示
    def __str__(self):
        return ""

    def __repr__(self):
        return "Empty"

    # Conversion / 中文：转换
    def __bool__(self):
        return False

    def __int__(self):
        return 0

    def __float__(self):
        return 0.0

    def __index__(self):
        return 0

    def __hash__(self):
        return hash(0)

    # Comparison / 中文：比较
    def __eq__(self, other):
        if other is self:
            return True
        if isinstance(other, bool):
            return not other
        if isinstance(other, (int, float)):
            return other == 0
        if isinstance(other, str):
            return other == ""
        return NotImplemented

    def __ne__(self, other):
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result

    def __lt__(self, other):
        return 0 < _as_number(other)

    def __le__(self, other):
        return 0 <= _as_number(other)

    def __gt__(self, other):
        return 0 > _as_number(other)

    def __ge__(self, other):
        return 0 >= _as_number(other)

    # Arithmetic: Empty counts as 0 in the four basic operations.
    # 中文：算术——Empty 参与四则运算时按 0 处理。
    def __add__(self, other):
        return 0 + other

    def __radd__(self, other):
        return other + 0

    def __sub__(self, other):
        return 0 - other

    def __rsub__(self, other):
        return other - 0

    def __mul__(self, other):
        return 0 * other

    def __rmul__(self, other):
        return other * 0

    def __truediv__(self, other):
        return 0 / other

    def __rtruediv__(self, other):
        return other / 0

    def __floordiv__(self, other):
        return 0 // other

    def __rfloordiv__(self, other):
        return other // 0

    def __mod__(self, other):
        return 0 % other

    def __rmod__(self, other):
        return other % 0

    def __neg__(self):
        return 0

    def __pos__(self):
        return 0

    def __abs__(self):
        return 0


def _as_number(value):
    if value is EMPTY:
        return 0
    return value


EMPTY = EmptyType()
