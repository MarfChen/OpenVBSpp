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

VBS 的 Empty 值。

`Dim x` 之后 x 是 Empty：打印为空行、参与算术时当 0、参与字符串连接时当 ""，
和 VBScript 的行为一致。
'''


class EmptyType:
    """空值哨兵。数值上等价于 0，字符串上等价于 ""。"""

    __slots__ = ()

    # ---------------------------------------------------------------- 显示
    def __str__(self):
        return ""

    def __repr__(self):
        return "Empty"

    # ---------------------------------------------------------------- 转换
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

    # ---------------------------------------------------------------- 比较
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

    # ---------------------------------------------------------------- 算术
    # Empty 参与四则运算时按 0 处理
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
