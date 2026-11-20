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
'''


class VBool(int):
    """A VBScript-style boolean.

    Numerically True is -1 and False is 0 (the VBScript convention), so
    `(1 < 2) + 0` gives -1, while printing still shows True / False, which keeps
    `echo` output readable. `VBool` subclasses `int`, so it takes part in
    arithmetic, comparison and concatenation directly.

    中文：VBS 风格的布尔值。数值上 True = -1、False = 0（VBScript 的约定），
    so `(1 < 2) + 0` yields -1, while printing still shows True / False and `echo` output
    中文：所以 `(1 < 2) + 0` 得到 -1；但打印出来仍然是 True / False，`echo` 的输出
    reads more naturally. `VBool` subclasses int, so arithmetic, comparison and concatenation all work.
    中文：更自然。`VBool` 是 int 的子类，可直接参与算术、比较与字符串拼接。
    """

    def __new__(cls, value=False):
        return super().__new__(cls, -1 if value else 0)

    def __repr__(self):
        return "True" if int(self) != 0 else "False"

    __str__ = __repr__

    def __format__(self, spec):
        return format(str(self), spec)


TRUE = VBool(True)
FALSE = VBool(False)


def to_bool(value):
    """Convert any value to a VBScript boolean (-1 / 0).

    中文：把任意值转成 VBS 布尔（-1 / 0）。
    """
    return TRUE if value else FALSE


# VBScript logical operators are bitwise: they do bit operations on integers and stay
# boolean on booleans. Because True is -1 and False is 0, both readings agree.
# 中文：VBS 的逻辑运算符是按位的——对整数做位运算、对布尔保持布尔；
# True is -1 and False is 0, so both meanings agree by construction.
# 中文：因为 True 就是 -1、False 就是 0，这两套语义天然统一。
def _both_boolean(a, b):
    return isinstance(a, VBool) and isinstance(b, VBool)


def _logical(a, b, op):
    left = int(a)
    right = int(b)
    result = op(left, right)
    return to_bool(result) if _both_boolean(a, b) else result


def ovpp_and(a, b):
    return _logical(a, b, lambda x, y: x & y)


def ovpp_or(a, b):
    return _logical(a, b, lambda x, y: x | y)


def ovpp_xor(a, b):
    return _logical(a, b, lambda x, y: x ^ y)


def ovpp_eqv(a, b):
    return _logical(a, b, lambda x, y: ~(x ^ y))


def ovpp_imp(a, b):
    return _logical(a, b, lambda x, y: (~x) | y)


def ovpp_not(a):
    result = ~int(a)
    return to_bool(result) if isinstance(a, VBool) else result
