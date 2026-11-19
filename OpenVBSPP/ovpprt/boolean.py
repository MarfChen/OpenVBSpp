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
'''


class VBool(int):
    """VBS 风格的布尔值。

    数值上 True = -1、False = 0（VBScript 的约定），所以
    `(1 < 2) + 0` 得到 -1；但它打印出来仍然是 True / False，
    这样 echo 的输出对初学者更自然。

    VBool 是 int 的子类，可以直接参与算术、比较和字符串拼接。
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
    """把任意值转成 VBS 布尔（-1 / 0）。"""
    return TRUE if value else FALSE


# VBS 的逻辑运算符是**按位**的：对整数做位运算，对布尔保持布尔。
# 由于 True 就是 -1、False 就是 0，这两套语义天然统一。
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
