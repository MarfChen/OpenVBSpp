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


from OpenVBSPP.ovpprt.format import to_text


def ovpp_str_concat(a, b):
    """The `&` operator: concatenate using VBScript text rules (5.0 becomes "5",
    Empty becomes "").

    When both sides are already `str` they are simply added, which is the common
    case in `s = s & "x"` loops and skips two `to_text` conversions.

    中文：`&` 运算符按 VBS 的文本规则拼接（5.0 拼成 "5"，Empty 拼成空串）。
    When both sides are strings it just adds them; `s = s & "x"` in a loop is the common case,
    中文：两边都是字符串时直接相加——`s = s & "x"` 这类循环里这是主流写法，
    and skipping both text conversions pays off.
    中文：可以省掉两次文本转换。
    """
    if type(a) is str and type(b) is str:
        return a + b
    return to_text(a) + to_text(b)


_ORDER = {
    "EQ": lambda value: value == 0,
    "NE": lambda value: value != 0,
    "LT": lambda value: value < 0,
    "GT": lambda value: value > 0,
    "LE": lambda value: value <= 0,
    "GE": lambda value: value >= 0,
}


def vb_compare(op, a, b):
    """Comparison under `Option Compare Text`: strings compare case-insensitively,
    every other value compares as usual.

    中文：`Option Compare Text` 下的比较——字符串不区分大小写，其余值照常比较。
    """
    if isinstance(a, str) and isinstance(b, str):
        left = a.lower()
        right = b.lower()
    else:
        left = a
        right = b

    if left == right:
        order = 0
    elif left < right:
        order = -1
    else:
        order = 1

    return _ORDER[op](order)