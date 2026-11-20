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

The VBScript built-in functions.

中文：VBS 常用内置函数。

The compiler turns a call such as `Len(x)` into `_ovpp_builtin.Len(x)`,
中文：编译器把 `Len(x)` 这样的调用翻译成 `_ovpp_builtin.Len(x)`，
so names here keep VBScript spelling with a capital first letter.
中文：所以这里的函数名保持 VBS 的写法（首字母大写）。
REGISTRY keys are lower-case names, which gives VBScript's case insensitivity.
中文：REGISTRY 用小写名字做键，从而实现 VBS 的大小写不敏感。
'''

import sys

from OpenVBSPP.ovpprt.array import lbound, ubound
from OpenVBSPP.ovpprt.boolean import VBool, to_bool
from OpenVBSPP.ovpprt.empty import EMPTY, EmptyType
from OpenVBSPP.exceptions import OVPPInvalidUseOfNullError
from OpenVBSPP.ovpprt.format import to_text
from OpenVBSPP.ovpprt.null import NULL, NullType


class ScriptEnd(BaseException):
    """Raised by a bare `End` to unwind the whole script.

    中文：单独的 `End` 抛出此异常，用来终止整个脚本。

    It derives from BaseException on purpose: `On Error Resume Next` and the
    `Try`/`Catch` handlers only catch `Exception`, so nothing can swallow it.
    中文：它故意继承 BaseException ——`On Error Resume Next` 与 `Try`/`Catch`
    只捕获 `Exception`，因此谁也接不住它。
    """

    __slots__ = ()


class _LazyModule:
    """A module proxy: the target module is imported on first attribute access.

    中文：按需 import 的模块代理 —— 第一次取属性时才真正 import 目标模块。
    """

    __slots__ = ("_name", "_module")

    def __init__(self, name):
        self._name = name
        self._module = None

    def __getattr__(self, attr):
        module = self._module
        if module is None:
            import importlib

            module = self._module = importlib.import_module(self._name)
        return getattr(module, attr)

    def __repr__(self):
        return f"_LazyModule({self._name!r})"


_datetime = _LazyModule("datetime")
math = _LazyModule("math")
random = _LazyModule("random")
re = _LazyModule("re")


# --------------------------------------------------------------------- Helpers / 中文：工具
def _to_str(value):
    return to_text(value)


def _to_number(value):
    if value is EMPTY or isinstance(value, EmptyType):
        return 0
    if isinstance(value, bool):
        return -1 if value else 0
    if isinstance(value, (int, float)):
        return value
    text = _to_str(value).strip()
    try:
        return int(text)
    except ValueError:
        return float(text)


def _to_list(value):
    if isinstance(value, list):
        return value
    raise TypeError("Expected an array, got " + type(value).__name__)


def _is_null(value):
    """True when ``value`` is VBScript's Null sentinel.

    中文：判断是否为 VBS 的 Null 哨兵。
    """
    return value is NULL or isinstance(value, NullType)


def _is_empty(value):
    return value is EMPTY or isinstance(value, EmptyType)


def _number_of(value):
    """Convert one operand to a number the way VBScript does it.

    中文：按 VBScript 的规则把操作数转成数字。

    Returns ``(number, is_double)``: VBScript only promotes to Double when the
    value is a fractional number or a string, so the caller can keep Integer
    arithmetic for plain integer operands.
    中文：返回 ``(数值, 是否 Double)``；只有小数或字符串才提升为 Double，
    这样整数与整数之间仍能保持 Integer 语义。

    Empty counts as 0. A string must look numeric (`"  3 "` is fine, `""` is
    not); anything else is a type mismatch, error 13.
    中文：Empty 当作 0；字符串必须是数字（`"  3 "` 合法，`""` 不合法）；
    其余一律算类型不匹配（错误 13）。
    """
    if value is EMPTY or isinstance(value, EmptyType):
        return 0, False
    kind = type(value)
    if kind is int:
        return value, False
    if kind is float:
        return value, True
    if kind is str:
        if not value.strip():
            raise TypeError("Type mismatch: an empty string is not a number")
        try:
            return int(value), False
        except ValueError:
            pass
        try:
            return float(value), True
        except ValueError:
            raise TypeError(f"Type mismatch: {value!r} is not a number")
    if isinstance(value, bool):
        return (-1 if value else 0), False
    if isinstance(value, VBool):
        return int(value), False
    if isinstance(value, (int, float)):
        return value, isinstance(value, float)
    raise TypeError(
        "Type mismatch: " + type(value).__name__ + " cannot be used as a number")


def _arith_pair(a, b):
    """Coerce both operands of ``- * / ^`` to one common numeric type.

    中文：把 `- * / ^` 的两个操作数归一到同一种数值类型。
    """
    left, left_double = _number_of(a)
    right, right_double = _number_of(b)
    if left_double or right_double:
        return float(left), float(right)
    return left, right


def _is_text_operand(value):
    """Operands that make `+` concatenate instead of add.

    中文：会触发 `+` 拼接（而非加法）的操作数。
    """
    return type(value) is str or value is EMPTY or isinstance(value, EmptyType)


# --------------------------------------------------------------------- Arithmetic / 中文：算术
# VBScript's `+` concatenates only when *both* sides are strings; every other
# combination is numeric. `- * / ^` are always numeric, `/` and `^` always
# produce a Double, and `\` / `Mod` stay integral.
# 中文：VBS 的 `+` 只在「两侧都是字符串」时拼接，其余一律按数值相加；
# `- * / ^` 恒为数值，`/` 与 `^` 恒得到 Double，`\` 与 `Mod` 保持整数。
#
# With `Option Native` every arithmetic operator rejects text instead of
# converting it: `"1" + "2"`, `1 + "2"` and `"6" / "3"` raise error 13 rather
# than quietly becoming "12", 3 and 2. Empty is still the default value 0 and
# Null still propagates — neither is a conversion of an actual value.
# 中文：打开 `Option Native` 后，算术运算符拒绝文本而不是转换它：
# `"1" + "2"`、`1 + "2"`、`"6" / "3"` 都报错误 13，而不是悄悄变成 "12"、3、2。
# Empty 仍然当作默认值 0、Null 仍然传播 —— 这两者都不是对真实值的转换。
def _plain_number(value):
    """True for a Python int/float that needs no VBScript conversion.

    中文：判断是否是不需要做 VBS 转换的普通 int/float。

    `VBool` subclasses `int` but VBScript treats True as -1, so bool-shaped
    values must not take this fast path; `type(...) is` keeps the check exact.
    中文：`VBool` 继承自 `int`，但 VBS 里 True 是 -1，所以布尔形状的值不能走这条
    快路径；用 `type(...) is` 做精确判断。

    The arithmetic operators inline this test instead of calling it: it runs once
    per operation, so a function call here would double the cost.
    中文：算术运算符里把这段判断内联展开而不调用本函数 —— 它每次运算都会跑，
    多一次函数调用等于把开销翻倍。
    """
    kind = type(value)
    return kind is int or kind is float


def ovpp_add(a, b):
    """VBScript `+`: add, or concatenate when both operands are text.

    中文：VBS 的 `+` —— 相加；两侧都是文本时改为拼接。
    """
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float):
        return a + b
    if _is_null(a) or _is_null(b):
        return NULL
    if _is_text_operand(a) and _is_text_operand(b):
        if _is_empty(a) and _is_empty(b):
            return 0
        return to_text(a) + to_text(b)
    left, right = _arith_pair(a, b)
    return left + right


def ovpp_sub(a, b):
    """VBScript `-`. 中文：VBS 的减法。"""
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float):
        return a - b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _arith_pair(a, b)
    return left - right


def ovpp_mul(a, b):
    """VBScript `*`. 中文：VBS 的乘法。"""
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float):
        return a * b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _arith_pair(a, b)
    return left * right


def ovpp_div(a, b):
    """VBScript `/`: always a Double, and `0 / 0` is an Overflow (error 6).

    中文：VBS 的 `/` —— 结果恒为 Double，且 `0 / 0` 是溢出（错误 6）。
    """
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float) and b:
        return a / b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _arith_pair(a, b)
    if right == 0:
        if left == 0:
            raise OverflowError("Overflow")
        raise ZeroDivisionError("Division by zero")
    return float(left) / float(right)


def ovpp_pow(a, b):
    """VBScript `^`: always a Double; a negative base with a fractional
    exponent is "Invalid procedure call" (error 5).

    中文：VBS 的 `^` —— 结果恒为 Double；负底数配小数指数算非法过程调用（错误 5）。
    """
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float) and b >= 0:
        return float(a) ** b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _arith_pair(a, b)
    if left == 0 and right < 0:
        raise ValueError("Invalid procedure call or argument")
    try:
        result = float(left) ** float(right)
    except OverflowError:
        raise OverflowError("Overflow")
    if isinstance(result, complex):
        raise ValueError("Invalid procedure call or argument")
    return result


def ovpp_neg(a):
    """VBScript unary minus. 中文：VBS 的一元负号。"""
    kind = type(a)
    if kind is int or kind is float:
        return -a
    if _is_null(a):
        return NULL
    number, is_double = _number_of(a)
    return -float(number) if is_double else -number


def ovpp_pos(a):
    """VBScript unary plus. 中文：VBS 的一元正号。"""
    kind = type(a)
    if kind is int or kind is float:
        return a
    if _is_null(a):
        return NULL
    number, is_double = _number_of(a)
    return float(number) if is_double else number


# --------------------------------------------------------------------- Operators / 中文：运算符
def _trunc_div(a, b):
    quotient = abs(a) // abs(b)
    return quotient if (a >= 0) == (b >= 0) else -quotient


def ovpp_intdiv(a, b):
    """VBScript `\\` (integer division, truncated toward zero).

    中文：VBS 的 `\\`（整除，向零截断）。
    """
    left = _to_number(a)
    right = _to_number(b)
    if right == 0:
        raise ZeroDivisionError("Division by zero")
    return _trunc_div(int(left), int(right))


def ovpp_mod(a, b):
    """VBScript `Mod`: the remainder takes the sign of the dividend (truncated toward
    zero).

    中文：VBS 的 Mod —— 余数符号与被除数一致（向零截断）。
    """
    left = _to_number(a)
    right = _to_number(b)
    if right == 0:
        raise ZeroDivisionError("Division by zero")
    return left - right * _trunc_div(int(left), int(right))


# --------------------------------------------------------------------- Native arithmetic
# The `Option Native` variants. Same operators, but a text operand is an error
# 13 instead of being converted. 中文：`Option Native` 版本 —— 运算符相同，
# 但文本操作数报错误 13，而不是被转换。
def _native_operand(value):
    """Return a numeric operand, or raise error 13 for text.

    中文：取数值操作数；文本直接报错误 13。
    """
    if isinstance(value, (int, float)):
        return value
    if _is_empty(value):
        return 0
    if _is_null(value):
        return NULL
    raise TypeError(
        "Type mismatch: " +
        ("String" if isinstance(value, str) else type(value).__name__) +
        " cannot be used as a number (Option Native is on)")


def _native_pair(a, b):
    left = _native_operand(a)
    right = _native_operand(b)
    if left is NULL or right is NULL:
        return NULL, NULL
    return left, right


def ovpp_add_native(a, b):
    """`+` under Option Native: text is an error, never a concatenation.

    中文：`Option Native` 下的 `+` —— 文本报错，绝不拼接。
    """
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float):
        return a + b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _native_pair(a, b)
    return left + right


def ovpp_sub_native(a, b):
    """`-` under Option Native. 中文：`Option Native` 下的减法。"""
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float):
        return a - b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _native_pair(a, b)
    return left - right


def ovpp_mul_native(a, b):
    """`*` under Option Native. 中文：`Option Native` 下的乘法。"""
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float):
        return a * b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _native_pair(a, b)
    return left * right


def ovpp_div_native(a, b):
    """`/` under Option Native: still always a Double.

    中文：`Option Native` 下的 `/` —— 结果仍恒为 Double。
    """
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float) and b:
        return a / b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _native_pair(a, b)
    if right == 0:
        if left == 0:
            raise OverflowError("Overflow")
        raise ZeroDivisionError("Division by zero")
    return float(left) / float(right)


def ovpp_pow_native(a, b):
    """`^` under Option Native: still always a Double.

    中文：`Option Native` 下的 `^` —— 结果仍恒为 Double。
    """
    ka = type(a)
    kb = type(b)
    if (ka is int or ka is float) and (kb is int or kb is float) and b >= 0:
        return float(a) ** b
    if _is_null(a) or _is_null(b):
        return NULL
    left, right = _native_pair(a, b)
    if left == 0 and right < 0:
        raise ValueError("Invalid procedure call or argument")
    try:
        result = float(left) ** float(right)
    except OverflowError:
        raise OverflowError("Overflow")
    if isinstance(result, complex):
        raise ValueError("Invalid procedure call or argument")
    return result


def ovpp_neg_native(a):
    """Unary minus under Option Native. 中文：`Option Native` 下的一元负号。"""
    kind = type(a)
    if kind is int or kind is float:
        return -a
    if _is_null(a):
        return NULL
    number, is_double = _number_of(_native_operand(a))
    return -float(number) if is_double else -number


def ovpp_pos_native(a):
    """Unary plus under Option Native. 中文：`Option Native` 下的一元正号。"""
    kind = type(a)
    if kind is int or kind is float:
        return a
    if _is_null(a):
        return NULL
    number, is_double = _number_of(_native_operand(a))
    return float(number) if is_double else number


def ovpp_intdiv_native(a, b):
    """`\\` under Option Native. 中文：`Option Native` 下的整除。"""
    left, right = _native_pair(a, b)
    if left is NULL:
        return NULL
    left = int(left)
    right = int(right)
    if right == 0:
        raise ZeroDivisionError("Division by zero")
    return _trunc_div(left, right)


def ovpp_mod_native(a, b):
    """`Mod` under Option Native. 中文：`Option Native` 下的取模。"""
    left, right = _native_pair(a, b)
    if left is NULL:
        return NULL
    left = int(left)
    right = int(right)
    if right == 0:
        raise ZeroDivisionError("Division by zero")
    return left - right * _trunc_div(left, right)


# --------------------------------------------------------------------- Strings / 中文：字符串
def Len(value):
    return len(_to_str(value))


def Left(text, length):
    return _to_str(text)[:int(length)]


def Right(text, length):
    n = int(length)
    text = _to_str(text)
    return text[-n:] if n else ""


def Mid(text, start, length=None):
    text = _to_str(text)
    begin = int(start) - 1
    if begin < 0:
        begin = 0
    if length is None:
        return text[begin:]
    n = int(length)
    if n < 0:
        raise ValueError("Mid: length cannot be negative")
    return text[begin:begin + n]


def UCase(text):
    return _to_str(text).upper()


def LCase(text):
    return _to_str(text).lower()


def LTrim(text):
    return _to_str(text).lstrip()


def RTrim(text):
    return _to_str(text).rstrip()


def Trim(text):
    return _to_str(text).strip()


def StrReverse(text):
    return _to_str(text)[::-1]


def Space(count):
    return " " * int(count)


def String(count, character):
    n = int(count)
    if isinstance(character, (int, float)) and not isinstance(character, bool):
        ch = chr(int(character))
    else:
        text = _to_str(character)
        ch = text[0] if text else ""
    return ch * n


def Asc(text):
    text = _to_str(text)
    return ord(text[0]) if text else 0


def Chr(code):
    return chr(int(code))


def InStr(*args):
    """`InStr([start,] string1, string2)` returns a 1-based position, or 0 when the
    text is not found.

    中文：`InStr([start,] string1, string2)` → 1-based 位置，找不到返回 0。
    """
    if len(args) == 2:
        start, haystack, needle = 1, args[0], args[1]
    elif len(args) == 3:
        start, haystack, needle = args[0], args[1], args[2]
    else:
        raise TypeError("InStr expects 2 or 3 arguments")

    haystack = _to_str(haystack)
    needle = _to_str(needle)
    pos = haystack.find(needle, max(0, int(start) - 1))
    return 0 if pos < 0 else pos + 1


def InStrRev(haystack, needle, start=None):
    haystack = _to_str(haystack)
    needle = _to_str(needle)
    end = len(haystack) if start is None else int(start)
    pos = haystack.rfind(needle, 0, end)
    return 0 if pos < 0 else pos + 1


def Replace(text, find, replacement, start=1, count=-1):
    text = _to_str(text)
    find = _to_str(find)
    replacement = _to_str(replacement)
    begin = max(0, int(start) - 1)
    if count is None or int(count) < 0:
        return text[:begin] + text[begin:].replace(find, replacement)
    return text[:begin] + text[begin:].replace(find, replacement, int(count))


def Split(text, delimiter=None, limit=-1):
    text = _to_str(text)
    if delimiter is None or delimiter is EMPTY:
        delimiter = " "
    delimiter = _to_str(delimiter)
    if int(limit) > 0:
        return text.split(delimiter, int(limit) - 1)
    return text.split(delimiter)


def Join(items, delimiter=None):
    if delimiter is None or delimiter is EMPTY:
        delimiter = " "
    return _to_str(delimiter).join(_to_str(item) for item in _to_list(items))


# --------------------------------------------------------------------- Conversions / 中文：转换
def CStr(value):
    return _to_str(value)


def CInt(value):
    """VBScript `CInt` rounds to the nearest even number (banker's rounding).

    中文：VBS 的 `CInt` 会做四舍五入（银行家舍入）。
    """
    return int(round(_to_number(value)))


def CLng(value):
    return int(round(_to_number(value)))


def CDbl(value):
    return float(_to_number(value))


def CBool(value):
    return to_bool(value)


def CByte(value):
    return int(_to_number(value)) & 0xFF


# --------------------------------------------------------------------- Math / 中文：数学
def Abs(value):
    return abs(_to_number(value))


def Int(value):
    """Round toward negative infinity.

    中文：向负无穷取整。
    """
    return math.floor(_to_number(value))


def Fix(value):
    """Round toward zero.

    中文：向零取整。
    """
    number = _to_number(value)
    return int(number)


def Sgn(value):
    number = _to_number(value)
    if number > 0:
        return 1
    if number < 0:
        return -1
    return 0


def Sqr(value):
    return math.sqrt(_to_number(value))


def Round(value, digits=0):
    """VBScript `Round` uses banker's rounding, matching Python's `round`.

    中文：VBS 的 `Round` 使用银行家舍入，和 Python 的 `round` 一致。
    """
    return round(_to_number(value), int(digits))


def Rnd(*args):
    if args and _to_number(args[0]) < 0:
        random.seed(0)
    return random.random()


def Randomize(*args):
    random.seed() if not args else random.seed(_to_number(args[0]))
    return EMPTY


# --------------------------------------------------------------------- Arrays / 中文：数组
def Array(*items):
    return list(items)


def UBound(items, dimension=1):
    return ubound(_to_list(items), dimension)


def LBound(items, dimension=1):
    _to_list(items)
    return lbound(items, dimension)


# --------------------------------------------------------------------- Type checks / 中文：判断
def IsNumeric(value):
    if isinstance(value, bool):
        return True
    if isinstance(value, (int, float)):
        return True
    if value is EMPTY:
        return True
    try:
        float(_to_str(value).strip())
    except (TypeError, ValueError):
        return False
    return True


def IsEmpty(value):
    return value is EMPTY or isinstance(value, EmptyType)


def IsNull(value):
    return value is NULL or isinstance(value, NullType)


def IsObject(value):
    scalar = (int, float, str, bool, list, tuple, EmptyType, NullType)
    return value is not None and not isinstance(value, scalar)


def IsArray(value):
    return isinstance(value, list)


def TypeName(value):
    try:
        from OpenVBSPP.ovpprt.nothing import NOTHING
        if value is NOTHING:
            return "Nothing"
    except Exception:
        pass
    if _is_null(value):
        return "Null"
    vbs_name = getattr(value, "_vbs_name", None)
    if vbs_name:
        return vbs_name
    if value is EMPTY or isinstance(value, EmptyType):
        return "Empty"
    if isinstance(value, list):
        return "Variant()"
    if isinstance(value, VBool) or isinstance(value, bool):
        return "Boolean"
    if isinstance(value, int):
        return "Integer"
    if isinstance(value, float):
        return "Double"
    if isinstance(value, str):
        return "String"
    return type(value).__name__


# --------------------------------------------------------------------- Miscellaneous / 中文：其它
def Hex(value):
    number = int(_to_number(value))
    return format(number & 0xFFFFFFFF, "X") if number < 0 else format(number, "X")


def Oct(value):
    number = int(_to_number(value))
    return format(number & 0xFFFFFFFF, "o") if number < 0 else format(number, "o")


def StrComp(string1, string2, compare=None):
    a = _to_str(string1)
    b = _to_str(string2)
    if compare is not None and int(compare) == 1:      # vbTextCompare
        a, b = a.lower(), b.lower()
    if a == b:
        return 0
    return -1 if a < b else 1


def Filter(items, match, include=True, compare=None):
    pattern = _to_str(match)
    flags = re.IGNORECASE if (compare is not None and int(compare) == 1) else 0
    keep_matching = True if include is None else bool(include)
    out = []
    for item in _to_list(items):
        hit = re.search(pattern, _to_str(item), flags) is not None
        if hit == keep_matching:
            out.append(item)
    return out


# --------------------------------------------------------------------- Date and time / 中文：日期时间
# Accepted shapes for a date written as text; the first match wins.
# 中文：日期文本可接受的形式，按顺序取第一个匹配成功的。
_DATE_FORMATS = (
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%Y/%m/%d %H:%M:%S", "%Y/%m/%d",
    "%m/%d/%Y %H:%M:%S", "%m/%d/%Y",
    "%d %B %Y", "%B %d, %Y",
)
_DATE_AMPM_FORMATS = (
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p",
)
_TIME_FORMATS = ("%H:%M:%S", "%H:%M")
_AMPM_FORMATS = ("%I:%M:%S %p", "%I:%M %p")

# VBScript's date zero: same anchor Python's serial-date helpers use.
# 中文：VBS 的日期零点，和序列日期换算用的是同一个锚点。
_DATE_ORIGIN = (1899, 12, 30)


def _time_only(moment):
    """Turn a parsed clock time into the VBS Date value that holds it.

    中文：把解析出的时间部分转成承载它的 VBS Date 值。
    """
    return _datetime.datetime(_DATE_ORIGIN[0], _DATE_ORIGIN[1], _DATE_ORIGIN[2],
                              moment.hour, moment.minute, moment.second)


def _parse_date_text(text):
    """Parse a date/time string the way VBScript's `CDate` accepts it.

    中文：按 VBScript `CDate` 的接受范围解析日期/时间字符串。
    """
    for fmt in _DATE_FORMATS:
        try:
            return _datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
    for fmt in _DATE_AMPM_FORMATS:
        try:
            return _datetime.datetime.strptime(text.upper(), fmt)
        except ValueError:
            continue
    for fmt in _TIME_FORMATS + _AMPM_FORMATS:
        try:
            return _time_only(_datetime.datetime.strptime(text.upper(), fmt))
        except ValueError:
            continue
    raise ValueError(f"Cannot convert '{text}' to a date")


def date_literal(text):
    """Implementation of the `#...#` literal: strip the delimiters, then parse.

    中文：`#...#` 字面量的实现 —— 去掉两侧的 #，再按日期解析。
    """
    return _parse_date_text(str(text).strip().strip("#").strip())


def _to_datetime(value):
    if value is None or value is EMPTY or isinstance(value, EmptyType):
        return _datetime.datetime.now()
    if isinstance(value, _datetime.datetime):
        return value
    if isinstance(value, _datetime.date):
        return _datetime.datetime(value.year, value.month, value.day)
    if isinstance(value, _datetime.time):
        # VBS has a single Date type, so `Hour(Time())` must keep working.
        # 中文：VBS 只有一种 Date 类型，所以 `Hour(Time())` 必须能正常工作。
        return _time_only(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # A serial number: day 1 is 1899-12-31 in VBS, day 0 in our anchor.
        # 中文：序列日期 —— VBS 里第 1 天是 1899-12-31，这里以 1899-12-30 为原点。
        origin = _datetime.datetime(_DATE_ORIGIN[0], _DATE_ORIGIN[1], _DATE_ORIGIN[2])
        return origin + _datetime.timedelta(days=float(value))
    return _parse_date_text(_to_str(value).strip())


def Now():
    return _datetime.datetime.now()


def Date():
    return _datetime.datetime.now().date()


def Time():
    return _datetime.datetime.now().time()


def Timer():
    now = _datetime.datetime.now()
    midnight = now.replace(hour=0, minute=0, second=0, microsecond=0)
    return (now - midnight).total_seconds()


def Year(value=None):
    return _to_datetime(value).year


def Month(value=None):
    return _to_datetime(value).month


def Day(value=None):
    return _to_datetime(value).day


def Hour(value=None):
    return _to_datetime(value).hour


def Minute(value=None):
    return _to_datetime(value).minute


def Second(value=None):
    return _to_datetime(value).second


def Weekday(value=None, firstdayofweek=1):
    # VBScript: Sunday = 1, Monday = 2, ...
    # 中文：VBS —— 星期日 = 1，星期一 = 2 ……
    return _to_datetime(value).isoweekday() % 7 + 1


def IsDate(value):
    try:
        _to_datetime(value)
    except (TypeError, ValueError):
        return False
    return True


_INTERVALS = {
    "yyyy": "year", "y": "day", "m": "month", "d": "day",
    "w": "day", "ww": "week", "q": "quarter",
    "h": "hour", "n": "minute", "s": "second",
}


def _days_in_month(year, month):
    import calendar
    return calendar.monthrange(year, month)[1]


def DateAdd(interval, number, value):
    kind = _INTERVALS.get(_to_str(interval).strip().lower())
    if kind is None:
        raise ValueError(f"Unknown date interval '{interval}'")

    base = _to_datetime(value)
    count = int(_to_number(number))

    if kind == "year":
        year = base.year + count
        return base.replace(year=year, day=min(base.day, _days_in_month(year, base.month)))
    if kind == "month":
        total = (base.year * 12 + base.month - 1) + count
        year, month = divmod(total, 12)
        month += 1
        return base.replace(year=year, month=month,
                            day=min(base.day, _days_in_month(year, month)))
    if kind == "quarter":
        return DateAdd("m", count * 3, base)

    delta = _datetime.timedelta(**{kind + "s": count})
    return base + delta


def DateDiff(interval, date1, date2, firstdayofweek=None, firstweekofyear=None):
    kind = _INTERVALS.get(_to_str(interval).strip().lower())
    if kind is None:
        raise ValueError(f"Unknown date interval '{interval}'")

    start = _to_datetime(date1)
    end = _to_datetime(date2)
    delta = end - start

    if kind == "year":
        return end.year - start.year
    if kind == "month":
        return (end.year - start.year) * 12 + (end.month - start.month)
    if kind == "quarter":
        return ((end.year - start.year) * 12 + (end.month - start.month)) // 3
    if kind == "day":
        return delta.days
    if kind == "week":
        return delta.days // 7
    if kind == "hour":
        return int(delta.total_seconds() // 3600)
    if kind == "minute":
        return int(delta.total_seconds() // 60)
    return int(delta.total_seconds())


# --------------------------------------------------------------------- Formatting / 中文：格式化
_MONTHS = ("January", "February", "March", "April", "May", "June", "July",
           "August", "September", "October", "November", "December")
_MONTHS_ABBR = tuple(name[:3] for name in _MONTHS)
_WEEKDAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday",
             "Friday", "Saturday")
_WEEKDAYS_ABBR = tuple(name[:3] for name in _WEEKDAYS)

_DATE_TOKENS = ("am/pm", "yyyy", "yy", "mmmm", "mmm", "mm", "dddd", "ddd",
                "dd", "hh", "nn", "ss", "m", "d", "h", "n", "s")
_NUMBER_CHARS = set("#0")


def _vbs_weekday(moment):
    """VBScript `Weekday`: Sunday = 1.

    中文：VBS 的 `Weekday` —— 周日 = 1。
    """
    return (moment.weekday() + 1) % 7 + 1


def _looks_like_date(fmt):
    lowered = fmt.lower()
    return any(token in lowered for token in
               ("yyyy", "yy", "dd", "hh", "nn", "ss", "am/pm",
                "mmmm", "ddd"))


def _format_date_pattern(moment, fmt):
    out = []
    index = 0
    lowered = fmt.lower()
    while index < len(fmt):
        if lowered.startswith("am/pm", index):
            out.append("AM" if moment.hour < 12 else "PM")
            index += 5
            continue
        replaced = False
        for token, value in (
                ("yyyy", f"{moment.year:04d}"), ("yy", f"{moment.year % 100:02d}"),
                ("mmmm", _MONTHS[moment.month - 1]), ("mmm", _MONTHS_ABBR[moment.month - 1]),
                ("mm", f"{moment.month:02d}"),
                ("dddd", _WEEKDAYS[_vbs_weekday(moment) - 1]),
                ("ddd", _WEEKDAYS_ABBR[_vbs_weekday(moment) - 1]),
                ("dd", f"{moment.day:02d}"),
                ("hh", f"{moment.hour:02d}"), ("nn", f"{moment.minute:02d}"),
                ("ss", f"{moment.second:02d}"),
                ("m", str(moment.month)), ("d", str(moment.day)),
                ("h", str(moment.hour)), ("n", str(moment.minute)),
                ("s", str(moment.second))):
            if lowered.startswith(token, index):
                out.append(value)
                index += len(token)
                replaced = True
                break
        if not replaced:
            out.append(fmt[index])
            index += 1
    return "".join(out)


def _format_number_pattern(number, fmt):
    core = fmt.strip()
    prefix = ""
    if core.startswith("$") or core.startswith("€") or core.startswith("¥"):
        prefix, core = core[0], core[1:]
    percent = core.endswith("%")
    if percent:
        core = core[:-1]
        number = number * 100
    if "." in core:
        left, right = core.split(".", 1)
        decimals = sum(1 for ch in right if ch in _NUMBER_CHARS)
    else:
        left, decimals = core, 0
    grouping = "," in left
    integer_pattern = left.replace(",", "")
    # An all-zero format such as "0000" needs leading zeros.
    # 中文：整段都是 0 的格式（如 "0000"）要补前导零。
    zero_pad = len(integer_pattern) if integer_pattern and "#" not in integer_pattern else 0
    if zero_pad and not grouping:
        width = zero_pad + (decimals + 1 if decimals else 0)
        text = f"{number:0{width}.{decimals}f}"
    elif grouping:
        text = f"{number:,.{decimals}f}"
    else:
        text = f"{number:.{decimals}f}"
    return prefix + text + ("%" if percent else "")


def Format(value, fmt=""):
    """VBScript `Format`: accepts date and number format strings.

    中文：VBS 的 `Format` —— 支持日期格式串与数字格式串。
    """
    if fmt is None or fmt == "":
        return to_text(value)
    text = str(fmt)
    if _looks_like_date(text):
        return _format_date_pattern(_to_datetime(value), text)
    if any(ch in _NUMBER_CHARS for ch in text):
        try:
            return _format_number_pattern(_to_number(value), text)
        except (TypeError, ValueError):
            return to_text(value)
    return to_text(value)


# --------------------------------------------------------------------- Date / 中文：日期
def _fmt_digits(digits, default=2):
    try:
        value = int(digits)
    except (TypeError, ValueError):
        return default
    return default if value < 0 else value


def FormatNumber(value, digits=-1, leadingdigit=-1, parens=-1, group=-1):
    number = _to_number(value)
    places = _fmt_digits(digits)
    text = (f"{number:,.{places}f}" if group != 0 else f"{number:.{places}f}")
    if parens == 1 and number < 0:
        text = "(" + text.lstrip("-") + ")"
    return text


def FormatPercent(value, digits=-1, leadingdigit=-1, parens=-1, group=-1):
    return FormatNumber(_to_number(value) * 100, digits,
                        leadingdigit, parens, group) + "%"


def FormatCurrency(value, digits=-1, leadingdigit=-1, parens=-1, group=-1):
    number = _to_number(value)
    text = FormatNumber(number, digits, leadingdigit, parens, group)
    if number < 0:
        return "-$" + text.lstrip("-")
    return "$" + text


def FormatDateTime(value, namedformat=0):
    moment = _to_datetime(value)
    fmt = int(namedformat)
    if fmt == 1:                       # vbLongDate
        return moment.strftime("%A, %B %d, %Y")
    if fmt == 2:                       # vbShortDate
        return moment.strftime("%Y-%m-%d")
    if fmt == 3:                       # vbLongTime
        return moment.strftime("%H:%M:%S")
    if fmt == 4:                       # vbShortTime
        return moment.strftime("%H:%M")
    return moment.strftime("%Y-%m-%d %H:%M:%S")


def DateSerial(year, month, day):
    """A date serial value; month and day may overflow (month 13 is January of the
    next year).

    中文：日期序列值 —— 月份/日可以越界（13 月 = 次年 1 月）。
    """
    year, month, day = int(year), int(month), int(day)
    year += (month - 1) // 12
    month = (month - 1) % 12 + 1
    return _datetime.datetime(year, month, 1) + _datetime.timedelta(days=day - 1)


def TimeSerial(hour, minute, second):
    """A time serial value; minutes and seconds may overflow.

    中文：时间序列值 —— 分/秒可以越界。
    """
    total = int(hour) * 3600 + int(minute) * 60 + int(second)
    total %= 86400
    return _datetime.datetime(1899, 12, 30) + _datetime.timedelta(seconds=total)


def DateValue(value):
    moment = _to_datetime(value)
    return _datetime.datetime(moment.year, moment.month, moment.day)


def TimeValue(value):
    moment = _to_datetime(value)
    return _datetime.datetime(1899, 12, 30, moment.hour, moment.minute,
                             moment.second)


def DatePart(interval, value):
    """DatePart("yyyy"/"m"/"d"/"h"/"n"/"s"/"q"/"y"/"w"/"ww", date)"""
    moment = _to_datetime(value)
    key = str(interval).strip().lower()
    if key in ("yyyy", "year"):
        return moment.year
    if key in ("q", "quarter"):
        return (moment.month - 1) // 3 + 1
    if key in ("m", "month"):
        return moment.month
    if key in ("y", "dayofyear"):
        return moment.timetuple().tm_yday
    if key in ("d", "day"):
        return moment.day
    if key in ("w", "weekday"):
        return _vbs_weekday(moment)
    if key in ("ww", "weekofyear"):
        return moment.isocalendar()[1]
    if key in ("h", "hour"):
        return moment.hour
    if key in ("n", "minute"):
        return moment.minute
    if key in ("s", "second"):
        return moment.second
    return 0


def MonthName(month, abbreviate=False):
    index = int(month) - 1
    if index < 0 or index > 11:
        return ""
    return _MONTHS_ABBR[index] if abbreviate else _MONTHS[index]


def WeekdayName(weekday, abbreviate=False, firstdayofweek=1):
    index = (int(weekday) - int(firstdayofweek)) % 7
    return _WEEKDAYS_ABBR[index] if abbreviate else _WEEKDAYS[index]


# --------------------------------------------------------------------- Value semantics / 中文：值语义

class Delegate:
    """The procedure value produced by `AddressOf`.

    A procedure body reads a ByRef parameter as `param.v`, so the arguments are
    wrapped in `Ref` at the ByRef positions recorded at compile time. Assignments are
    not propagated back, which is the trade-off made for procedure values.

    中文：`AddressOf` 产生的「过程值」。过程体的 ByRef 参数读的是 `参数.v`，所以这里
    Boxes the arguments recorded as ByRef at compile time; write-backs do not propagate, a deliberate trade-off for procedure values.
    中文：按编译期记录的 ByRef 位置把实参包成 Ref；写回不传播（过程值语义下的取舍）。
    """

    def __init__(self, function, name="", byref_flags=None):
        self.function = function
        self.name = name or getattr(function, "__name__", "delegate")
        self.byref_flags = list(byref_flags or [])

    def __call__(self, *args):
        if not self.byref_flags:
            return self.function(*args)
        from OpenVBSPP.ovpprt.ref import Ref
        wrapped = []
        for index, value in enumerate(args):
            flag = self.byref_flags[index] if index < len(self.byref_flags) else False
            wrapped.append(Ref(value) if flag else value)
        return self.function(*wrapped)

    def __repr__(self):
        return f"Delegate({self.name})"


def delegate(function, name="", byref_flags=None):
    return Delegate(function, name, byref_flags)


def call_value(value, args):
    """`f(a)`: either array indexing or calling a procedure value.

    中文：`f(a)` —— 既是「数组下标」也是「调用过程值」。
    """
    if isinstance(value, Delegate) or (callable(value)
                                       and not isinstance(value, (list, tuple, str))):
        return value(*args)
    for index in args[:-1]:
        value = value[index]
    return value[args[-1]] if args else value


def coalesce(left, right):
    """`a ?? b`: uses the right side when the left side is Empty, Null or Nothing.

    中文：`a ?? b` —— 左边是 Empty/Null/Nothing 时取右边。
    """
    from OpenVBSPP.ovpprt.empty import EMPTY, EmptyType
    from OpenVBSPP.ovpprt.null import NULL, NullType
    from OpenVBSPP.ovpprt.nothing import NOTHING

    if left is NOTHING or left is NULL or isinstance(left, (EmptyType, NullType)):
        return right
    return left


# --------------------------------------------------------------------- Output / 中文：输出
def _safe_write(text):
    """Write one line to stdout, degrading instead of raising when the console cannot
    encode a character.

    Old Windows consoles default to GBK/CP1252, where printing Chinese text or emoji
    raises UnicodeEncodeError and shows up in a script as a runtime error. A failed
    write should never fail the program.

    中文：写一行到 stdout；遇到控制台编码不支持的字符时降级而不是抛异常。Windows 的
    Legacy consoles default to GBK or CP1252, where printing CJK text or emoji raises UnicodeEncodeError and then
    中文：旧控制台默认是 GBK/CP1252，直接 print 中文/emoji 会抛 UnicodeEncodeError，进而
    surfaces as a run-time error in the user's script. A failed write must not fail the program.
    中文：在用户脚本里表现为「运行时错误」——输出失败不应该让程序失败。
    """
    stream = sys.stdout
    line = text + "\n"
    try:
        stream.write(line)
    except UnicodeEncodeError:
        buffer = getattr(stream, "buffer", None)
        if buffer is None:
            raise
        encoding = getattr(stream, "encoding", None) or "utf-8"
        buffer.write(line.encode(encoding, "backslashreplace"))
        buffer.flush()
        return
    try:
        stream.flush()
    except Exception:
        pass


def Echo(*values):
    """The implementation of the Echo statement: VBScript-style output, so 3.0 prints
    as 3 and Empty prints as nothing.

    中文：Echo 语句的实现 —— 按 VBS 的习惯输出（3.0 打印成 3，Empty 打印成空）。
    """
    _safe_write(" ".join(to_text(value) for value in values))


def _echo_text(value):
    return to_text(value)


# --------------------------------------------------------------------- Interaction / 中文：交互
def MsgBox(message="", buttons=0, title=""):
    """Show a modal message box and return the button the user clicked.

    The result is a VBScript code: `1` vbOK, `2` vbCancel, `3` vbAbort,
    `4` vbRetry, `5` vbIgnore, `6` vbYes, `7` vbNo. Dialogs come from the
    bundled `ui-support` over package; when they are not enabled (library use,
    tests, CI, a non-interactive run, or no `tkinter`) the message is printed to
    stdout and `1` is returned, which is the historical behaviour.

    中文：弹出模态消息框，返回用户点击的按钮码：`1` vbOK、`2` vbCancel、
    `3` vbAbort、`4` vbRetry、`5` vbIgnore、`6` vbYes、`7` vbNo。
    对话框由内置 `ui-support` 包提供；未启用时（库调用、测试、CI、非交互运行、
    或没有 `tkinter`）退回「打印文本并返回 1」的历史行为。
    """
    from OpenVBSPP.ovpprt import ui

    answer = ui.msgbox(_to_str(message), buttons, _to_str(title))
    if answer is None:
        print(_to_str(message))
        return 1
    return answer


def InputBox(prompt="", title="", default=""):
    """Show a modal single-line prompt and return the typed text.

    Cancel returns `""`, exactly like VBScript. Without dialogs the prompt is
    printed and one line is read from stdin (the command-line stand-in).

    中文：弹出模态单行输入框，返回输入的文本。取消时返回 `""`，与 VBScript 一致。
    没有对话框时把提示打印出来，再从标准输入读一行（命令行环境下的替代实现）。
    """
    from OpenVBSPP.ovpprt import ui

    answer = ui.inputbox(_to_str(prompt), _to_str(title), _to_str(default))
    if answer is not None:
        return answer
    print(_to_str(prompt))
    return input()


# --------------------------------------------------------------------- Function registry / 中文：注册表
# The functions exposed to scripts are listed explicitly so internal helpers never
# leak into the language.
# 中文：显式列出暴露给 VBS 的内置函数，避免把模块内部的辅助函数混进去。

# --------------------------------------------------------------------- Math extras / 中文：数学补充
# VBScript exposes the trig / log family straight from the underlying library.
# 中文：VBScript 直接把三角/对数这一族暴露给脚本。
def Atn(value):
    return math.atan(_to_number(value))


def Cos(value):
    return math.cos(_to_number(value))


def Sin(value):
    return math.sin(_to_number(value))


def Tan(value):
    return math.tan(_to_number(value))


def Exp(value):
    return math.exp(_to_number(value))


def Log(value):
    """Natural logarithm; VBS raises on zero and negatives.

    中文：自然对数；VBS 对 0 与负数报错。
    """
    number = _to_number(value)
    if number <= 0:
        raise ValueError("Log: argument must be positive")
    return math.log(number)


# --------------------------------------------------------------------- Conversions / 中文：转换
def CCur(value):
    """Currency is a fixed-point number with four decimals.

    中文：Currency 是四位小数的定点数。
    """
    return round(_to_number(value), 4)


def CSng(value):
    """Single precision: keep about seven significant digits.

    中文：单精度 —— 保留约七位有效数字。
    """
    return float(f"{_to_number(value):.7g}")


def CDate(value):
    """Convert to the VBS Date value (a `datetime` when it runs here).

    中文：转换成 VBS 的 Date 值（在这里就是 `datetime`）。
    """
    return _to_datetime(value)


# --------------------------------------------------------------------- Type info / 中文：类型信息
# The vbVarType constants a script can compare against.
# 中文：脚本里可以拿来比对的 vbVarType 常量。
VT_EMPTY, VT_NULL, VT_INTEGER, VT_LONG, VT_SINGLE = 0, 1, 2, 3, 4
VT_DOUBLE, VT_CURRENCY, VT_DATE, VT_STRING, VT_OBJECT = 5, 6, 7, 8, 9
VT_BOOLEAN, VT_VARIANT = 11, 12
VT_ARRAY = 8192


def VarType(value):
    """Return the vbVarType number of a value.

    中文：返回值的 vbVarType 编号。
    """
    if value is EMPTY or isinstance(value, EmptyType):
        return VT_EMPTY
    if value is NULL or isinstance(value, NullType):
        return VT_NULL
    if isinstance(value, list):
        return VT_ARRAY + VT_VARIANT
    if isinstance(value, VBool):
        return VT_BOOLEAN
    if isinstance(value, bool):
        return VT_BOOLEAN
    if isinstance(value, int):
        return VT_LONG
    if isinstance(value, float):
        return VT_DOUBLE
    if isinstance(value, str):
        return VT_STRING
    if isinstance(value, (_datetime.datetime, _datetime.date, _datetime.time)):
        return VT_DATE
    return VT_OBJECT


def RGB(red, green, blue):
    """Pack a colour the way VBS does: red + green*256 + blue*65536.

    中文：按 VBS 的方式打包颜色：红 + 绿*256 + 蓝*65536。
    """
    return int(red) + int(green) * 256 + int(blue) * 65536


# --------------------------------------------------------------------- Engine / 中文：引擎信息
# OpenVBS++ answers like a VBScript 5.8 host so version checks in old scripts keep
# working; the build number is ours, not Microsoft's.
# 中文：OpenVBS++ 以 VBScript 5.8 宿主的口径回答，好让老脚本里的版本判断继续成立；
# 构建号是我们自己的，不是微软的。
_ENGINE = "VBScript"
_ENGINE_BUILD = 0
_ENGINE_MAJOR = 5
_ENGINE_MINOR = 8
_LCID = 1033


def ScriptEngine():
    return _ENGINE


def ScriptEngineBuildVersion():
    return _ENGINE_BUILD


def ScriptEngineMajorVersion():
    return _ENGINE_MAJOR


def ScriptEngineMinorVersion():
    return _ENGINE_MINOR


def SetLocale(lcid):
    """Set the locale id and return the previous one.

    中文：设置区域 ID，并返回之前的那个。
    """
    global _LCID
    previous = _LCID
    _LCID = int(lcid)
    return previous


def GetLocale():
    return _LCID


# --------------------------------------------------------------------- Reflection / 中文：反射
def IsMissing(value):
    """True when an optional argument was left out.

    OpenVBS++ fills an omitted optional argument with `Empty`, so this cannot tell
    "omitted" from "explicitly Empty" — same result for both.

    中文：判断可选参数是否被省略。OpenVBS++ 用 `Empty` 填充省略的实参，因此无法区分
    「省略」与「显式传 Empty」，两者结果相同。
    """
    return value is EMPTY or isinstance(value, EmptyType)


class _ProcedureRef(object):
    """Handle returned by GetRef: call it, or invoke it like an argument list.

    中文：GetRef 返回的句柄：可以直接调用，也可以像带参数的调用那样使用。
    """
    _vbs_name = "Object"

    def __init__(self, target, name):
        self._target = target
        self.name = name

    def __call__(self, *args):
        return self._target(*args)

    def __getitem__(self, key):
        # (21) where f holds a reference compiles to a subscript on the variable,
        # so route that form to the procedure as well.
        # 中文：变量持有过程引用时，(21) 会被编译成对该变量的下标取值，
        # 中文：因此这种写法也要转成一次过程调用。
        return self._target(key)

    def __repr__(self):
        return "<procedure ref {0}>".format(self.name)


def GetRef(name):
    """Return a reference to a named procedure.

    中文：返回指定名字的过程引用。
    """
    frame = sys._getframe(1)
    key = str(name)
    lower = key.lower()
    # Compiler-generated procedure functions are named _ovpp_fn_<name> and
    # _ovpp_sub_<name>; plain names are accepted too for NATIVE procs.
    # 中文：编译器生成的过程函数名为 _ovpp_fn_<名字> 与 _ovpp_sub_<名字>；
    # 中文：也接受直接的名字，便于引用原生过程。
    candidates = ("_ovpp_fn_" + lower, "_ovpp_sub_" + lower, key, lower)
    target = None
    for candidate in candidates:
        for table in (frame.f_locals, frame.f_globals):
            value = table.get(candidate)
            if callable(value):
                target = value
                break
        if target is not None:
            break
    if target is None:
        from OpenVBSPP.exceptions import OVPPProcedureNotDefinedError

        raise OVPPProcedureNotDefinedError(
            f"GetRef: no procedure named '{key}'", line_number=None)
    return _ProcedureRef(target, key)


def LoadPicture(path):
    """Not available: OpenVBS++ has no GUI picture library.

    中文：不支持 —— OpenVBS++ 没有图形图像库。
    """
    from OpenVBSPP.exceptions import OVPPObjectNotSupportedError

    raise OVPPObjectNotSupportedError(
        "LoadPicture is not available: OpenVBS++ has no GUI picture library")

# --------------------------------------------------------------------- Null rules
# VBScript splits the built-in functions into three camps when an argument is
# Null. The lists below were verified against the real VBScript host: each one
# names the behaviour its group gets.
# 中文：VBS 的内置函数遇到 Null 时分三类，下面的清单是按真实 VBS 宿主实测出来的。
_NULL_PROPAGATES = (
    "Len Left Right Mid LTrim RTrim Trim UCase LCase InStr StrComp Hex Oct "
    "Abs Int Fix Year Month Day Hour Minute Second Weekday DateAdd DateDiff DatePart"
).split()

_NULL_REJECTED = (
    "StrReverse Space String Asc Chr InStrRev Replace Split Join Filter "
    "Sgn Sqr Round Atn Cos Sin Tan Exp Log "
    "CStr CInt CLng CDbl CSng CCur CByte CBool CDate MonthName WeekdayName"
).split()

# These five report a plain type mismatch (13) instead of 94.
# 中文：这五个报普通的类型不匹配（13），而不是 94。
_NULL_MISMATCH = (
    "Format FormatNumber FormatPercent FormatCurrency FormatDateTime"
).split()


def _null_argument(args, kwargs):
    for value in args:
        if _is_null(value):
            return True
    for value in kwargs.values():
        if _is_null(value):
            return True
    return False


def _propagate_null(func):
    """Return Null as soon as any argument is Null.

    中文：任一参数为 Null 时直接返回 Null。
    """

    def wrapper(*args, **kwargs):
        if _null_argument(args, kwargs):
            return NULL
        return func(*args, **kwargs)

    wrapper.__name__ = func.__name__
    wrapper.__doc__ = func.__doc__
    return wrapper


def _reject_null(error):
    """Raise ``error`` (94 or 13) instead of accepting a Null argument.

    中文：参数为 Null 时抛 ``error``（94 或 13），而不是硬算下去。
    """

    def decorate(func):
        def wrapper(*args, **kwargs):
            if _null_argument(args, kwargs):
                raise error("Invalid use of Null" if error is OVPPInvalidUseOfNullError
                            else "Type mismatch: Null is not a valid argument")
            return func(*args, **kwargs)

        wrapper.__name__ = func.__name__
        wrapper.__doc__ = func.__doc__
        return wrapper

    return decorate


for _name in _NULL_PROPAGATES:
    globals()[_name] = _propagate_null(globals()[_name])
for _name in _NULL_REJECTED:
    globals()[_name] = _reject_null(OVPPInvalidUseOfNullError)(globals()[_name])
for _name in _NULL_MISMATCH:
    globals()[_name] = _reject_null(TypeError)(globals()[_name])
del _name

_BUILTIN_NAMES = (
    "Len Left Right Mid UCase LCase LTrim RTrim Trim StrReverse Space String "
    "Asc Chr InStr InStrRev Replace Split Join "
    "CStr CInt CLng CDbl CBool CByte "
    "Abs Int Fix Sgn Sqr Round Rnd Randomize "
    "Array UBound LBound "
    "IsNumeric IsEmpty IsNull IsObject IsArray IsDate TypeName "
    "Hex Oct StrComp Filter "
    "Now Date Time Timer Year Month Day Hour Minute Second Weekday DateAdd DateDiff "
    "Format FormatNumber FormatCurrency FormatPercent FormatDateTime "
    "DateSerial TimeSerial DateValue TimeValue DatePart MonthName WeekdayName "
    "MsgBox InputBox "
    "Atn Cos Sin Tan Exp Log CCur CSng CDate VarType RGB "
    "ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion "
    "SetLocale GetLocale IsMissing GetRef LoadPicture"
).split()

REGISTRY = {name.lower(): globals()[name] for name in _BUILTIN_NAMES}

