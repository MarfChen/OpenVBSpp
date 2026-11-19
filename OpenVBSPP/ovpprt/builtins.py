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

VBS 常用内置函数。

编译器把 `Len(x)` 这样的调用翻译成 `_ovpp_builtin.Len(x)`，
所以这里的函数名保持 VBS 的写法（首字母大写）。
REGISTRY 用小写名字做键，从而实现 VBS 的大小写不敏感。
'''

import datetime as _datetime
import math
import random
import re
import sys

from OpenVBSPP.ovpprt.array import lbound, ubound
from OpenVBSPP.ovpprt.boolean import VBool, to_bool
from OpenVBSPP.ovpprt.empty import EMPTY, EmptyType
from OpenVBSPP.ovpprt.format import to_text
from OpenVBSPP.ovpprt.null import NULL, NullType


# --------------------------------------------------------------------- 工具
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


# --------------------------------------------------------------------- 运算符
def _trunc_div(a, b):
    quotient = abs(a) // abs(b)
    return quotient if (a >= 0) == (b >= 0) else -quotient


def ovpp_intdiv(a, b):
    """VBS 的 \\（整除，向零截断）。"""
    left = _to_number(a)
    right = _to_number(b)
    if right == 0:
        raise ZeroDivisionError("Division by zero")
    return _trunc_div(int(left), int(right))


def ovpp_mod(a, b):
    """VBS 的 Mod：余数符号与被除数一致（向零截断）。"""
    left = _to_number(a)
    right = _to_number(b)
    if right == 0:
        raise ZeroDivisionError("Division by zero")
    return left - right * _trunc_div(int(left), int(right))


# --------------------------------------------------------------------- 字符串
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
    """InStr([start,] string1, string2) → 1-based 位置，找不到返回 0。"""
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


# --------------------------------------------------------------------- 转换
def CStr(value):
    return _to_str(value)


def CInt(value):
    """VBS 的 CInt 会做四舍五入（银行家舍入）。"""
    return int(round(_to_number(value)))


def CLng(value):
    return int(round(_to_number(value)))


def CDbl(value):
    return float(_to_number(value))


def CBool(value):
    return to_bool(value)


def CByte(value):
    return int(_to_number(value)) & 0xFF


# --------------------------------------------------------------------- 数学
def Abs(value):
    return abs(_to_number(value))


def Int(value):
    """向负无穷取整。"""
    return math.floor(_to_number(value))


def Fix(value):
    """向零取整。"""
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
    """VBS 的 Round 使用银行家舍入，和 Python 的 round 一致。"""
    return round(_to_number(value), int(digits))


def Rnd(*args):
    if args and _to_number(args[0]) < 0:
        random.seed(0)
    return random.random()


def Randomize(*args):
    random.seed() if not args else random.seed(_to_number(args[0]))
    return EMPTY


# --------------------------------------------------------------------- 数组
def Array(*items):
    return list(items)


def UBound(items, dimension=1):
    return ubound(_to_list(items), dimension)


def LBound(items, dimension=1):
    _to_list(items)
    return lbound(items, dimension)


# --------------------------------------------------------------------- 判断
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


# --------------------------------------------------------------------- 其它
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


# --------------------------------------------------------------------- 日期时间
def _to_datetime(value):
    if value is None or value is EMPTY or isinstance(value, EmptyType):
        return _datetime.datetime.now()
    if isinstance(value, _datetime.datetime):
        return value
    if isinstance(value, _datetime.date):
        return _datetime.datetime(value.year, value.month, value.day)
    text = _to_str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d", "%H:%M:%S"):
        try:
            return _datetime.datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ValueError(f"Cannot convert '{text}' to a date")


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
    # VBS：星期日 = 1，星期一 = 2 ……
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


# --------------------------------------------------------------------- 格式化
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
    """VBS 的 Weekday：周日 = 1。"""
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
    # 整段都是 0 的格式（如 "0000"）要补前导零
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
    """VBS 的 Format：支持日期格式串与数字格式串。"""
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


# --------------------------------------------------------------------- 日期
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
    """日期序列值：月份/日可以越界（13 月 = 次年 1 月）。"""
    year, month, day = int(year), int(month), int(day)
    year += (month - 1) // 12
    month = (month - 1) % 12 + 1
    return _datetime.datetime(year, month, 1) + _datetime.timedelta(days=day - 1)


def TimeSerial(hour, minute, second):
    """时间序列值：分/秒可以越界。"""
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


# --------------------------------------------------------------------- 值语义

class Delegate:
    """AddressOf 产生的"过程值"。

    过程体的 ByRef 参数读的是 `参数.v`，所以这里按编译期记录的
    ByRef 位置把实参包成 Ref；写回不传播（过程值语义下的取舍）。
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
    """`f(a)`：既是"数组下标"也是"调用过程值"。"""
    if isinstance(value, Delegate) or (callable(value)
                                       and not isinstance(value, (list, tuple, str))):
        return value(*args)
    for index in args[:-1]:
        value = value[index]
    return value[args[-1]] if args else value


def coalesce(left, right):
    """`a ?? b`：左边是 Empty/Null/Nothing 时取右边。"""
    from OpenVBSPP.ovpprt.empty import EMPTY, EmptyType
    from OpenVBSPP.ovpprt.null import NULL, NullType
    from OpenVBSPP.ovpprt.nothing import NOTHING

    if left is NOTHING or left is NULL or isinstance(left, (EmptyType, NullType)):
        return right
    return left


# --------------------------------------------------------------------- 输出
def _safe_write(text):
    """写一行到 stdout；遇到控制台编码不支持的字符时降级而不是抛异常。

    Windows 的旧控制台默认是 GBK/CP1252，直接 print 中文/emoji 会抛
    UnicodeEncodeError，进而在用户脚本里表现为"运行时错误"——这不合理，
    输出失败不应该让程序失败。
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
    """Echo 语句的实现：按 VBS 的习惯输出（3.0 打印成 3，Empty 打印成空）。"""
    _safe_write(" ".join(to_text(value) for value in values))


def _echo_text(value):
    return to_text(value)


# --------------------------------------------------------------------- 交互
def MsgBox(message="", *args):
    """无 GUI：把消息打印到标准输出，返回 1（相当于点了"确定"）。"""
    print(_to_str(message))
    return 1


def InputBox(prompt="", *args):
    """从标准输入读一行（命令行环境下的替代实现）。"""
    print(_to_str(prompt))
    return input()


# --------------------------------------------------------------------- 注册表
# 显式列出暴露给 VBS 的内置函数，避免把模块内部的辅助函数混进去。
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
    "MsgBox InputBox"
).split()

REGISTRY = {name.lower(): globals()[name] for name in _BUILTIN_NAMES}

