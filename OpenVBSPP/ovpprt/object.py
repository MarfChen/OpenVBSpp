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

Object runtime support for VBScript objects.

中文：对象运行时支持。

VBScript objects come from COM (IDispatch); this project does not emulate COM and offers two kinds of objects:

中文：VBS 的对象来自 COM（IDispatch）；本项目不模拟 COM，只提供两种对象：
  1. classes a script defines with `Class ... End Class`;
中文：脚本里用 `Class ... End Class` 定义的类；
  2. built-in classes, currently Scripting.Dictionary.
中文：内置类（目前是 Scripting.Dictionary）。

ProgID resolution for CreateObject happens here.
中文：`CreateObject` 的 ProgID 解析就发生在这里。
'''

from OpenVBSPP.exceptions import OVPPObjectNotSupportedError


_DEFAULT_MEMBER = "_default_member"


def _vbs_scalar(value):
    """Turn a plain Python `bool` into the VBScript Boolean (-1 / 0).

    Host objects are written in Python, where `True` is 1. VBScript's `True` is -1
    and its logical operators are bitwise, so leaking a Python bool makes
    `Not value` return -2 (truthy) instead of False. Every value that crosses the
    object boundary goes through here.

    中文：把普通 Python `bool` 变成 VBS 布尔（-1 / 0）。宿主对象是用 Python 写的，
    其中 `True` 是 1；VBS 的 `True` 是 -1 且逻辑运算符是按位的，漏出去的 Python
    bool 会让 `Not value` 得到 -2（仍为真）而不是 False。所有跨对象边界的值都经过这里。
    """
    if isinstance(value, bool):
        from OpenVBSPP.ovpprt.boolean import FALSE, TRUE
        return TRUE if value else FALSE
    return value


def create_object(progid, registry, *args):
    """The runtime implementation of `CreateObject("ProgID")`.

    The second VBScript argument names a remote server; remote activation is not
    supported locally, so passing it is an error.

    中文：`CreateObject("ProgID")` 的运行时实现。VBS 的第二个参数是 servername
    VBScript's second argument is servername, for remote activation, which a local
    implementation cannot do, so an argument raises an error.

    中文：VBS 的第二个参数是 servername（远程激活），本地无法实现，所以带参数时报错。
    """
    if args:
        raise OVPPObjectNotSupportedError(
            "Remote activation is not supported: CreateObject(class, servername)")
    key = str(progid).strip().lower()

    if key == "scripting.dictionary":
        from OpenVBSPP.ovpprt.collections import Dictionary
        return Dictionary()

    if key == "scripting.filesystemobject":
        from OpenVBSPP.ovpprt.filesystem import FileSystemObject
        return FileSystemObject()

    if key in ("vbscript.regexp", "regexp"):
        from OpenVBSPP.ovpprt.regexp import RegExp
        return RegExp()

    cls = registry.get(key)
    if cls is not None:
        return cls(*args)

    # Unknown ProgID: hand it to the COM bridge (the com-support over package on
    # Windows). Scripts import nothing; CreateObject itself reaches this point.
    # 中文：未知 ProgID 交给 COM 桥接（Windows 上的 com-support over 包）。
    # A script imports nothing: CreateObject reaches this point on its own.
    # 中文：脚本里不需要 import 任何东西，CreateObject 本身就会走到这里。
    from OpenVBSPP.ovpprt import com as _com
    active = _com.provider()
    if active is not None:
        return active.create(progid)

    raise OVPPObjectNotSupportedError(
        f"Cannot create object '{progid}': only script classes and "
        f"'Scripting.Dictionary' / 'Scripting.FileSystemObject' are supported "
        f"(COM objects are not available: {_com.describe()})")


def type_name(value):
    """The VBS class name of an object, or None when the value is not an object.

    中文：对象返回 VBS 类名；不是对象时返回 None。
    """
    return getattr(value, "_vbs_name", None)


def is_object(value):
    return getattr(value, "_vbs_name", None) is not None


def default_member_name(value):
    return getattr(value, _DEFAULT_MEMBER, None)


def default_value(value):
    """The value of `x = obj`: the default member when the object has one, else the
    object itself.

    中文：`x = obj` 的取值——对象有默认成员时取它，否则原样返回。
    """
    if not is_object(value):
        return value
    member = getattr(value, _DEFAULT_MEMBER, None)
    if not member:
        return value
    return member_get(value, member)


def get_object(pathname="", registry=None, classname=None):
    """`GetObject([pathname], [class])`: attach to an existing COM object.

    VBScript passes either a display-name moniker (`winmgmts:`), a file name that
    binds to its class, or an empty pathname plus a class name to attach to a
    running instance.  A plain ProgID keeps the historical fallback: it behaves
    like `CreateObject`.
    中文：`GetObject([pathname], [class])` —— 挂接已存在的 COM 对象。VBS 可以传
    显示名 moniker（`winmgmts:`）、绑定到类别的文件名，或空 pathname + 类名去挂接
    运行中的实例。普通 ProgID 保持历史行为：按 `CreateObject` 处理。
    """
    from OpenVBSPP.ovpprt.empty import EMPTY

    text = "" if pathname is EMPTY or pathname is None else str(pathname).strip()
    class_text = ("" if classname is None or classname is EMPTY
                  else str(classname).strip())

    if not text and not class_text:
        raise OVPPObjectNotSupportedError(
            "GetObject needs a pathname, a class name, or both")

    if not text:
        from OpenVBSPP.ovpprt import com
        return com.get("", class_text)

    if _looks_like_moniker(text) or _file_exists(text):
        from OpenVBSPP.ovpprt import com
        return com.get(text, class_text or None)

    return create_object(pathname, registry)


def _looks_like_moniker(text):
    """A display-name moniker or an absolute path, not a ProgID.

    中文：显示名 moniker 或绝对路径（而不是 ProgID）。
    """
    lowered = text.lower()
    if ":" in text:
        return True
    if "\\" in text or "/" in text:
        return True
    return lowered.startswith(("new:", "file:", "clsid:", "queue:"))


def _file_exists(text):
    """Whether the text names an existing file (a file moniker target).

    中文：该文本是否指向一个已存在的文件（file moniker 的目标）。
    """
    import os
    try:
        return os.path.exists(text)
    except (OSError, ValueError):
        return False


_MISSING = object()


def _lookup(obj, name):
    """Look a member up case insensitively, since VBScript identifiers are not
    case sensitive.

    Built-in objects expose CamelCase Python names (`CreateTextFile`), while scripts
    may write any casing, so a fallback lookup happens here.

    中文：大小写不敏感地取成员。内置对象的 Python 方法名是 CamelCase
    (CreateTextFile) while a script may use any casing, so fall back to a case-insensitive lookup.
    中文：（CreateTextFile），而 VBS 里写的是任意大小写，所以这里做一次回退查找。
    """
    target = getattr(obj, name, _MISSING)
    if target is not _MISSING:
        return target
    lowered = name.lower()
    for attribute in dir(obj):
        if attribute.lower() == lowered:
            return getattr(obj, attribute)
    return _MISSING


def arity_error(name, count, arities):
    """The error raised when overload dispatch fails; it lists the accepted arities.

    中文：重载分派失败时的错误（给出可用的参数个数）。
    """
    from OpenVBSPP.exceptions import OVPPArgumentError
    options = ", ".join(str(item) for item in arities)
    raise OVPPArgumentError(
        f"'{name}' has no overload taking {count} argument(s) (available: {options})")


#: Member names of Dictionary itself (lower case). In real VBScript `d.Add` always
#: resolves to the method and a data key has to be read as `d("Add")`, so these names
#: never need to scan the keys.
#: 中文：Dictionary 自身的成员名（小写）。真实 VBS 里 `d.Add` 永远解析成方法，
#: Data keys are reached as `d("Add")`, so these names need no key scan at all.
#: 中文：数据键要用 `d("Add")` 访问；因此这些名字不需要、也不应该去扫描键。
_DICT_MEMBER_NAMES = frozenset({
    "add", "comparemode", "count", "exists", "hashval", "item", "items",
    "key", "keys", "remove", "removeall",
})


def _dict_key(dictionary, name):
    """Find a user data key in a Dictionary, case insensitively.

    Method names always return None so normal member resolution handles them;
    otherwise every `d.Add(...)` would scan all keys and large dictionaries would
    degrade to O(n²).

    中文：大小写不敏感地在 Dictionary 里找用户数据键。方法名一律返回 None
    and falls through to normal member resolution, or every `d.Add(...)` would scan every key.
    中文：（交给常规成员解析），否则每次 `d.Add(...)` 都要线性扫描全部键。
    """
    lowered = name.lower()
    if lowered in _DICT_MEMBER_NAMES:
        return None
    for existing in dictionary.keys():
        if str(existing).lower() == lowered:
            return existing
    return None


def call_method(obj, name, refs):
    """The shared method-call protocol.

    VBScript class methods always take `Ref` arguments: ByRef writes straight to `.v`
    and ByVal unwraps at the method entry. For built-in objects (Dictionary and
    friends) that means taking `.v` and calling by value.

    中文：统一的方法调用协议。VBS 类的方法参数一律通过 `Ref` 传——ByRef 直接改
    `.v` and ByVal unpacks at the method entry, so a built-in object simply gets the unpacked value by value.
    中文：`.v`，ByVal 在方法入口解包，因此对内置对象就是把 `.v` 取出来按值调用。
    """
    from OpenVBSPP.ovpprt.collections import Dictionary
    from OpenVBSPP.ovpprt.array import get_item

    if isinstance(obj, Dictionary):
        key = _dict_key(obj, name)
        if key is not None:
            member = obj.item(key)
            values = [ref.v for ref in refs]
            if callable(member):
                return _vbs_scalar(member(*values))
            if not values:
                return _vbs_scalar(member)
            for index in values:
                member = get_item(member, index)
            return _vbs_scalar(member)

    target = _lookup(obj, name)
    if target is _MISSING:
        raise OVPPObjectNotSupportedError(
            f"Object of type '{type(obj).__name__}' has no member '{name}'")

    if callable(target):
        if getattr(obj, "_ovpp_vbs_class", False):
            return _vbs_scalar(target(*refs))
        return _vbs_scalar(target(*[ref.v for ref in refs]))

    # Not a method: treat it as array indexing.
    # 中文：不是方法——按数组下标处理。
    values = [ref.v for ref in refs]
    if not values:
        return _vbs_scalar(target)
    for index in values:
        target = get_item(target, index)
    return _vbs_scalar(target)


def _is_nothing(value):
    from OpenVBSPP.ovpprt.nothing import NOTHING
    return value is NOTHING or value is None


def safe_member(obj, name):
    """`a?.b`: when `a` is Nothing the result is Nothing instead of an error.

    中文：`a?.b` —— a 是 Nothing 时结果是 Nothing，不报错。
    """
    if _is_nothing(obj):
        from OpenVBSPP.ovpprt.nothing import NOTHING
        return NOTHING
    return member_get(obj, name)


def safe_call(obj, name, refs):
    """`a?.b(args)`: when `a` is Nothing the result is Nothing.

    中文：`a?.b(args)` —— a 是 Nothing 时直接返回 Nothing。
    """
    if _is_nothing(obj):
        from OpenVBSPP.ovpprt.nothing import NOTHING
        return NOTHING
    return call_method(obj, name, refs)


def write_only_error(message):
    """The error raised when a write-only property is read.

    中文：读取只写属性时的错误。
    """
    from OpenVBSPP.exceptions import OVPPObjectNotSupportedError
    raise OVPPObjectNotSupportedError(str(message))


def set_member(obj, name, value, is_set=False):
    """Member assignment: `Set obj.P = x` uses Property Set, everything else uses a
    setter or a plain attribute.

    中文：成员赋值 —— `Set obj.P = x` 走 Property Set，其余走 setter / 属性。
    """
    if is_set:
        writer = getattr(obj, "_ovpp_pset_" + name, None)
        if writer is not None:
            writer(value)
            return
    setattr(obj, name, value)


def member_get(obj, name):
    """`obj.name`: fields and properties are returned directly, while a zero-argument
    method is called (VBScript allows omitting the parentheses).

    中文：`obj.name` 取值 —— 字段/属性直接返回，无参方法则调用（VBS 允许省略括号）。
    """
    from OpenVBSPP.ovpprt.collections import Dictionary
    if isinstance(obj, Dictionary):
        key = _dict_key(obj, name)
        if key is not None:
            member = obj.item(key)
            return _vbs_scalar(member() if callable(member) else member)

    target = _lookup(obj, name)
    if target is _MISSING:
        raise OVPPObjectNotSupportedError(
            f"Object of type '{type(obj).__name__}' has no member '{name}'")
    if callable(target):
        return _vbs_scalar(target())
    return _vbs_scalar(target)


def member_call(obj, name, *args):
    """`obj.name(args...)`: either a method call or indexing into an array field.

    中文：`obj.name(args...)` —— 既可能是方法调用，也可能是数组字段的下标访问。
    """
    from OpenVBSPP.ovpprt.collections import Dictionary
    if isinstance(obj, Dictionary):
        key = _dict_key(obj, name)
        if key is not None:
            member = obj.item(key)
            values = list(args)
            if callable(member):
                return member(*values)
            if not values:
                return member
            for index in values[:-1]:
                member = member[index]
            return member[values[-1]]

    target = _lookup(obj, name)
    if target is _MISSING:
        raise OVPPObjectNotSupportedError(
            f"Object of type '{type(obj).__name__}' has no member '{name}'")

    if callable(target):
        return _vbs_scalar(target(*args))

    # Not a method: treat it as array indexing.
    # 中文：不是方法——按数组下标处理。
    if not args:
        return _vbs_scalar(target)
    for index in args[:-1]:
        target = target[index]
    return target[args[-1]]


def set_member_index(obj, name, args, value):
    """`obj.name(i, j) = value`: assign one element of an array field.

    中文：`obj.name(i, j) = value` —— 数组字段的元素赋值。
    """
    from OpenVBSPP.ovpprt.collections import Dictionary
    from OpenVBSPP.ovpprt.array import get_item, checked_target

    if isinstance(obj, Dictionary):
        key = _dict_key(obj, name)
        if key is not None:
            member = obj.item(key)
            for index in args[:-1]:
                member = get_item(member, index)
            checked_target(member)[args[-1]] = value
            return

    target = _lookup(obj, name)
    if target is _MISSING:
        raise OVPPObjectNotSupportedError(
            f"Object of type '{type(obj).__name__}' has no member '{name}'")

    for index in args[:-1]:
        target = get_item(target, index)
    checked_target(target)[args[-1]] = value
