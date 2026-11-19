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

对象运行时支持。

VBS 的对象来自 COM（IDispatch）；本项目不模拟 COM，只提供两种对象：
  1. 脚本里用 `Class ... End Class` 定义的类；
  2. 内置类（目前是 Scripting.Dictionary）。

`CreateObject` 的 ProgID 解析就发生在这里。
'''

from OpenVBSPP.exceptions import OVPPObjectNotSupportedError
from OpenVBSPP.ovpprt.collections import Dictionary
from OpenVBSPP.ovpprt.filesystem import FileSystemObject
from OpenVBSPP.ovpprt.regexp import RegExp


_DEFAULT_MEMBER = "_default_member"


def create_object(progid, registry, *args):
    """CreateObject("ProgID") 的运行时实现。

    VBS 的第二个参数是 servername（远程激活），本地无法实现，所以带参数时报错。
    """
    if args:
        raise OVPPObjectNotSupportedError(
            "Remote activation is not supported: CreateObject(class, servername)")
    key = str(progid).strip().lower()

    if key == "scripting.dictionary":
        return Dictionary()

    if key == "scripting.filesystemobject":
        return FileSystemObject()

    if key in ("vbscript.regexp", "regexp"):
        return RegExp()

    cls = registry.get(key)
    if cls is not None:
        return cls(*args)

    # 未知 ProgID：交给 COM 桥接（Windows 上的 com-support over 包）。
    # 脚本里不需要 import 任何东西，CreateObject 本身就会走到这里。
    from OpenVBSPP.ovpprt import com as _com
    active = _com.provider()
    if active is not None:
        return active.create(progid)

    raise OVPPObjectNotSupportedError(
        f"Cannot create object '{progid}': only script classes and "
        f"'Scripting.Dictionary' / 'Scripting.FileSystemObject' are supported "
        f"(COM objects are not available: {_com.describe()})")


def type_name(value):
    """对象返回 VBS 类名；不是对象时返回 None。"""
    return getattr(value, "_vbs_name", None)


def is_object(value):
    return getattr(value, "_vbs_name", None) is not None


def default_member_name(value):
    return getattr(value, _DEFAULT_MEMBER, None)


def default_value(value):
    """`x = obj` 的取值：对象有默认成员时取它，否则原样返回。"""
    if not is_object(value):
        return value
    member = getattr(value, _DEFAULT_MEMBER, None)
    if not member:
        return value
    return member_get(value, member)


def get_object(progid, registry, *args):
    """GetObject：本项目不做对象持久化，按 CreateObject 处理（见文档）。"""
    return create_object(progid, registry, *args)


_MISSING = object()


def _lookup(obj, name):
    """大小写不敏感地取成员（VBS 的标识符不区分大小写）。

    内置对象的 Python 方法名是 CamelCase（CreateTextFile），
    而 VBS 里写的是任意大小写，所以这里做一次回退查找。
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
    """重载分派失败时的错误（给出可用的参数个数）。"""
    from OpenVBSPP.exceptions import OVPPArgumentError
    options = ", ".join(str(item) for item in arities)
    raise OVPPArgumentError(
        f"'{name}' has no overload taking {count} argument(s) (available: {options})")


def _dict_key(dictionary, name):
    """大小写不敏感地在 Dictionary 里找键。"""
    lowered = name.lower()
    for existing in dictionary.keys():
        if str(existing).lower() == lowered:
            return existing
    return None


def call_method(obj, name, refs):
    """统一的方法调用协议。

    VBS 类的方法参数一律通过 `Ref` 传：ByRef 直接改 `.v`，ByVal 在方法入口解包，
    因此对内置对象（Dictionary 等）就是把 `.v` 取出来按值调用。
    """
    from OpenVBSPP.ovpprt.collections import Dictionary
    if isinstance(obj, Dictionary):
        key = _dict_key(obj, name)
        if key is not None:
            member = obj.item(key)
            values = [ref.v for ref in refs]
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
        if getattr(obj, "_ovpp_vbs_class", False):
            return target(*refs)
        return target(*[ref.v for ref in refs])

    # 不是方法：按数组下标处理
    values = [ref.v for ref in refs]
    if not values:
        return target
    for index in values[:-1]:
        target = target[index]
    return target[values[-1]]


def _is_nothing(value):
    from OpenVBSPP.ovpprt.nothing import NOTHING
    return value is NOTHING or value is None


def safe_member(obj, name):
    """`a?.b`：a 是 Nothing 时结果是 Nothing，不报错。"""
    if _is_nothing(obj):
        from OpenVBSPP.ovpprt.nothing import NOTHING
        return NOTHING
    return member_get(obj, name)


def safe_call(obj, name, refs):
    """`a?.b(args)`：a 是 Nothing 时直接返回 Nothing。"""
    if _is_nothing(obj):
        from OpenVBSPP.ovpprt.nothing import NOTHING
        return NOTHING
    return call_method(obj, name, refs)


def write_only_error(message):
    """读取只写属性时的错误。"""
    from OpenVBSPP.exceptions import OVPPObjectNotSupportedError
    raise OVPPObjectNotSupportedError(str(message))


def set_member(obj, name, value, is_set=False):
    """成员赋值：`Set obj.P = x` 走 Property Set，其余走 setter / 属性。"""
    if is_set:
        writer = getattr(obj, "_ovpp_pset_" + name, None)
        if writer is not None:
            writer(value)
            return
    setattr(obj, name, value)


def member_get(obj, name):
    """obj.name 取值：字段/属性直接返回，无参方法则调用（VBS 允许省略括号）。"""
    from OpenVBSPP.ovpprt.collections import Dictionary
    if isinstance(obj, Dictionary):
        key = _dict_key(obj, name)
        if key is not None:
            member = obj.item(key)
            return member() if callable(member) else member

    target = _lookup(obj, name)
    if target is _MISSING:
        raise OVPPObjectNotSupportedError(
            f"Object of type '{type(obj).__name__}' has no member '{name}'")
    if callable(target):
        return target()
    return target


def member_call(obj, name, *args):
    """obj.name(args...)：既可能是方法调用，也可能是数组字段的下标访问。"""
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
        return target(*args)

    # 不是方法：按数组下标处理
    if not args:
        return target
    for index in args[:-1]:
        target = target[index]
    return target[args[-1]]


def set_member_index(obj, name, args, value):
    """obj.name(i, j) = value —— 数组字段的元素赋值。"""
    from OpenVBSPP.ovpprt.collections import Dictionary
    if isinstance(obj, Dictionary):
        key = _dict_key(obj, name)
        if key is not None:
            member = obj.item(key)
            for index in args[:-1]:
                member = member[index]
            member[args[-1]] = value
            return

    target = _lookup(obj, name)
    if target is _MISSING:
        raise OVPPObjectNotSupportedError(
            f"Object of type '{type(obj).__name__}' has no member '{name}'")

    for index in args[:-1]:
        target = target[index]
    target[args[-1]] = value
