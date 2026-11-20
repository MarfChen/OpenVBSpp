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

com-support is the Windows COM/DCOM adapter for OpenVBS++.

中文：com-support 是 OpenVBS++ 的 Windows COM/DCOM 适配层。

It is a **standalone over package**: the directory holds only `over.json` and this
file, it imports nothing from `OpenVBSPP`, and it can be copied, packed or removed
on its own.

中文：它是一个独立的 over 包：包内只有 `over.json` 与这个文件，不 import
`OpenVBSPP` 的任何模块，因此可以被单独复制、打包、卸载。

Usage, with no import inside the script:

中文：使用方式（脚本里不需要 import）：

    Set shell = CreateObject("WScript.Shell")
    shell.Run "cmd /c echo hi"
    Set fso = CreateObject("Scripting.FileSystemObject")   ' built-in wins
                                                           ' 内置实现优先

Conventions

中文：约定

`create(progid, *args)` is the only entry point and returns a fake object: the
OpenVBSPP object layer recognises it through `_vbs_name`, member access goes through
`__getattr__` and member assignment through `__setattr__`, so `obj.Prop` and
`obj.Method(a, b)` both work from VBS.

中文：`create(progid, *args)` 是唯一入口，返回一个「假对象」；对象层通过
`_vbs_name` 认出它，成员访问走 `__getattr__`，属性赋值走 `__setattr__`，
因此 VBS 侧的 `obj.Prop` / `obj.Method(a, b)` 都能自然工作。

Known limit: ByRef write-back is not supported, so a value COM changes is not
written back into the VBS variable.

中文：已知限制：`ByRef` 参数的写回（COM 侧改值后回写 VBS 变量）不支持。
'''

import sys

VERSION = "1.0.0"

#: The name a script sees from `TypeName(obj)`
#: 中文：允许 VBS 侧 `TypeName(obj)` 看到的名字
VBS_TYPE_NAME = "Object"


class ComSupportError(Exception):
    """A COM support error: unknown ProgID, pywin32 missing and so on.

    中文：COM 支持相关的错误，例如 ProgID 不存在、pywin32 缺失。"""


def is_available():
    """Whether this interpreter can do COM: true when pywin32 imports.

    中文：当前解释器是否具备 COM 能力，能 import pywin32 即为真。"""
    if sys.platform != "win32":
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


def _require_pywin32():
    if sys.platform != "win32":
        raise ComSupportError(
            "COM is only available on Windows "
            f"(current platform: {sys.platform})")
    try:
        import pythoncom
        import win32com.client as win32client
    # Not covered by tests: it depends on the machine's Python setup
    # 中文：不纳入覆盖率，取决于环境
    except ImportError as exc:      # pragma: no cover
        raise ComSupportError(
            "COM support needs pywin32: run 'pip install pywin32' "
            "(or 'openvbs over install' inside a project)") from exc
    return pythoncom, win32client


# ------------------------------------------------------------ value conversion
# 中文：值转换
def _unwrap(value):
    """Translate a script-side value into something COM accepts.

    中文：把脚本侧的值翻译成 COM 能接受的 Python 值。"""
    if isinstance(value, COMProxy):
        return object.__getattribute__(value, "_com_dispatch")
    return value


def _is_dispatch(value):
    return hasattr(value, "_oleobj_") or type(value).__name__ in (
        "CDispatch", "Dispatch", "DispatchEx")


def _wrap(value, progid=""):
    """Translate a COM return value back, wrapping IDispatch in a COMProxy.

    中文：把 COM 返回值翻译成脚本侧的值，IDispatch 包成 COMProxy。"""
    if isinstance(value, (tuple, list)):
        return type(value)(_wrap(item, progid) for item in value)
    if _is_dispatch(value):
        return COMProxy(value, progid)
    return value


def _convert_error(exc):
    """Turn a pywin32 com_error into a message a script author can act on.

    中文：把 pywin32 的 com_error 变成一条对脚本作者友好的消息。"""
    text = str(exc).strip() or exc.__class__.__name__
    return ComSupportError(f"COM automation failed: {text}")


# --------------------------------------------------------------------- proxy
# 中文：代理
class COMProxy:
    """A Python-side shell around one COM object; `_vbs_name` makes it recognisable.

    中文：一个 COM 对象的 Python 侧外壳，`_vbs_name` 让对象层认得出它。"""

    _vbs_name = VBS_TYPE_NAME

    def __init__(self, dispatch, progid=""):
        object.__setattr__(self, "_com_dispatch", dispatch)
        object.__setattr__(self, "_com_progid", progid)

    # -- members -------------------------------------------------------
    # 中文：成员
    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        return _ComMember(self, name)

    def __setattr__(self, name, value):
        if name.startswith("_"):
            object.__setattr__(self, name, value)
            return
        dispatch = object.__getattribute__(self, "_com_dispatch")
        try:
            setattr(dispatch, name, _unwrap(value))
        except Exception as exc:
            raise _convert_error(exc) from exc

    # -- collection behaviour (For Each / obj(i)) -----------------------
    # 中文：集合行为（For Each / obj(i)）
    def __iter__(self):
        dispatch = object.__getattribute__(self, "_com_dispatch")
        try:
            return iter(_wrap(list(dispatch)))
        except Exception as exc:
            raise _convert_error(exc) from exc

    def __len__(self):
        dispatch = object.__getattribute__(self, "_com_dispatch")
        try:
            return len(dispatch)
        except Exception as exc:
            raise _convert_error(exc) from exc

    def __getitem__(self, key):
        dispatch = object.__getattribute__(self, "_com_dispatch")
        try:
            return _wrap(dispatch[_unwrap(key)])
        except Exception as exc:
            raise _convert_error(exc) from exc

    def __call__(self, *args):
        """`obj(args)` calls COM's default member.

        中文：调用 COM 的默认成员。"""
        dispatch = object.__getattribute__(self, "_com_dispatch")
        try:
            return _wrap(dispatch(*[_unwrap(item) for item in args]))
        except Exception as exc:
            raise _convert_error(exc) from exc

    # -- display -------------------------------------------------------
    # 中文：显示
    def __repr__(self):
        progid = object.__getattribute__(self, "_com_progid")
        return f"<COM {progid or 'object'}>"

    def __str__(self):
        return self.__repr__()

    def __bool__(self):
        return True


class _ComMember:
    """An intermediate layer that only touches COM when it is called.

    `obj.Text`   -- member_get sees a callable and calls it with no argument, so it reads
                    the property;
    中文：`obj.Text` —— member_get 看到可调用对象会以 0 参调用，即读属性；
    `obj.Run "x"` -- call_method calls it with the arguments, so it invokes the method.
    中文：`obj.Run "x"` —— call_method 以实参调用，即调方法。
    """

    __slots__ = ("_owner", "_name")

    def __init__(self, owner, name):
        self._owner = owner
        self._name = name

    def __call__(self, *args):
        dispatch = object.__getattribute__(self._owner, "_com_dispatch")
        try:
            member = getattr(dispatch, self._name)
        except Exception as exc:
            raise _convert_error(exc) from exc

        if not callable(member):
            if args:
                raise ComSupportError(
                    f"'{self._name}' is a property, not a method")
            return _wrap(member)

        try:
            return _wrap(member(*[_unwrap(item) for item in args]))
        except Exception as exc:
            raise _convert_error(exc) from exc

    def __repr__(self):
        return f"<COM member {self._name}>"


# --------------------------------------------------------------------- entry
# 中文：入口
def create(progid, *args):
    """The COM implementation behind CreateObject(progid).

    中文：CreateObject(progid) 的 COM 实现。"""
    if args:
        raise ComSupportError(
            "CreateObject(class, servername) is not supported by com-support")
    if not str(progid).strip():
        raise ComSupportError("CreateObject needs a non-empty ProgID")

    pythoncom, win32client = _require_pywin32()
    try:
        pythoncom.CoInitialize()
    # Already initialised, possibly on another thread
    # 中文：已在其它线程初始化 / 已初始化
    except Exception:
        pass

    factory = getattr(win32client, "DispatchEx", win32client.Dispatch)
    try:
        dispatch = factory(str(progid))
    except Exception as exc:
        raise _convert_error(exc) from exc

    if dispatch is None:
        raise ComSupportError(f"COM class '{progid}' could not be created")
    return COMProxy(dispatch, str(progid))


def get(progid, *args):
    """An approximation of GetObject(progid): attach to a running instance.

    中文：GetObject(progid) 的近似实现，拿到已运行实例。"""
    pythoncom, win32client = _require_pywin32()
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        dispatch = win32client.GetActiveObject(str(progid))
    except Exception as exc:
        raise _convert_error(exc) from exc
    return COMProxy(dispatch, str(progid))
