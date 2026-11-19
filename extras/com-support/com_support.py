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

com-support —— OpenVBS++ 的 Windows COM/DCOM 适配层。

它是一个 **独立的 over 包**：包内只有 `over.json` 与这个文件，不 import
`OpenVBSPP` 的任何模块，因此可以被单独复制、打包、卸载。

使用方式（脚本里不需要 import）：
    Set shell = CreateObject("WScript.Shell")
    shell.Run "cmd /c echo hi"
    Set fso = CreateObject("Scripting.FileSystemObject")   ' 内置实现优先

约定
----
`create(progid, *args)` 是唯一入口，返回一个「假对象」：`OpenVBSPP` 的
对象层通过 `_vbs_name` 认出它，成员访问走 `__getattr__`，属性赋值走
`__setattr__`，因此 VBS 侧的 `obj.Prop` / `obj.Method(a, b)` 都能自然工作。

已知限制：`ByRef` 参数的写回（COM 侧改值后回写 VBS 变量）不支持。
'''

import sys

VERSION = "1.0.0"

#: 允许 VBS 侧 `TypeName(obj)` 看到的名字
VBS_TYPE_NAME = "Object"


class ComSupportError(Exception):
    """COM 支持相关的错误（ProgID 不存在、pywin32 缺失等）。"""


def is_available():
    """当前解释器是否具备 COM 能力（能 import pywin32 即为真）。"""
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
    except ImportError as exc:      # pragma: no cover - 取决于环境
        raise ComSupportError(
            "COM support needs pywin32: run 'pip install pywin32' "
            "(or 'openvbs over install' inside a project)") from exc
    return pythoncom, win32client


# --------------------------------------------------------------------- 值转换
def _unwrap(value):
    """把 VBS 侧的值翻译成 COM 能接受的 Python 值。"""
    if isinstance(value, COMProxy):
        return object.__getattribute__(value, "_com_dispatch")
    return value


def _is_dispatch(value):
    return hasattr(value, "_oleobj_") or type(value).__name__ in (
        "CDispatch", "Dispatch", "DispatchEx")


def _wrap(value, progid=""):
    """把 COM 返回值翻译成 VBS 侧的值（IDispatch -> COMProxy）。"""
    if isinstance(value, (tuple, list)):
        return type(value)(_wrap(item, progid) for item in value)
    if _is_dispatch(value):
        return COMProxy(value, progid)
    return value


def _convert_error(exc):
    """把 pywin32 的 com_error 变成一条对 VBS 用户友好的消息。"""
    text = str(exc).strip() or exc.__class__.__name__
    return ComSupportError(f"COM automation failed: {text}")


# --------------------------------------------------------------------- 代理
class COMProxy:
    """一个 COM 对象的 Python 侧外壳（`_vbs_name` 让对象层认得出它）。"""

    _vbs_name = VBS_TYPE_NAME

    def __init__(self, dispatch, progid=""):
        object.__setattr__(self, "_com_dispatch", dispatch)
        object.__setattr__(self, "_com_progid", progid)

    # -- 成员 ----------------------------------------------------------
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

    # -- 集合行为（For Each / obj(i)） ----------------------------------
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
        """`obj(args)` —— 调用 COM 的默认成员。"""
        dispatch = object.__getattribute__(self, "_com_dispatch")
        try:
            return _wrap(dispatch(*[_unwrap(item) for item in args]))
        except Exception as exc:
            raise _convert_error(exc) from exc

    # -- 显示 ----------------------------------------------------------
    def __repr__(self):
        progid = object.__getattribute__(self, "_com_progid")
        return f"<COM {progid or 'object'}>"

    def __str__(self):
        return self.__repr__()

    def __bool__(self):
        return True


class _ComMember:
    """延迟到调用时才真正访问 COM 成员的中间层。

    `obj.Text`  —— `member_get` 看到可调用对象会以 0 参调用 -> 读属性；
    `obj.Run "x"` —— `call_method` 以实参调用 -> 调方法。
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


# --------------------------------------------------------------------- 入口
def create(progid, *args):
    """CreateObject(progid) 的 COM 实现。"""
    if args:
        raise ComSupportError(
            "CreateObject(class, servername) is not supported by com-support")
    if not str(progid).strip():
        raise ComSupportError("CreateObject needs a non-empty ProgID")

    pythoncom, win32client = _require_pywin32()
    try:
        pythoncom.CoInitialize()
    except Exception:               # 已在其它线程初始化 / 已初始化
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
    """GetObject(progid) 的近似实现（拿到已运行实例）。"""
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
