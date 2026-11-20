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


class ComCreateError(ComSupportError):
    """`CreateObject` could not create the class (VBScript error 429).

    中文：`CreateObject` 建不出对象，对应 VBScript 错误 429。"""

    #: VBScript error number the interpreter reports for this failure.
    #: 中文：解释器报告给脚本的 VBScript 错误号。
    number = 429
    #: VBScript error name, used in the error text and `Err.Description`.
    #: 中文：VBScript 错误名，用于错误文本与 `Err.Description`。
    vbs_name = "ActiveX component can't create object"


class ComAttachError(ComSupportError):
    """`GetObject` could not attach to an existing object (VBScript error 432).

    中文：`GetObject` 挂不上已有对象，对应 VBScript 错误 432。"""

    number = 432
    vbs_name = "File name or class name not found"


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


def _com_error_text(exc):
    """A readable message for a pywin32 `com_error`.

    pywin32 formats it as `(-2147221005, '无效的类字符串', None, None)`; VBScript would
    only show `无效的类字符串`, so the HRESULT keeps the detail.
    中文：把 pywin32 的 `com_error` 变成可读消息。pywin32 的格式是
    `(-2147221005, '无效的类字符串', None, None)`，VBS 只会显示 `无效的类字符串`，
    这里把 HRESULT 附在后面保留细节。
    """
    args = getattr(exc, "args", ())
    number = text = None
    if len(args) >= 2 and isinstance(args[1], str):
        number, text = args[0], args[1]
    elif args and isinstance(args[0], tuple) and len(args[0]) >= 2:
        number, text = args[0][0], args[0][1]
    text = str(text or "").strip()
    if text:
        try:
            code = int(number) & 0xFFFFFFFF
        except (TypeError, ValueError):
            return text
        return f"{text} (0x{code:08X})"
    return str(exc).strip()


def _dispatch_type_name(dispatch):
    """The COM type name from the type library, e.g. `FileSystemObject`.

    `TypeName(obj)` in VBScript answers with this name, not with `Object`.
    中文：从类型库取 COM 类型名（如 `FileSystemObject`）。VBS 的 `TypeName(obj)`
    返回的是它，而不是 `Object`。
    """
    try:
        import pythoncom
        oleobj = dispatch._oleobj_
        try:
            # IProvideClassInfo answers with the coclass name, exactly like
            # VBScript's TypeName ("FileSystemObject", "SWbemServicesEx").
            # 中文：IProvideClassInfo 给的是组件类名，和 VBS 的 TypeName 完全一致。
            info = oleobj.QueryInterface(pythoncom.IID_IProvideClassInfo)
            name = info.GetClassInfo().GetDocumentation(-1)[0]
            if name:
                return str(name).strip()
        # IProvideClassInfo is optional; most objects expose only the interface.
        # 中文：IProvideClassInfo 是可选的，多数对象只暴露接口的类型信息。
        except Exception:   # pragma: no cover - depends on the COM object
            pass
        info = oleobj.GetTypeInfo()
        if info is None:
            return ""
        name = info.GetDocumentation(-1)[0]
        return str(name or "").strip()
    # Not covered by tests: only a genuine COM object reaches here.
    # 中文：不纳入覆盖率，只有真实 COM 对象才会走到。
    except Exception:       # pragma: no cover
        return ""


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
        # TypeName(obj) reports the type-library name, not Object.
        # 中文：TypeName(obj) 用类型库里的名字，而不是 Object。
        object.__setattr__(self, "_vbs_name",
                           _dispatch_type_name(dispatch) or VBS_TYPE_NAME)

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

        # A property whose value is itself a COM object comes back as a dispatch that is
        # also callable (it has a default member); `obj.Files` must yield the collection
        # rather than call it. 中文：属性值是 COM 对象时，拿到的 dispatch 也可调用（有默认
        # 成员）；`obj.Files` 应当返回集合，而不是调用它。
        if not args and hasattr(member, "_oleobj_"):
            return _wrap(member)

        if not args and not callable(member):
            return _wrap(member)

        if not callable(member):
            raise ComSupportError(
                f"'{self._name}' is a property, not a method")

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
        detail = _com_error_text(exc)
        message = f"ActiveX component can't create object: '{progid}'"
        if detail:
            message += f" ({detail})"
        raise ComCreateError(message) from exc

    if dispatch is None:
        raise ComCreateError(
            f"ActiveX component can't create object: '{progid}'")
    return COMProxy(dispatch, str(progid))


def get(pathname="", classname=None):
    """The implementation behind `GetObject([pathname], [class])`.

    `pathname` may be a display-name moniker such as `winmgmts:`, a file name that
    binds to its class, or empty to attach to a class already in the running-object
    table. 中文：`GetObject([pathname], [class])` 的实现。`pathname` 可以是
    `winmgmts:` 这类显示名 moniker、绑定到类别的文件名，或留空以从运行对象表里按
    类名挂接。

    Returns a COMProxy. 中文：返回 COMProxy。
    """
    pythoncom, win32client = _require_pywin32()
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    path = str(pathname or "").strip()
    class_name = str(classname or "").strip()
    try:
        if path and class_name:
            dispatch = win32client.GetObject(path, class_name)
        elif path:
            dispatch = win32client.GetObject(path)
        elif class_name:
            try:
                dispatch = win32client.GetActiveObject(class_name)
            except Exception as exc:
                # A class with no running instance is 429 in VBScript, not 432.
                # 中文：类名挂接失败（没有运行实例）在 VBS 里是 429，不是 432。
                detail = _com_error_text(exc)
                message = f"ActiveX component can't create object: '{class_name}'"
                if detail:
                    message += f" ({detail})"
                raise ComCreateError(message) from exc
        else:
            raise ComAttachError(
                "GetObject needs a pathname, a class name, or both")
    except ComSupportError:
        raise
    except Exception as exc:
        detail = _com_error_text(exc)
        target = path or class_name
        message = f"File name or class name not found: '{target}'"
        if detail:
            message += f" ({detail})"
        raise ComAttachError(message) from exc

    if dispatch is None:
        raise ComAttachError(
            f"File name or class name not found: '{path or class_name}'")
    return COMProxy(dispatch, class_name or path)
