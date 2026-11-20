"""
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

The `WScript` host object.

VBScript files are normally run by Windows Script Host, which hands the script a
global `WScript` object: `WScript.Echo` writes a line, `WScript.Arguments` holds
the command line, `WScript.ScriptName` names the running file, and so on.  Old
VBScript code uses these objects heavily, so OpenVBS++ provides a console host
with the same surface.  The object is injected by the transformer under the name
`_ovpp_host.HOST`; `bind()` is called once per run to record which file is being
executed and which arguments it received.

中文：`WScript` 宿主对象。VBScript 文件通常由 Windows Script Host 执行，它会向脚本提供
一个全局的 `WScript` 对象：`WScript.Echo` 输出一行，`WScript.Arguments` 保存命令行参数，
`WScript.ScriptName` 是正在运行的文件名，等等。老 VBScript 代码大量使用这些对象，因此
OpenVBS++ 提供一个具有相同接口的控制台宿主。该对象由 transformer 以
`_ovpp_host.HOST` 的名字注入；每次运行前调用 `bind()` 记录正在执行的文件和收到的参数。
"""

import os
import sys
import time


#: The error number of "Input past end of file", as VBScript reports it.
#: 中文：VBScript 里「输入超出文件尾」的错误号。
ERR_INPUT_PAST_EOF = 62

#: The error number of "Subscript out of range".
#: 中文：「下标越界」的错误号。
ERR_SUBSCRIPT_OUT_OF_RANGE = 9


def _to_text(value):
    """Convert a VBScript value to the text a stream writes.

    中文：把 VBS 值转换成流写出的文本。
    """
    from OpenVBSPP.ovpprt.builtins import to_text
    return to_text(value)


def _error(number, message):
    """Build a runtime error carrying a VBScript error number.

    中文：构造一个带 VBS 错误号的运行时错误。
    """
    from OpenVBSPP.exceptions import OVPPRuntimeError
    error = OVPPRuntimeError(message)
    error.number = number
    error.name = message
    return error


class _TextStream:
    """A `TextStream`: the type of `WScript.StdOut`, `StdErr` and `StdIn`.

    The console streams are thin wrappers around `sys.stdout`, `sys.stderr` and
    `sys.stdin`.  Reads are only meaningful for `StdIn`; writing to a read-only
    stream is allowed by VBScript but pointless here, so it succeeds quietly.

    中文：`TextStream` —— `WScript.StdOut`、`StdErr`、`StdIn` 的类型。控制台流是
    `sys.stdout`、`sys.stderr`、`sys.stdin` 的薄包装。读取只对 `StdIn` 有意义；向只读流
    写入在 VBScript 里合法但没有实际效果，这里也安静地成功。
    """

    _vbs_name = "TextStream"

    def __init__(self, stream):
        self._stream = stream
        self._pending = None

    # ------------------------------------------------------------------ writing
    # 中文：写
    def Write(self, text=""):
        """Write text without a trailing newline.

        中文：写文本，不换行。
        """
        self._emit(text, "")

    def WriteLine(self, text=""):
        """Write text followed by a newline.

        中文：写文本并换行。
        """
        self._emit(text, "\n")

    def WriteBlankLines(self, count=0):
        """Write `count` empty lines.

        中文：写出 `count` 个空行。
        """
        try:
            total = int(count or 0)
        except (TypeError, ValueError):
            total = 0
        for _ in range(max(0, total)):
            self._emit("", "\n")

    def _emit(self, text, newline):
        payload = _to_text(text) + newline
        try:
            self._stream.write(payload)
            self._stream.flush()
        except UnicodeEncodeError:
            # Old Windows consoles default to GBK/CP1252; degrade instead of raising.
            # 中文：老式 Windows 控制台默认 GBK/CP1252，这里降级输出而不是抛错。
            buffer = getattr(self._stream, "buffer", None)
            if buffer is None:
                raise
            encoding = getattr(self._stream, "encoding", None) or "utf-8"
            buffer.write(payload.encode(encoding, "backslashreplace"))
            buffer.flush()

    # ------------------------------------------------------------------ reading
    # 中文：读
    def ReadLine(self):
        """Read one line without its line terminator; at end of file it raises 62.

        中文：读取一行（不含行尾）；到文件尾时抛错误 62。
        """
        first = self._take_pending()
        rest = self._stream.readline()
        line = first + rest
        if line == "":
            raise _error(ERR_INPUT_PAST_EOF, "Input past end of file")
        return line.rstrip("\r\n")

    def ReadAll(self):
        """Read the rest of the stream.

        中文：读取剩余的流内容。
        """
        first = self._take_pending()
        return first + self._stream.read()

    def Skip(self, characters=0):
        """Read and discard `characters` characters.

        中文：读取并丢弃 `characters` 个字符。
        """
        try:
            count = int(characters or 0)
        except (TypeError, ValueError):
            count = 0
        if self._pending is not None:
            if count <= 0:
                return
            self._pending = None
            count -= 1
        if count > 0:
            self._stream.read(count)

    def SkipLine(self):
        """Read and discard one line.

        中文：读取并丢弃一行。
        """
        self._take_pending()
        self._stream.readline()

    @property
    def AtEndOfStream(self):
        """Whether there is nothing left to read.

        中文：是否已经读到末尾。
        """
        return self._peek() == ""

    def Close(self):
        """Close the stream. The console streams stay open, so this does nothing.

        中文：关闭流。控制台流由进程持有，因此这里什么也不做。
        """

    # ---------------------------------------------------------------- internals
    # 中文：内部实现
    def _peek(self):
        if self._pending is None:
            self._pending = self._stream.read(1)
        return self._pending

    def _take_pending(self):
        if self._pending is None:
            return ""
        pending = self._pending
        self._pending = None
        return pending

    def __repr__(self):
        name = getattr(self._stream, "name", "")
        return f"_TextStream({name})"


class _NamedArguments:
    """`WScript.Arguments.Named`: the `/name:value` arguments, keyed by name.

    中文：`WScript.Arguments.Named` —— `/name:value` 形式的具名参数。
    """

    _vbs_name = "Object"

    def __init__(self, values):
        self._values = {}
        for raw in values:
            name, value = _split_named(raw)
            if name is not None:
                self._values[name.lower()] = value

    @property
    def Count(self):
        """The number of named arguments.

        中文：具名参数的个数。
        """
        return len(self._values)

    def Exists(self, name):
        """Whether a named argument was supplied.

        中文：某个具名参数是否存在。
        """
        return str(name).lower() in self._values

    def Item(self, name=""):
        """The value of a named argument; a missing name yields an empty string.

        中文：具名参数的值；名字不存在时返回空字符串。
        """
        return self._values.get(str(name).lower(), "")

    def __getitem__(self, name):
        return self.Item(name)

    def __iter__(self):
        return iter(list(self._values))

    def __len__(self):
        return len(self._values)

    def __contains__(self, name):
        return self.Exists(name)

    def __repr__(self):
        return f"_NamedArguments({self._values!r})"


class _Arguments:
    """`WScript.Arguments`: the command line the script was started with.

    VBScript indexes arguments from zero, and `Count` returns how many there are.
    `Named` exposes `/name:value` style arguments.

    中文：`WScript.Arguments` —— 启动脚本时传入的命令行参数。VBScript 从 0 开始编号，
    `Count` 是参数个数；`Named` 提供 `/name:value` 风格的具名参数。
    """

    _vbs_name = "Object"

    def __init__(self, values=()):
        self._values = [str(value) for value in values]

    @property
    def Count(self):
        """The number of arguments.

        中文：参数个数。
        """
        return len(self._values)

    def Item(self, index=0):
        """The argument at `index` (zero based).

        中文：`index` 位置的参数（从 0 开始）。
        """
        try:
            return self._values[int(index)]
        except (IndexError, TypeError, ValueError):
            raise _error(ERR_SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range")

    def __getitem__(self, index):
        return self.Item(index)

    def __iter__(self):
        return iter(self._values)

    def __len__(self):
        return len(self._values)

    @property
    def Named(self):
        """The named (`/name:value`) arguments.

        中文：具名（`/name:value`）参数。
        """
        return _NamedArguments(self._values)

    def __repr__(self):
        return f"_Arguments({self._values!r})"


def _split_named(raw):
    """Split `/name:value` or `/name=value` into a name and a value.

    Anything else is not a named argument and yields `(None, None)`.

    中文：把 `/name:value` 或 `/name=value` 拆成名字和值；其它形式不是具名参数，
    返回 `(None, None)`。
    """
    text = str(raw)
    if not text.startswith("/") or len(text) < 3:
        return None, None
    body = text[1:]
    for separator in (":", "="):
        index = body.find(separator)
        if index > 0:
            return body[:index], body[index + 1:]
    return None, None


class _Host:
    """The `WScript` object itself.

    The properties mirror Windows Script Host, but they describe the OpenVBS++
    console runner instead of `cscript.exe`: `FullName` is the Python
    interpreter, `Path` its directory, and dialogs are never interactive.

    中文：`WScript` 对象本身。各属性仿照 Windows Script Host，但描述的是 OpenVBS++
    的控制台运行器而不是 `cscript.exe`：`FullName` 是 Python 解释器，`Path` 是它所在
    的目录，且没有交互式对话框。
    """

    _vbs_name = "Object"

    def __init__(self):
        self._script = ""
        self._arguments = _Arguments()
        self._timeout = 0

    # ---------------------------------------------------------------- identity
    # 中文：身份信息
    @property
    def Name(self):
        """The name of the host, as VBScript reports it.

        中文：宿主名称（与 VBScript 报告的一致）。
        """
        return "Windows Script Host"

    @property
    def Version(self):
        """The host version VBScript scripts expect.

        中文：VBScript 脚本期望的宿主版本号。
        """
        return "5.812"

    @property
    def ScriptName(self):
        """The file name of the running script, without its directory.

        中文：正在运行的脚本文件名（不含目录）。
        """
        if not self._script:
            return ""
        return os.path.basename(self._script)

    @property
    def ScriptFullName(self):
        """The full path of the running script.

        中文：正在运行的脚本完整路径。
        """
        if not self._script:
            return ""
        try:
            return os.path.abspath(self._script)
        except (OSError, ValueError):
            return self._script

    @property
    def FullName(self):
        """The full path of the program running the script.

        中文：运行脚本的程序的完整路径。
        """
        return sys.executable or ""

    @property
    def Path(self):
        """The directory holding the program running the script.

        中文：运行脚本的程序所在目录。
        """
        executable = sys.executable or ""
        return os.path.dirname(executable)

    @property
    def Interactive(self):
        """Whether the host can show dialogs (a console session with a terminal).

        中文：宿主是否处于可以弹窗的交互环境（有终端的控制台会话）。
        """
        stdin = sys.stdin
        stdout = sys.stdout
        return bool(stdin is not None and stdout is not None
                    and stdin.isatty() and stdout.isatty())

    @property
    def Timeout(self):
        """How long a script may run, in seconds; `0` means no limit.

        中文：脚本最长运行时间（秒）；`0` 表示不限制。
        """
        return self._timeout

    @Timeout.setter
    def Timeout(self, value):
        try:
            self._timeout = int(value or 0)
        except (TypeError, ValueError):
            self._timeout = 0

    @property
    def Arguments(self):
        """The command line arguments of the running script.

        中文：正在运行的脚本的命令行参数。
        """
        return self._arguments

    @property
    def StdIn(self):
        """Standard input as a `TextStream`.

        中文：标准输入（`TextStream`）。
        """
        return _TextStream(sys.stdin)

    @property
    def StdOut(self):
        """Standard output as a `TextStream`.

        中文：标准输出（`TextStream`）。
        """
        return _TextStream(sys.stdout)

    @property
    def StdErr(self):
        """Standard error as a `TextStream`.

        中文：标准错误（`TextStream`）。
        """
        return _TextStream(sys.stderr)

    # ------------------------------------------------------------------ methods
    # 中文：方法
    def Echo(self, *values):
        """Print the values separated by spaces, like the `Echo` statement.

        中文：按 `Echo` 语句的习惯，用空格分隔输出各值。
        """
        from OpenVBSPP.ovpprt.builtins import Echo as _echo
        _echo(*values)

    def Sleep(self, milliseconds=0):
        """Pause the script for the given number of milliseconds.

        中文：让脚本暂停指定的毫秒数。
        """
        try:
            delay = float(milliseconds or 0) / 1000.0
        except (TypeError, ValueError):
            delay = 0.0
        if delay > 0:
            time.sleep(delay)

    def Quit(self, exit_code=0):
        """Stop the script and exit with the given code.

        中文：结束脚本，并以给定退出码退出。
        """
        try:
            code = int(exit_code or 0)
        except (TypeError, ValueError):
            code = 0
        raise SystemExit(code)

    def CreateObject(self, progid="", prefix=""):
        """Create a COM object; `prefix` names the event handlers and is ignored.

        中文：创建 COM 对象；`prefix` 是事件处理函数前缀，这里忽略。
        """
        from OpenVBSPP.ovpprt.object import create_object
        return create_object(progid, {})

    def GetObject(self, pathname="", classname=None):
        """Attach to a running COM object, like the `GetObject` function.

        中文：挂接已运行的 COM 对象，与 `GetObject` 函数一致。
        """
        from OpenVBSPP.ovpprt.object import get_object
        return get_object(pathname, {}, classname)

    def __repr__(self):
        return f"_Host({self._script!r}, {self._arguments._values!r})"


#: The single host object shared by every script in the process.
#: 中文：进程内所有脚本共享的唯一宿主对象。
HOST = _Host()


def bind(filename=None, arguments=None):
    """Record which script is running and with which arguments.

    The runtime calls this before executing a program.  Tests and embedders may
    call it directly to simulate a command line.

    中文：记录正在运行的脚本以及它的参数。运行时在执行程序前调用它；测试与嵌入方也
    可以直接调用来模拟命令行。
    """
    HOST._script = str(filename or "")
    HOST._arguments = _Arguments(arguments or ())


def describe():
    """A one-line description of the host, for `openvbs com` style listings.

    中文：宿主的一行描述，用于 `openvbs com` 之类的清单输出。
    """
    return (f"WScript host: {HOST.Name} {HOST.Version} "
            f"(console; script arguments: {HOST.Arguments.Count})")
