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

A minimal `Scripting.FileSystemObject`.

Half of a typical legacy VBS script reads and writes files, which is why this
subset matters for "old scripts keep working unchanged". Only plain text is
supported: no COM, no registry, no binary I/O.

中文：`Scripting.FileSystemObject` 的最小子集。老 VBS 脚本里一半的工作是读写文件，
which is why this module is the key to running old scripts unchanged. It covers plain text only: no COM, no registry,
中文：所以这一块是「老脚本不用改」的关键。只实现纯文本相关的能力——不碰 COM、不碰注册表、
Only plain text is covered: no COM, no registry, no binary I/O.
中文：只实现纯文本相关的能力，不碰 COM、不碰注册表、不做二进制读写。

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.CreateTextFile("out.txt", True)
  f.WriteLine "hello"
  f.Close
'''

import os
import tempfile
from pathlib import Path

from OpenVBSPP.exceptions import OVPPRuntimeError


# TextStream open modes / 中文：TextStream 的打开模式
FOR_READING = 1
FOR_WRITING = 2
FOR_APPENDING = 8


class TextStream:
    """The stream returned by `fso.CreateTextFile` / `OpenTextFile`.

    中文：`fso.CreateTextFile` / `OpenTextFile` 返回的流对象。
    """

    def __init__(self, path, mode, handle):
        self._path = str(path)
        self._mode = mode
        self._handle = handle
        self._closed = False
        self._line = 1
        self._column = 1

    # Reading / 中文：读
    def ReadAll(self):
        self._require_read()
        return self._handle.read()

    def ReadLine(self):
        self._require_read()
        text = self._handle.readline()
        if text.endswith("\n"):
            text = text[:-1]
        if text.endswith("\r"):
            text = text[:-1]
        self._line += 1
        self._column = 1
        return text

    def Read(self, count):
        self._require_read()
        return self._handle.read(int(count))

    def SkipLine(self):
        self.ReadLine()

    # Writing / 中文：写
    def Write(self, text):
        self._require_write()
        self._handle.write(str(text))
        self._column += len(str(text))

    def WriteLine(self, text=""):
        self._require_write()
        self._handle.write(str(text) + "\n")
        self._line += 1
        self._column = 1

    def WriteBlankLines(self, count):
        for _ in range(int(count)):
            self.WriteLine()

    # State / 中文：状态
    def Close(self):
        if not self._closed:
            self._handle.close()
            self._closed = True

    def AtEndOfStream(self):
        self._require_read()
        position = self._handle.tell()
        char = self._handle.read(1)
        self._handle.seek(position)
        return char == ""

    Line = property(lambda self: self._line)
    Column = property(lambda self: self._column)

    def __del__(self):
        try:
            self.Close()
        except Exception:
            pass

    # Internals / 中文：内部
    def _require_read(self):
        if self._mode not in (FOR_READING,):
            raise OVPPRuntimeError("TextStream is not open for reading")
        if self._closed:
            raise OVPPRuntimeError("TextStream is closed")

    def _require_write(self):
        if self._mode not in (FOR_WRITING, FOR_APPENDING):
            raise OVPPRuntimeError("TextStream is not open for writing")
        if self._closed:
            raise OVPPRuntimeError("TextStream is closed")


class FileSystemObject:
    """A subset of `Scripting.FileSystemObject`: text files, paths, existence checks.

    中文：`Scripting.FileSystemObject` 的子集——文本文件 + 路径 + 存在性判断。
    """

    # Existence checks / 中文：判断
    def FileExists(self, path):
        return Path(str(path)).is_file()

    def FolderExists(self, path):
        return Path(str(path)).is_dir()

    def DriveExists(self, path):
        return Path(str(path)).exists()

    # Folders / 中文：目录
    def CreateFolder(self, path):
        target = Path(str(path))
        if target.exists():
            raise OVPPRuntimeError(f"Folder already exists: {target}")
        target.mkdir(parents=True)
        return Folder(target)

    def DeleteFolder(self, path, force=None):
        target = Path(str(path))
        if not target.is_dir():
            raise OVPPRuntimeError(f"Folder not found: {target}")
        if force:
            for child in sorted(target.rglob("*"), reverse=True):
                child.unlink() if child.is_file() else child.rmdir()
        target.rmdir()

    def GetFolder(self, path):
        target = Path(str(path))
        if not target.is_dir():
            raise OVPPRuntimeError(f"Folder not found: {target}")
        return Folder(target)

    def GetParentFolderName(self, path):
        return str(Path(str(path)).parent)

    # Paths / 中文：路径
    def GetFileName(self, path):
        return Path(str(path)).name

    def GetBaseName(self, path):
        return Path(str(path)).stem

    def GetExtensionName(self, path):
        return Path(str(path)).suffix.lstrip(".")

    def BuildPath(self, folder, name):
        return str(Path(str(folder)) / str(name))

    def GetTempName(self):
        handle, name = tempfile.mkstemp(suffix=".tmp")
        os.close(handle)
        return name

    # Files / 中文：文件
    def DeleteFile(self, path, force=None):
        target = Path(str(path))
        if not target.is_file():
            raise OVPPRuntimeError(f"File not found: {target}")
        try:
            target.unlink()
        except OSError as exc:
            raise OVPPRuntimeError(f"DeleteFile failed: {exc}") from exc

    def CreateTextFile(self, path, overwrite=True, unicode=None):
        target = Path(str(path))
        if target.exists() and not overwrite:
            raise OVPPRuntimeError(f"File already exists: {target}")
        target.parent.mkdir(parents=True, exist_ok=True)
        handle = open(target, "w", encoding="utf-8", newline="\n")
        return TextStream(target, FOR_WRITING, handle)

    def OpenTextFile(self, path, iomode=FOR_READING, create=False):
        target = Path(str(path))
        mode = int(iomode or FOR_READING)
        if mode == FOR_READING:
            if not target.is_file():
                raise OVPPRuntimeError(f"File not found: {target}")
            handle = open(target, "r", encoding="utf-8-sig", newline="")
        elif mode == FOR_APPENDING:
            target.parent.mkdir(parents=True, exist_ok=True)
            handle = open(target, "a", encoding="utf-8", newline="\n")
        else:
            if not target.exists() and not create:
                raise OVPPRuntimeError(f"File not found: {target}")
            target.parent.mkdir(parents=True, exist_ok=True)
            handle = open(target, "w", encoding="utf-8", newline="\n")
        return TextStream(target, mode, handle)

    def CopyFile(self, source, target, overwrite=True):
        import shutil
        destination = Path(str(target))
        if destination.exists() and not overwrite:
            raise OVPPRuntimeError(f"File already exists: {destination}")
        shutil.copyfile(str(source), str(destination))

    def MoveFile(self, source, target):
        import shutil
        shutil.move(str(source), str(target))

    def __repr__(self):
        return "FileSystemObject()"


class Folder:
    """The folder object returned by `fso.GetFolder` / `CreateFolder` (minimal).

    中文：`fso.GetFolder` / `CreateFolder` 返回的文件夹对象（最小集）。
    """

    def __init__(self, path):
        self._path = Path(path)

    def Files(self):
        return [File(item) for item in sorted(self._path.glob("*")) if item.is_file()]

    def SubFolders(self):
        return [Folder(item) for item in sorted(self._path.glob("*")) if item.is_dir()]

    Name = property(lambda self: self._path.name)
    Path = property(lambda self: str(self._path))
    ShortPath = property(lambda self: str(self._path))

    def __repr__(self):
        return f"Folder({self._path})"


class File:
    """The file object returned by `fso.GetFile` (minimal).

    中文：`fso.GetFile` 返回的文件对象（最小集）。
    """

    def __init__(self, path):
        self._path = Path(path)

    def Delete(self):
        self._path.unlink()

    Name = property(lambda self: self._path.name)
    Path = property(lambda self: str(self._path))
    Size = property(lambda self: self._path.stat().st_size if self._path.exists() else 0)

    def __repr__(self):
        return f"File({self._path})"
