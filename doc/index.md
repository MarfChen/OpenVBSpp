# OpenVBS++ documentation

> English | **[中文](index.zh-CN.md)**

**OpenVBS++** is a VBScript-flavoured scripting language: its syntax is a superset of
VBScript, compiled to a standard Python AST and executed by CPython. Existing `.vbs`
scripts mostly run as they are, while classes, `Try/Catch`, optional typing, the `over`
package manager and real COM on Windows are available when you want them.

- Project home / quick start / command line / performance: see the
  [README on GitHub](https://github.com/MarfChen/OpenVBSpp) ([中文版](https://github.com/MarfChen/OpenVBSpp/blob/main/README.zh-CN.md))
- What changed between versions: see the [CHANGELOG](https://github.com/MarfChen/OpenVBSpp/blob/main/CHANGELOG.md)
- Install: `pip install openvbspp` (on Windows with COM: `pip install "openvbspp[com]"`)

## Start here

| What you want to know | Read |
|---|---|
| Which VBScript constructs are supported and which are not | [VBScript superset](VBS_SUPERSET.md) |
| How objects, properties, default members and `With` work | [Object model](OBJECT_MODEL.md) |
| `Class` / `Inherits` / `Implements` / `Property` | [Object-oriented programming](OOP.md) |
| `Try/Catch`, `Option Typing`, `?.`, `??` and other new syntax | [Modern features](MODERN_FEATURES.md) |
| How to install third-party packages, publish your own, and lock versions | [The over package manager](OVER.md) |
| The built-in web IDE, `/api/run`, `/api/check` | [Web IDE](WEB_IDE.md) |
| Editor support (VS Code and others) | [README — editor and IDE support](https://github.com/MarfChen/OpenVBSpp#editor-and-ide-support) |

## A small example

```vbscript
' hello.vbs — run with openvbs hello.vbs or python -m OpenVBSPP hello.vbs
Class Counter
    Private m_n
    Public Property Get Value()
        Value = m_n
    End Property
    Public Sub Add(ByVal n)
        m_n = m_n + n
    End Sub
End Class

Set c = New Counter
c.Add 2
c.Add 3
echo "total = " & c.Value

If IsObject(c) Then
    echo "Windows COM works too: CreateObject(""WScript.Shell"")"
End If
```

## Status

- Tests: 557 cases green (hard gate on Windows 3.10–3.13, Linux is a gate as well)
- Performance: 2.5–3× faster end to end than `cscript.exe`; loops and arrays are on par with hand-written Python
- Legal: an independent implementation that does not distribute Microsoft code or documentation — see [Legal and compliance](LEGAL.md)
