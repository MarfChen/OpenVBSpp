# OpenVBS++ 文档

**OpenVBS++** 是一门 VBScript 风格的脚本语言：语法是 VBScript 的超集，编译成标准
Python AST 后交给 CPython 执行。老 `.vbs` 脚本基本能直接跑，同时可以按需使用
类、`Try/Catch`、可选类型、包管理器 `over` 和 Windows 上真正的 COM。

- 项目主页 / 快速开始 / 命令行 / 性能：见 [GitHub 上的 README](https://github.com/MarfChen/OpenVBSpp)（[中文版](https://github.com/MarfChen/OpenVBSpp/blob/main/README.zh-CN.md)）
- 版本变更：见 [CHANGELOG](https://github.com/MarfChen/OpenVBSpp/blob/main/CHANGELOG.md)
- 安装：`pip install openvbspp`（Windows 需要 COM 时：`pip install "openvbspp[com]"`）

## 从这里开始

| 想知道什么 | 看这里 |
|---|---|
| 哪些 VBScript 写法支持、哪些不支持 | [VBScript 超集](VBS_SUPERSET.md) |
| 对象、属性、默认成员、`With` 怎么工作 | [对象模型](OBJECT_MODEL.md) |
| `Class` / `Inherits` / `Implements` / `Property` | [面向对象](OOP.md) |
| `Try/Catch`、`Option Typing`、`?.`、`??` 等新语法 | [现代特性](MODERN_FEATURES.md) |
| 怎么装第三方包、怎么写包、怎么锁版本 | [over 包管理器](OVER.md) |
| 内置网页 IDE、`/api/run`、`/api/check` | [网页 IDE](WEB_IDE.md) |
| VS Code 与其它编辑器支持 | [README — 编辑器 / IDE 支持](https://github.com/MarfChen/OpenVBSpp/blob/main/README.zh-CN.md#编辑器--ide-支持) |

## 一个小例子

```vbscript
' hello.vbs —— 用 openvbs hello.vbs 或 python -m OpenVBSPP hello.vbs 运行
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
    echo "windows COM 也可用：CreateObject(\"WScript.Shell\")"
End If
```

## 状态

- 测试：644 个用例全绿（Windows 3.10–3.13 硬门禁，Linux 同为门禁）
- 性能：端到端比 `cscript.exe` 快 2.5–3 倍；循环与数组已接近手写 Python
- 法律：独立实现，不分发微软代码或文档，见 [法务与合规](LEGAL.md)
