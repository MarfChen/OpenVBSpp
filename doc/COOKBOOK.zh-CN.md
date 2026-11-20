> **[English](COOKBOOK.md)** | 中文

# 实战手册（Cookbook）

这里每个配方都是可以完整运行的脚本：存成 `recipe.vbs`，然后 `openvbs recipe.vbs`（或
`python -m OpenVBSPP recipe.vbs`）即可。标了 **Windows** 的配方会调 COM，只能在 Windows 上跑，
其余配方在任何能跑 CPython 的系统上都行。

如果只看一节，看 [迁移老脚本](#迁移老脚本)：那里列的是真正会绊住人的几处差异。

---

## 快速运行

```powershell
# 运行文件
openvbs hello.vbs

# 不建文件，直接跑一行
openvbs -c "echo ""hello"" & "" from the command line"""

# 从标准输入读（管道里很方便）
Get-Content hello.vbs | openvbs -

# 只解析、输出 JSON 诊断（有错误时退出码为 1）
openvbs check broken.vbs --json
```

`check` 只编译不运行，`run` 编译后再运行——编辑器和 CI 用的就是这两个。`openvbs --version`
打印版本号。输出是纯文本，Tab 与 ANSI 颜色都能正常工作，所以 `| Out-File log.txt` 和
`> log.txt` 都可以。

---

## 参数、退出码与宿主对象

```vbscript
' args.vbs —— 运行：openvbs args.vbs alpha beta /mode:fast
WScript.Echo "script: " & WScript.ScriptName
WScript.Echo "args: " & WScript.Arguments.Count

Dim i
For i = 0 To WScript.Arguments.Count - 1
    WScript.Echo "  [" & i & "] " & WScript.Arguments(i)
Next

If WScript.Arguments.Named.Exists("mode") Then
    WScript.Echo "mode: " & WScript.Arguments.Named("mode")
End If

WScript.Quit 3          ' 退出码 3，这一行之后的代码都不再执行
```

`WScript.Arguments` 从 0 开始，且**包含** `/mode:fast` 这类命名参数（与真实 WSH 一致），
命名参数同时可以通过 `Named` 按名字取，`/name:value` 与 `/name=value` 两种写法都认。
`Quit` 可以带退出码（默认 `0`），这是设置非零退出状态的唯一方式。`WScript.Echo` 用空格
连接各参数，和 `cscript` 完全一样。

---

## 读写文本文件

```vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.BuildPath(fso.GetSpecialFolder(2), "openvbs-demo.txt")   ' 2 = 临时目录

Set out = fso.CreateTextFile(path, True)      ' True = 覆盖已有文件
out.WriteLine "first"
out.WriteLine "second"
out.Close

Set src = fso.OpenTextFile(path, 1)           ' 1 = 只读
Do While Not src.AtEndOfStream
    line = src.ReadLine
    WScript.Echo "line: " & line
Loop
src.Close

WScript.Echo "exists: " & fso.FileExists(path)
WScript.Echo "size: " & fso.GetFile(path).Size
fso.DeleteFile path
```

只实现纯文本流（没有二进制流，也没有 `ADODB.Stream`）。`folder.Files` / `folder.SubFolders`
的 `.Count` 可用，`.Item(i)` 从 1 开始，`For Each` 也能遍历。文件不存在报 **53**、目录不存在
报 **76**，与 VBScript 一致。

---

## 把 CSV 汇总进 Dictionary

```vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set totals = CreateObject("Scripting.Dictionary")   ' 键大小写不敏感

Set data = fso.OpenTextFile("sales.csv", 1)         ' region,amount
Do While Not data.AtEndOfStream
    line = data.ReadLine
    parts = Split(line, ",")
    If UBound(parts) = 1 Then
        region = Trim(parts(0))
        amount = CDbl(Trim(parts(1)))
        If totals.Exists(region) Then
            totals(region) = totals(region) + amount
        Else
            totals.Add region, amount
        End If
    End If
Loop
data.Close

Dim key
For Each key In totals.Keys
    WScript.Echo key & ": " & FormatNumber(totals(key), 2)
Next
```

`Scripting.Dictionary` 是内置实现（不需要 COM），大小写不敏感，和真的一样。

---

## 错误处理：两种风格

`On Error` 是 VBScript 的老写法，照样能用：

```vbscript
On Error Resume Next
Set shell = CreateObject("No.Such.Class.Anyway")
If Err.Number <> 0 Then
    WScript.Echo "failed with " & Err.Number        ' Windows 上是 429
    Err.Clear
End If
On Error GoTo 0
```

`Try` / `Catch` / `Finally` 更简洁，处理器对象带 `Number` 与 `Description`：

```vbscript
Try
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile("nope.txt", 1)
Catch e When e.Number = 53
    WScript.Echo "missing file: " & e.Description
Catch e
    WScript.Echo "error " & e.Number & ": " & e.Description
Finally
    WScript.Echo "cleanup"
End Try
```

`Finally` 一定会执行，包括 `Try` 块里遇到裸 `End` 的时候。裸 `End` 立即结束整个脚本并以
退出码 0 结束，且无法被捕获——需要失败退出码时用 `Quit 1`。

---

## 用 COM 自动化 Windows **Windows**

```vbscript
' 环境变量、外部命令、WMI 查询
Set shell = CreateObject("WScript.Shell")
WScript.Echo "OS: " & shell.ExpandEnvironmentStrings("%OS%")

' 0 = 隐藏窗口，True = 等待，返回值是子进程退出码
code = shell.Run("cmd /c exit 5", 0, True)
WScript.Echo "child exit: " & code

Set wmi = GetObject("winmgmts:")
Set rows = wmi.ExecQuery("SELECT Caption, Version FROM Win32_OperatingSystem")
For Each row In rows
    WScript.Echo row.Caption & " (" & row.Version & ")"
Next
```

对象型属性可以像真 COM 那样链式取值（`folder.Files.Count`），`TypeName` 返回类型库里的名字
（`IWshShell3`、`FileSystemObject`、`SWbemServicesEx`），失败时报 VBScript 的错误号：
"无法创建" 是 `429`，"文件/类找不到" 是 `432`。`GetObject` 支持 moniker（`"winmgmts:"`）、
文件路径，以及 `GetObject(, "WScript.Shell")` 这种挂接已运行实例的写法。`WScript` 本身
始终可用，不依赖 COM。

---

## 和用户交互 **跨平台，Windows 上是真对话框**

```vbscript
answer = MsgBox("Save the report?", vbYesNoCancel + vbQuestion, "OpenVBS++")
WScript.Echo "picked: " & answer          ' 1 ok、2 cancel、6 yes、7 no

name = InputBox("Report name:", "OpenVBS++", "quarterly-report")
WScript.Echo "name: " & name
```

在满足启用条件的桌面环境里（从命令行启动、交互式终端、不在 CI 里、Tk 可用），内置的
`ui-support` 包会弹出真正的 Tk 对话框，并返回 VBScript 的返回码。`OPENVBS_UI=gui` 强制弹窗，
`OPENVBS_UI=console` 强制只用文本，`OVPP_UI_PATH` 可以指向自己的实现，`over add ui-support`
则把它装进项目里。

在网页 IDE 或任何没有终端的宿主里会退回控制台行为：`MsgBox` 打印文本并返回 `1`，`InputBox`
从 stdin 读一行。

---

## 把老脚本拆成模块

两种写法，区别很重要：

```vbscript
' 合并：对方文件里的过程直接成为本脚本的一部分
<script src="lib/format.vbs" />
WScript.Echo Money(12.5)

' 隔离：一切都挂在别名下面，不会泄漏出来
<namespace src="lib/strings.vbs" alias="s" />
WScript.Echo s.Shout("hello")
```

`<script src>` 适合自己拥有的代码，相当于文本级包含；`<namespace src alias>` 更像 import：
模块的全局变量留在命名空间里，因此两个模块可以有同名辅助函数。语言里没有 `import` 关键字。

---

## 让 Python 处理难的部分

```vbscript
<script src="python-embed" />          ' 显式打开 Python 桥

' Python 干活，结果通过 COP 池取回
#python
import json
from OpenVBSPP.ovpprt.cop import COP

payload = '{"b": 1, "a": 2}'
data = json.loads(payload)
COP["keys"] = ", ".join(sorted(data))
COP["total"] = sum(data.values())
#endpython

WScript.Echo "keys: " & COP("keys")     ' a, b
WScript.Echo "total: " & COP("total")   ' 3
```

`#python ... #endpython` 里是真正的 Python，运行在脚本的全局作用域；`COP` 是一份两边共享的
字典（Python 侧 `from OpenVBSPP.ovpprt.cop import COP`，VBS 侧 `COP("name")`），天然适合在两个
世界之间传值。JSON、正则、HTTP、二进制格式等 VBScript 表面覆盖不到的东西，都可以从这里走。
必须写 `<script src="python-embed">`，否则代码块不会生效。

---

## 用 `over` 打包可复用代码

项目根目录的 `over.json` 描述依赖：

```json
{
  "name": "reports",
  "version": "1.0.0",
  "entry": "main.vbs",
  "dependencies": {"format-utils": "file:../packages/format-utils"}
}
```

```powershell
openvbs over install            # 解析依赖、写 over.lock、安装 Python 依赖
openvbs over install --frozen   # CI：锁文件不满足就失败
openvbs over list               # over 包、内置包、Python 包
openvbs over add ui-support     # 内置包也能按名字加进项目
```

临时试验时，把目录放到 `packages/<name>/` 就行——项目里的包会覆盖内置的同名包，不需要安装
步骤。`over pack` / `over publish` 生成 `.ovppkg` 归档；完整参考见 [OVER.zh-CN.md](OVER.zh-CN.md)。

---

## 打开类型检查与原生算术

```vbscript
Option Explicit        ' 每个变量都要 Dim（VBScript 行为）
Option Typing          ' 静态检查：声明和参数可以带 -> 类型标注
Option Native          ' + - * / 不再转换文本："1" + "2" 抛 13

Dim count -> Integer
count = 2

Function Twice(n -> Integer) -> Integer
    Twice = n * 2
End Function

WScript.Echo Twice(count)
```

`Option Native` 会改变语义，所以是选择性开启；`Option Typing` 隐含了它。新写的代码建议打开，
让它大声报错而不是悄悄做类型转换。

---

## 迁移老脚本

1. **先直接跑。** `openvbs old.vbs` 通常就够了——VBScript 核心（控制流、类、`Property`、
   `With`、`On Error`、`ReDim`、`Erase`、`Select Case`、`#date#` 字面量、`&H`/`&O`、`GetRef`、
   `ExecuteGlobal`、95+ 内置函数、`WScript`）都已实现。
2. **只有 8.3 短名或 UNC 路径才需要改。** `C:\data\x.txt` 原样可用；路径函数在 Linux 上也保持
   Windows 语义。
3. **不支持 `Declare`**（有意为之）。需要原生调用时改用 COM、`#python` 块或 `over` 包。
4. **`LoadPicture` 会直接抛错**，没有图形对象类型；图片请留在脚本之外。
5. **`MsgBox`/`InputBox` 会退化成控制台调用**，除非满足弹窗的启用条件。
6. **向后 `GoTo` 会被拒绝**，请改用 `Do ... Loop` 重写（向前 `GoTo` 可用）。
7. **选项必须写在最前面**——`Option` 语句不能出现在脚本中间。
8. **退出码**：支持 `WScript.Quit n` 与裸 `End`（退出码 0）；`Err.Raise` 的错误号会以
   `OVPPRuntimeError.number` 的形式浮出来。

CI 里可以用 `openvbs check old.vbs --json` 自动跑一遍，它给出带行列号的诊断。

---

## 接下来看什么

- [VBS_SUPERSET.zh-CN.md](VBS_SUPERSET.zh-CN.md)——语言参考与已知缺口表。
- [OBJECT_MODEL.zh-CN.md](OBJECT_MODEL.zh-CN.md) / [OOP.zh-CN.md](OOP.zh-CN.md)——类、`Property`、`Inherits`。
- [MODERN_FEATURES.zh-CN.md](MODERN_FEATURES.zh-CN.md)——`Option Typing`、`Try/Catch`、`#python`、语法糖。
- [OVER.zh-CN.md](OVER.zh-CN.md) 与 [WEB_IDE.zh-CN.md](WEB_IDE.zh-CN.md)——打包与网页 IDE。
- 仓库里的 `examples/`——同样的配方，可直接运行；`examples/demo/` 是端到端演示。
