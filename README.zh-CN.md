# OpenVBS++

![tests](https://img.shields.io/badge/tests-577%20passing-brightgreen)
![version](https://img.shields.io/badge/version-0.9.21-blue)
![license](https://img.shields.io/badge/license-Apache--2.0-blue)
![python](https://img.shields.io/badge/python-3.10%2B-yellow)
![deps](https://img.shields.io/badge/dependencies-none-success)
[![CI](https://github.com/MarfChen/OpenVBSpp/actions/workflows/ci.yml/badge.svg)](https://github.com/MarfChen/OpenVBSpp/actions/workflows/ci.yml)

> **English** — the primary document is the English [README.md](README.md);
> this file **README.zh-CN.md** is the full Chinese version of it.
>
> **中文** —— OpenVBS++ 把 VBScript 风格源码编译成标准 Python AST 后在 CPython 上运行。
> 它**必须是 VBScript 的超集**：任何合法 VBS 程序都要继续能跑，扩展只能新增语法。
> 主文档是英文版 [README.md](README.md)，本文件是完整中文版。
> 第一次来？看 [CONTRIBUTING.md](CONTRIBUTING.md)，挑一个
> [新手任务](dev-docs/good-first-issues.md)。
>
> 文档站：<https://marfchen.github.io/OpenVBSpp/> · 变更记录：[CHANGELOG.md](CHANGELOG.md) ·
> 法务自查：[doc/LEGAL.md](doc/LEGAL.md)

>
> **免责声明 / Disclaimer** —— OpenVBS++ 是独立实现，与 Microsoft 无任何关联；
> VBScript 是 Microsoft Corporation 的商标。本仓库不包含 Microsoft 的任何代码或文档。
> OpenVBS++ is an independent reimplementation, not affiliated with or endorsed by
> Microsoft. VBScript is a trademark of Microsoft Corporation.
> 详细的自查记录与合规清单见 [doc/LEGAL.md](doc/LEGAL.md)。

**OpenVBS++**（Python 包名 `OpenVBSPP`，当前版本 **0.9.22**）是一个以 **VBScript 语法风格**为基础的编译器。
它不解释源码，而是把 VBS 风格的脚本编译成**标准 Python AST**，再由 CPython 执行：

```
.vbs 源码
  → lexer.py        正则词法分析（注释 / CRLF / 行续接 / : / 行列号）
  → parser.py       递归下降语法分析 → OpenVBS++ 自己的 AST
  → transformer.py  AST → 标准库 ast（行号一并带过去）
  → compiler.py     compile() → code object → exec
```

**项目铁律：OpenVBS++ 必须是 VBScript 的超集。** 任何合法 VBS 程序都要能跑，
扩展只新增语法、不改变已有语义。状态矩阵见 [doc/VBS_SUPERSET.md](doc/VBS_SUPERSET.md)。
语言规范见 [doc/OpenVBS spec.docx](doc/OpenVBS%20spec.docx)（随 PyPI 分发包一起发出）；
设计初衷见 [dev-docs/plan.docx](dev-docs/plan.docx)（只在仓库里，不随包分发）。

## 快速开始

```bash
# 仓库自带虚拟环境（Python 3.10+；0.9.14 是最后一个声明支持 3.9 的版本）
.\venv\Scripts\python.exe cli.py examples\procedures.vbs

# 不带参数跑内置示例
.\venv\Scripts\python.exe cli.py
```

```vb
Option Explicit

Sub PrintLine(text)
    echo "-- " & text
End Sub

Function Fact(n)
    If n <= 1 Then Fact = 1 Else Fact = n * Fact(n - 1)
End Function

Dim x, y
x = 1 : y = 2                 ' 冒号分隔多条语句
PrintLine "factorial"
echo Fact(5)                  ' 120
```

在 Python 里调用：

```python
from OpenVBSPP.runtime import vbs_run

vbs_run('''
Function Add(a, b)
    Add = a + b
End Function
echo Add(2, 3)
''', filename="demo.vbs")
```

只想要 code object：

```python
from OpenVBSPP.compiler import vbs_compile
code_obj = vbs_compile("echo 1\n", filename="demo.vbs")
exec(code_obj, {})
```

## 命令行

```
python cli.py [script.vbs] [-v] [--save] [--save-dir DIR]
python cli.py -c "echo 1 + 1"            # 直接运行一段源码，不用建文件
echo "echo 6 * 7" | python cli.py -c -   # 从标准输入读源码
```

| 参数 | 说明 |
|---|---|
| `script.vbs` | 要运行的文件；省略则运行内置示例 |
| `-c, --code CODE` | 直接运行 CODE（`-c -` 从标准输入读）；与脚本文件互斥 |
| `-` | `-c -` 的简写，从标准输入读源码 |
| `-v, --verbose` | 打印 token 流、AST 与生成的 Python 代码 |
| `--save` | 额外写出 `.pyc` 与 `.ast.txt` |
| `--save-dir` | 落盘目录，默认 `./__pycache__`（自动创建） |
| `check <script>` | 只编译不运行，报诊断；`--json` 输出机器可读结果 |
| `run <script>` | 运行；`--json` 输出 `{ok, output, error, duration_ms}` |
| `over <子命令>` | 包管理器：`install` / `list` / `info <name>` / `remove <name> [--python]` / `com` …（见 [doc/OVER.md](doc/OVER.md)） |
| `ide [--port N]` | 启动网页 IDE（默认 `http://127.0.0.1:7860`） |

安装成包之后可以直接用 `openvbs script.vbs`：

```bash
pip install -e .            # 本地开发安装
pip install openvbspp       # 从 PyPI 安装（发布后）
openvbs examples/fizzbuzz.vbs
openvbs --version
python -m OpenVBSPP examples/control1.vbs   # 等价写法
```

## 网页 IDE / Web IDE

```bash
openvbs ide                 # http://127.0.0.1:7860
openvbs ide --port 8000 --workspace ./scripts
```

浏览器里写代码、Ctrl+Enter 运行、看输出与诊断、保存/读取脚本。零第三方依赖，
包管理器已经落地（`over`，见 [doc/OVER.md](doc/OVER.md)），网页 IDE 里的包管理接口仍是预留。

## 编辑器 / IDE 支持

| 编辑器 / 环境 | 现状 | 怎么用 |
|---|---|---|
| 内置网页 IDE | ✅ 完整 | `openvbs ide`：编辑 / 运行 / 检查 / 保存，零依赖 |
| 任意编辑器、任务、CI | ✅ 完整 | `openvbs check --json`、`openvbs run --json` 提供机器可读结果 |
| VS Code | ⚠️ 只有语法与片段 | [extras/vscode-openvbspp](extras/vscode-openvbspp)：TextMate 语法、注释/括号配置、代码片段；没有智能补全 |
| Neovim / Emacs / Sublime / JetBrains | ⚠️ VBScript 高亮 + CLI 诊断 | 把编辑器接到 `openvbs check --json`；语言服务器尚未提供 |
| 语言服务器（LSP） | ❌ 未实现 | 设计与里程碑见 [dev-docs/IDE_LSP_DESIGN.md](dev-docs/IDE_LSP_DESIGN.md) |
| 调试器（DAP） | ❌ 未实现 | `Err.Line` 与 Try/Catch 栈帧是前置基础 |

所有集成共用的机器可读契约：

```bash
openvbs check --json script.vbs   # {"ok": true|false, "diagnostics": [{line, column, message, source_line, ...}]}
openvbs run --json script.vbs     # {"ok": true|false, "output": "...", "error": "...", "duration_ms": 12.3}
```

IDE 服务还导出带版本号的 capabilities 文档（[OpenVBSPP/ide/service.py](OpenVBSPP/ide/service.py)），
未来的 LSP 服务器会复用同一份声明；完整的集成方案见 [dev-docs/IDE_LSP_DESIGN.md](dev-docs/IDE_LSP_DESIGN.md)。

## 安装与打包

| 目标 | 命令 | 产物 |
|---|---|---|
| 本地开发 | `pip install -e ".[all]"` | 可编辑安装 + 全部开发依赖 |
| PyPI 包 | `.\tools\build.ps1 -Target wheel` | `dist/*.whl` + `dist/*.tar.gz` |
| Windows exe | `.\tools\build.ps1 -Target exe` | `dist\openvbs.exe`（单文件，无需 Python） |
| Windows COM 支持 | `pip install "openvbspp[com]"` | `pywin32`（也可以让 `over install` 自动装） |

发布流程（TestPyPI 验证 → PyPI → 打 exe）见 [dev-docs/RELEASE.md](dev-docs/RELEASE.md)。
项目已预留好：`[project.scripts]` 入口、`MANIFEST.in`、`packaging/openvbs.spec`
（PyInstaller 配置）、`tools/build.ps1`。

**支持的 Python：3.10 – 3.13（CPython）。** 本项目是纯 Python，发布的是
`py3-none-any` 通用 wheel，没有需要编译的东西，也不需要按版本分别打包：
`pip install openvbspp` 在 x64 与 ARM64 上通用。577 个测试在每个受支持的版本上都跑 CI；
运行期代码本身仍然能被 3.9 解析并运行（`vermin` 给出的下限就是 3.9），
但**声明支持 3.9 的最后一个版本是 0.9.14**。版本策略与 ARM 相关要做的事见
[dev-docs/RELEASE.md](dev-docs/RELEASE.md) §9。

## 已支持的语法

| 类别 | 支持情况 |
|---|---|
| 语句分隔 | 换行、`:`、行尾 `_` 续行；`rem` / `'` 注释；CRLF / CR / LF / BOM |
| 标识符 | 大小写不敏感 |
| 变量声明 | `Dim a, b(10), c()`；`ReDim [Preserve]`；`Const`；`Option Explicit`；`Let` / `Set` |
| 数据类型 | 整数、浮点、字符串、布尔、`Empty` / `Null` / `Nothing`、一维 / 多维数组 |
| 运算符 | `+ - * / \ Mod ^`、`&`、`= <> < > <= >=`（也接受 `==`）、`Is`、`And Or Not Xor Eqv Imp` |
| 条件 | `If ... Then ... End If`、**单行 `If ... Then ... [Else ...]`**、`ElseIf`、`Select Case`（含 `Is`、`To`、逗号列表） |
| 循环 | `For ... Next`、`For Each ... Next`、`Do [While\|Until] ... Loop [While\|Until]`、`While ... Wend`、`Exit Do/For` |
| 过程 | `Sub` / `Function`、`ByRef`（默认）/ `ByVal`、递归、`Exit Sub/Function`、`Call` |
| 包管理 | over：本地目录 / .ovppkg / git+ 源，锁文件，冲突路径报告；声明的 Python 依赖自动交给 pip 装 |
| 多文件 | <script src="lib.vbs"> 合并、<namespace src="lib.vbs" alias="u"> 命名空间 |
| 一等函数 | `f = AddressOf Foo`，过程可以存进变量、当参数传递 |
| 正则 / JSON | `New RegExp`（`Test`/`Replace`/`Execute`）、`Json.Parse` / `Json.Stringify` |
| 语法糖 | 字符串插值 $"..."、可选链 ?.、空合并 ??、复合赋值 += &= |
| 结构化异常 | Try / Catch e [When 条件] / Finally / End Try，与 On Error 共存 |
| 错误处理 | `On Error Resume Next`、`On Error GoTo 0`、`Err.Number/.Description/.Source`、`Err.Clear`、`Err.Raise` |
| 对象语法 | `Class ... End Class`、字段、`Sub`/`Function` 方法、`Property Get/Let/Set`、`Me`、`Class_Initialize`/`Class_Terminate`、`New`、`CreateObject`、`Dictionary`、`With ... End With`、`.成员` / `.方法()` |
| 类型系统 | `Option Typing` + `->` 注解（关闭时完全等同普通 VBS） |
| 继承 | `Class Child Inherits Parent`（超集扩展：方法覆盖、字段/属性继承） |
| 网页 IDE | openvbs ide：浏览器里编辑/运行/检查/保存，零第三方依赖 |
| Python 互操作 | `#python ... #endpython` 代码块、`COP` 公共对象池；用 `<script src="python-embed" />` 显式导入后可用 |
| 动态执行 | `Execute` / `Eval`（复用编译器，在全局作用域运行） |
| 其它 | `Erase`、`Option Compare Text`、`FormatNumber` 系列、`vbCrLf` 等预定义常量、约 95 个内置函数 |

## 过程：Sub / Function

```vb
' 无参数 / 有参数
Sub Hello()
    echo "hi"
End Sub

Sub Show(a, b)
    echo a & "-" & b
End Sub

' 三种调用写法都支持
Hello
Call Show("x", "y")
Show "x", "y"

' 函数用"函数名"承载返回值
Function Add(a, b)
    Add = a + b
End Function
```

- **过程可以先用后定义**：编译器先扫描整个脚本登记签名。
- **函数返回值**：给函数名赋值；一次都没赋值则返回 `Empty`。
- **`Exit Sub` / `Exit Function`**：提前返回。
- **参数默认 ByRef**（VBS 约定），会写回调用方；`ByVal` 不写回。

```vb
Sub Swap(ByRef a, ByRef b)
    Dim t
    t = a : a = b : b = t
End Sub

Function Bump(ByRef a)
    a = a + 1
    Bump = a * 10
End Function

x = 1 : y = 2
Swap x, y            ' x=2, y=1
x = 1
y = Bump(x)          ' x=2（ByRef 写回），y=20
x = Bump(x)          ' 先写回 2，再用返回值 20 覆盖 x —— 与 VBS 一致
```

实现方式：实参包进 `Ref` 盒子（`ovpprt/ref.py`），过程内读写成 `参数.v`，调用后写回；
调用出现在表达式或 `If` 条件里时，编译器把"建盒子 / 写回"的语句提升到正确位置。

## 语法糖：`:`、单行 If、行续接

```vb
a = 1 : b = 2 : echo a + b          ' 冒号分隔

If b = 2 Then echo "two"             ' 单行 If
If b = 2 Then echo "two" Else echo "other"
If b = 2 Then echo "a" : echo "b" Else echo "c"

total = 1 + _
        2                            ' 行尾下划线续行
```

## 变量、数组与 Option Explicit

```vb
Option Explicit            ' 开启后所有变量必须先声明

Dim total, i               ' 标量，初值是 Empty
Dim scores(4)              ' 数组：下界 0、上界 4，共 5 个元素
Dim grid(1, 2)             ' 二维数组，访问 grid(1, 2)

For i = 0 To UBound(scores)
    scores(i) = (i + 1) * 10
Next

ReDim Preserve scores(5)   ' 扩容并保留已有值
Erase scores               ' 元素重置为 Empty
Const NAME = "openvbs"
```

- `Empty` 与 VBS 一致：打印为空行；参与算术当 0；参与 `&` 当 `""`。
- `Null` 会"传染"：`Null + 1`、`Null = 0` 都是 `Null`，用 `IsNull()` 判断。
- `Dim a()` 是动态数组，之后用 `ReDim a(n)` 定尺寸。
- `Option Explicit` 在**编译期**检查未声明变量，报错带行号与源码行。

## 类与对象

```vb
Class Counter
    Private n
    Public Property Get Value()
        Value = n
    End Property
    Public Sub Bump(ByVal step)
        n = n + step
    End Sub
    Private Sub Class_Initialize()
        n = 0
    End Sub
End Class

Set c = New Counter
c.Bump 5
echo c.Value            ' 5
```

- 支持字段（`Public` / `Private` / `Dim`）、方法（`Sub` / `Function`）、
  `Property Get/Let/Set`、`Me`、`Class_Initialize` / `Class_Terminate`；
- `New Foo` 与 `CreateObject("Foo")` 都能创建脚本类；
- `CreateObject("Scripting.Dictionary")` 提供 `Add/Exists/Item/Remove/RemoveAll/Count/Keys/Items`，
  并且 `d("k")`、`d("k") = v`、`For Each k In d` 都能直接用；
- Windows 上 `CreateObject("WScript.Shell")` / `Excel.Application` 这些 COM 对象也能真连：
  属性、方法、集合、`For Each` 都可用（由 `com-support` 包提供，无需 `import`；
  见 [extras/com-support](extras/com-support)）；
- `Set x = Nothing`、`x Is Nothing`、`TypeName(obj)`、`IsObject(obj)` 语义与 VBS 一致；
- `With obj ... End With` 的表达式只求值一次。

细节与限制见 [doc/OBJECT_MODEL.md](doc/OBJECT_MODEL.md)（设计报告）和
[doc/VBS_SUPERSET.md](doc/VBS_SUPERSET.md) §3.7（方法参数的 ByRef 差异）。

## Option Typing：编译期类型检查

默认关闭 —— 关闭时行为和普通 VBS 完全一致，类型注解会被忽略。
打开后，注解（`->`，不是 `:`，避免与语句分隔符冲突）会参与编译期检查：

```vb
Option Typing

Dim count -> Integer
Dim ratio -> Double
Dim name -> String

count = 10
name = "openvbs"

Function Twice(n -> Integer) -> Integer
    Twice = n * 2
End Function

echo Twice(count)
```

检查范围：赋值、实参、返回值，以及能静态确定接收者类型时的类字段 / 属性。
`Variant`、`Empty`、`Null`、`Object` 通配一切，所以不会误伤没写类型的代码：

```
Type mismatch in assignment to 'count': expected Integer, got String
Location: Line 12
Source: count = "oops"
```

## With

```vb
With Err
    .Number = 3
    .Clear()
End With
echo Err.Number
```

## 错误处理

```vb
On Error Resume Next
value = 10 / 0
If Err.Number <> 0 Then
    echo "failed: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
```

- `On Error Resume Next` 之后的语句被逐条保护，出错时记录到 `Err` 并继续执行；
- 每个过程一个错误处理帧，子过程里的 `On Error` 不会泄漏给调用方；
- `Err.Raise number, source, description` 可以由脚本主动抛错；
- 没有开启保护时，错误照常向上抛，并被翻译成带行号的 `OVPP*` 错误。

## 内置函数

| 分类 | 函数 |
|---|---|
| 字符串 | `Len Left Right Mid UCase LCase LTrim RTrim Trim StrReverse Space String Asc Chr InStr InStrRev Replace Split Join Filter StrComp Hex Oct` |
| 转换 | `CStr CInt CLng CDbl CBool CByte CCur CSng CDate` |
| 数学 | `Abs Int Fix Sgn Sqr Round Rnd Randomize Sin Cos Tan Atn Exp Log` |
| 数组 | `Array UBound LBound` |
| 判断 | `IsNumeric IsEmpty IsNull IsObject IsArray IsDate IsMissing TypeName VarType` |
| 日期 | `Now Date Time Timer Year Month Day Hour Minute Second Weekday DateAdd DateDiff DatePart DateSerial DateValue TimeSerial TimeValue MonthName WeekdayName` |
| 格式化 | `Format FormatNumber FormatCurrency FormatDateTime FormatPercent` |
| 运行时信息 | `RGB ScriptEngine ScriptEngineMajorVersion ScriptEngineMinorVersion ScriptEngineBuildVersion GetLocale SetLocale` |
| 过程 | `Execute` / `ExecuteGlobal` / `Eval` / `GetRef` / `LoadPicture`（没有图形库，会报错） |
| 交互 | `MsgBox` / `InputBox`（命令行交互时弹真对话框，否则控制台替身 —— 见[已知限制](#与规范的差距--已知限制)） |

预定义常量：`vbCrLf vbCr vbLf vbTab vbNullString vbNullChar vbBack vbFormFeed vbVerticalTab vbTrue vbFalse`。

## 数字与布尔语义

- 比较运算返回 **VBS 布尔**：真 = `-1`、假 = `0`，`Echo` 打印成 `True` / `False`。
- **逻辑运算符是按位的**（与 VBS 一致）：`5 And 3` → `1`、`Not 0` → `-1`、
  `(1 = 1) And (2 = 2)` → `True`（两边都是布尔时保持布尔）。
- `Echo` 与 `&` 都按 VBS 习惯格式化数字：`3.0` 打印成 `3`，`7 / 2` 打印成 `3.5`。
- `\` 与 `Mod` 向零截断：`-7 Mod 2` → `-1`；`^` 左结合：`2 ^ 3 ^ 2` → `64`。

## 错误信息

语法错误与运行期错误都会带**行号 + 源码行**，错误类型见 [OpenVBSPP/exceptions.py](OpenVBSPP/exceptions.py)：

```
Syntax error (1002): Missing 'End If': unexpected end of file while reading the If ... End If block
Location: Line 2, Column 1
Source: echo 1
```

```
Variable is undefined (1001): Variable 'zzz' is not defined
Description: Python NameError: name 'zzz' is not defined
Location: Line 3
Source: echo zzz
```

其它常用错误：`Sub or Function not defined (35)`、`Type mismatch (13)`、
`Wrong number of arguments (450)`、`Subscript out of range (9)`、`Division by zero (11)`。

## 性能

先给结论：**作为「老脚本的迁移执行层」，它比原生 `cscript.exe` 快 2.5–3 倍，启动也快约 1.5 倍；
循环、数组、函数调用的运行时开销已经压到接近手写 Python（1.0–1.4 倍），
剩下明显偏慢的只有 `Dictionary` 的成员访问和深递归。** 适合批处理、胶水、运维自动化，
不适合把百万次字典插入或深递归塞进 VBS 里做核心计算。

自己复现（不需要任何参数，读取仓库内源码，不含硬编码路径）：

```bash
python tools/bench.py          # 进程内：编译一次、反复执行
python tools/bench.py --e2e    # 额外跑端到端（含解释器启动），Windows 上自动对照 cscript.exe
```

### 端到端（含进程启动，这才是用户感受到的数字）

| 脚本 | OpenVBS++ | cscript.exe | CPython |
|---|---:|---:|---:|
| 只 echo 一行 | **114 ms** | 169 ms | 70 ms |
| 1M 次算术循环 | **142 ms** | 438 ms | 101 ms |
| fib(24) | **150 ms** | 394 ms | 76 ms |

### 进程内（编译一次、反复执行，纯运行开销；中位数）

| 用例 | OpenVBS++ | 等价 CPython | 倍数 |
|---|---:|---:|---:|
| 1M 算术循环（过程内） | 29.7 ms | 30.1 ms | **0.99×** |
| 1M 算术循环（模块级） | 69.7 ms | 67.0 ms | **1.04×** |
| 100k 数组读写 | 12.8 ms | 12.7 ms | **1.00×** |
| 200k 函数调用 | 23.1 ms | 17.1 ms | 1.35× |
| 20k 字符串拼接 | 6.3 ms | 4.5 ms | 1.39× |
| fib(24) 递归 | 37.9 ms | 5.5 ms | 6.85× |
| 50k Dictionary 插入 | 58.4 ms | 4.5 ms | 13.1× |

### 这些数字是怎么测出来的

- **机器**：13th Gen Intel Core i5-13500H、15.6 GB 内存、Windows 11 家庭中文版、
  仓库 `venv` 里的 CPython 3.13.15，`cscript.exe` 来自同一套 Windows。
  所有数字都出自这一台机器；你机器上的绝对耗时必然不同，值得信的是**倍数**。
- **工具**：`tools/bench.py`，无硬编码路径，每个用例都配一份等价的、手写风格的
  Python 源码。`--repeat N` 采样 N 次（默认 5）取**中位数**，避免杀毒扫描或 GC
  偶尔卡一下就把表格带偏。
- **进程内**：脚本只编译一次，然后在同一进程里反复执行生成的代码，用于隔离
  「运行期开销」与「启动开销」。CPython 列是按 Python 习惯写的等价实现
  （原生 `list`/`dict`、不加兼容层），所以它是**下限而非法力对等**——OpenVBS++
  还要维持 VBScript 语义。
- **端到端**：每次采样都新起进程，用墙钟时间包住
  `python -m OpenVBSPP <file>`、`cscript //nologo <file>`、`python <file>`，
  这才是定时任务里等结果的用户真正感受到的东西。
- **抖动**：本机上多次运行之间大约 ±5%。对外引数字前先重跑
  `python tools/bench.py`，并写清测的是哪个 commit。
- **没有测**：内存占用、多线程、超大文件解析，以及非 Windows 平台。不要把
  `cscript` 那两列当成对其它 VBScript 实现的结论。

### v0.9.15 做的四项优化

这四项的共同前提是「只改生成的代码，不改语言语义」，每一项都由
`tests/test_perf_semantics.py` 钉住边界：

| 优化 | 做法 | 效果 |
|---|---|---|
| 启动懒加载 | `regexp`/`jsonobj`/`object`/`builtins` 里的 `re`、`json`、`datetime`、`math`、`random` 改成首次使用时才导入；`cli.py` 只在真正执行脚本时才导入 `runtime` | 冷启动开销 27 ms → 13 ms，`openvbs --version` 82 ms |
| 只读 `ByRef` 参数不再装箱 | 编译器静态判断形参是否真的被写（赋值、`ReDim`、`For`、或继续传给下层），只读的按值传，被调方内联拆包 `if type(n) is Ref: n = n.v` | 200k 次函数调用 118 ms → 23 ms（13.7× → 1.35×） |
| 无 `On Error` 的过程不再建错误帧 | 过程体里没有 `On Error ...`/`Resume`/`Err` 就不生成 `enter()/try/finally/leave()`，模块加载时已有一个基帧兜底 | 递归 fib(24) 67 ms → 38 ms（12.0× → 6.85×） |
| `For` 安全时编译成 `range` | 步长是常量、循环体不给计数器赋值时，生成 `for i in for_seq(...)` + `else: i = i + step`，复现「正常结束后计数器 = 末值 + 步长、`Exit For` 不加、空区间保留起点」 | 1M 循环 50 → 29.7 ms（1.65× → 0.99×），模块级 144 → 69.7 ms（4.19× → 1.04×） |

早先还做过两项：`Dictionary` 成员分派从「线性扫全部键」改成「方法名直接判定」
（O(n²) → O(1)，5 万次插入从 103 秒降到 58 ms）；`&` 拼接在两端都是字符串时走原生 `+` 快路径。

### 还能再快、但收益与代价要再权衡的项

- `Dictionary` 每次成员访问仍要过一遍运行时函数调用（13×），要彻底解决得让编译器识别 `d.Add` 这种模式。
- 递归慢在「每次调用都是一次 Python 函数调用 + 参数绑定」，要更快需要内联或尾调用优化，会牵动 `On Error` 语义。
- `Dim a(n)` 仍是运行时数组对象，改成原生 `list` 能再省一点，但会破坏 `LBound/UBound` 的可扩展性。

## 为什么不用 X？

这里给诚实的答案 —— 它不是所有场景的最优解。

**为什么不用原生 VBScript / `cscript.exe`？**
如果脚本已经稳定、不需要更快的速度、环境里也不允许安装任何东西，就用 `cscript`，
它是系统自带、也是参考实现。需要类 / try-catch / 包管理、需要批处理快 2–3 倍、
需要可机读的诊断与 JSON 输出时，才值得换 OpenVBS++。代价是多一个依赖
（Python 或 8 MB 的 exe），而且它不是厂商官方运行时。

**为什么不用 Python？**
写新代码时，Python 通常是更好的选择：生态更大、跨平台、工具链成熟。
OpenVBS++ 赢在**必须保留 `.vbs` 源码和入口**的场合（计划任务、登录脚本、客户现场、
老构建系统），或者维护脚本的人只懂 VBScript。它是迁移与执行层，不是 Python 的替代品。

**为什么不用 PowerShell？**
PowerShell 是更现代的 Windows 管理外壳，模块生态大得多。但很多环境对 PowerShell 有
执行策略限制、脚本要签名，企业里还堆着大量能跑的 `.vbs` —— 重写是纯风险、几乎没有收益。
如果你在给可控机器写**新**自动化，PowerShell 可能更合适；OpenVBS++ 的定位是让**已有的**
脚本继续跑下去，而且跑得更快、更好调试。

**为什么不用 JavaScript / Node？**
为了一件只需要文件、注册表、COM 访问的事情去引入运行时和 npm 供应链并不划算，
TS/JS 在这方面也给不了什么。如果你同时要做 Web UI、或者要用那个生态的库，Node 才是对的。

**为什么不用 C# / .NET？**
适合长期维护、强类型、编译型的应用和真正的服务。对胶水脚本来说太重：项目文件、构建链、
运行时体积。当脚本开始变成产品时，再上 .NET。

**为什么不用 AutoHotkey / AutoIt？**
它们强在 UI 自动化、热键、窗口控制。OpenVBS++ 面向的是文件、注册表、WMI、COM 自动化，
并且沿用 VBScript 的语法与语义。

**什么情况下完全不该用 OpenVBS++**
需要跨平台、需要 CPU 密集的重计算、需要长期运行的服务，或者环境禁止安装任何东西、
必须用厂商运行时。另外 VBScript 的语义边界仍有极少数差异，见
[与规范的差距 / 已知限制](#与规范的差距--已知限制)。

## 项目结构

```
OpenVBSPP/
├── lexer.py            词法分析（注释 / CRLF / 行续接 / : / 行列号）
├── parser.py           递归下降语法分析（过程、声明、数组、调用、With、On Error）
├── transformer.py      OpenVBS++ AST → Python ast（两遍：先登记签名再生成）
├── compiler.py         编译入口 vbs_compile()，支持 exec / save
├── runtime.py          vbs_run()：编译 + 执行 + 异常翻译
├── cli.py              命令行入口（安装后为 openvbs 命令）
├── exceptions.py       OVPP* 错误类（沿用 VBS 错误号）
├── ast_nodes/          AST 节点定义（表达式 / 语句 / 控制流 / 过程）
├── ovpprt/             运行时库
│   ├── boolean.py      VBool + 按位逻辑运算
│   ├── string.py       & 字符串连接
│   ├── format.py       VBS 风格的数字/空值格式化
│   ├── empty.py        Empty 哨兵
│   ├── null.py         Null 哨兵（传染语义）
│   ├── nothing.py      Nothing 哨兵
│   ├── array.py        数组创建 / ReDim / Erase / UBound
│   ├── ref.py          ByRef 盒子
│   ├── err.py          Err 对象 + On Error 状态机
│   ├── object.py       对象运行时（CreateObject / 成员分派）
│   ├── com.py          COM 桥（定位 com-support 包并挂进 CreateObject）
│   ├── collections.py  Scripting.Dictionary
│   └── builtins.py     内置函数 + Echo
├── typecheck.py        Option Typing 的编译期类型检查
├── over.py             over 包管理器（依赖图 / 锁文件 / Python 依赖 / bundled 包）
├── plugins/            预处理与词法/语法辅助模块
└── vendor/             随包分发的 over 包：com-support（COM 支持）、python-embed（Python 互操作）
examples/               示例脚本（控制流 / 过程 / 数组 / 语法糖 / 错误处理 / 类 / 类型 / FizzBuzz）
examples/demo/          录屏用六段演示项目（旧脚本 / 类与 Try / Python / COM / 命名空间 / 包管理器）
tests/                  回归测试（577 个用例）
doc/                    语言规范、VBS 超集矩阵、对象模型设计
```

## 测试

```bash
.\venv\Scripts\python.exe -m unittest discover -s tests -v
```

- `tests/test_openvbspp.py`：示例输出、运算符、布尔语义、控制流、词法、错误定位、编译器与命令行。
- `tests/test_procedures.py`：Sub / Function / ByRef / 递归 / Dim / 数组 / Option Explicit / 内置函数 / Empty。
- `tests/test_objects.py`：Class / Property / Me / Class_Initialize / Dictionary / 私有成员 / 默认成员 / 对象错误。
- `tests/test_option_typing.py`：`Option Typing` 的类型检查（含数组元素类型、继承链）。
- `tests/test_interop.py`：`#python` 代码块、`COP` 对象池与显式导入要求。
- `tests/test_vbs_superset.py`：**VBS 超集一致性**——100+ 条合法 VBS 片段必须跑通；
  v0.5 起已没有 expectedFailure（缺口全部补齐），新增缺口时按老规矩加。
- `tests/test_vbs_coverage.py`：v0.9.21 补齐的每个 VBScript 内置与字面量（`&H`/`&O`、`#date#`、数学、`VarType`、`GetRef`、`ExecuteGlobal` 等）。
- `tests/test_over.py` / `tests/test_over_registry.py`：over 的依赖图、锁文件、本地源与 registry。
- `tests/test_over_python.py`：over 声明 Python 依赖时的合并、钉版本、frozen 还原与卸载。
- `tests/test_com.py`：COM 桥的定位优先级、开关、错误翻译与适配层行为（不依赖真实 COM）。
- `tests/test_ui_dialogs.py`：`ui-support` 的按钮/图标/默认按钮映射、启用规则与控制台回退。

## 与规范的差距 / 已知限制

**尚未实现**（`doc/VBS_SUPERSET.md` §4）：

- `Collection` 内置类（VBScript 本身没有，属于 VBA）
- `Declare`（声明外部 DLL 函数）—— 短期有意不支持：原生调用写错会让进程无法捕获地退出，见 [doc/VBS_SUPERSET.zh-CN.md §3.11](doc/VBS_SUPERSET.zh-CN.md)
- `LoadPicture` 直接报错而不加载图片
- `ExecuteGlobal` 在运行期定义的过程无法按名字调用（改用 `GetRef` 绑定），且只支持向前 `GoTo`

`MsgBox` 与 `InputBox` **已经支持**：自 v0.9.22 起，内置的 `ui-support` 包在脚本由
`openvbs` 命令行交互式启动时用 `tkinter` 弹出真对话框（按钮组合、图标、默认按钮与
返回码都按 VBScript 的语义）。在 CI、管道、Python 里嵌入 `vbs_run()` 以及 Tk 不可用
的场合，它们退回控制台行为（`MsgBox` 打印并返回 1，`InputBox` 读一行）。详见
[doc/VBS_SUPERSET.zh-CN.md §3.10](doc/VBS_SUPERSET.zh-CN.md) 与
[dev-docs/GUI_DIALOGS_DESIGN.zh-CN.md](dev-docs/GUI_DIALOGS_DESIGN.zh-CN.md)。

COM 宿主对象（`Excel.Application`、`WScript.Shell` 等）**已经支持**：Windows 上由
`com-support` 包真连，非 Windows 平台、`OPENVBS_COM=0`、或没装 `pywin32` 时会给出明确报错
（不再按"未实现"处理）。失败编号与 VBScript 一致：找不到 ProgID 报 429，连接失败报 432，
`TypeName` 返回组件类名；`GetObject` 支持 moniker（`winmgmts:`）与 `GetObject(, "类名")` 形式。
`WScript` 宿主对象本身没有提供 —— `WScript.Echo` 等会报 `Variable 'WScript' is undefined`（1001），
输出请用 `echo`，shell 功能用 `CreateObject("WScript.Shell")`。

Tcl/Tk 的许可与义务见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)：它是 BSD 风格、
允许任意目的使用，但**不是 OSI 认证**的许可；只接受 OSI 清单的组织可设
`OPENVBS_UI=console`，走不含 Tk 的路径。

其它已知限制：

- 数组下标不做负值检查（Python 语义下 `a(-1)` 会取最后一个元素）。
- `Exit For` 若嵌在 `Do` 循环里，只跳出最内层循环（Python 没有具名循环）。
- `#python` 代码块在全局作用域运行（不能直接改过程的局部变量）。
- `Class_Terminate` 依赖 Python 的引用计数/GC，时机不保证。
- 没有 `WScript` 宿主对象（`WScript.Echo`、`.Arguments`、`.Sleep`、`.Quit`）。

完整的对象模型清单见 [doc/OOP.md](doc/OOP.md)（继承、重载、`Set/Get`，以及"为什么不做 abstract/static"的说明）。

仓库里还有一个**把网页 IDE 打成 over 包**的示例：[extras/web-ide](extras/web-ide)，
用 over pack / publish / index 就能发布到本地 registry，不依赖任何服务器。

`extras/com-support` 则是**把 COM 支持做成 over 包**的示例：它声明 `platforms: ["win32"]`
与 `python: {"pywin32": ">=306"}`，装进项目后 `over install` 会顺手把 pywin32 装好；
不装也能用 —— OpenVBS++ 自带一份 vendor 副本，Windows 上零配置即可 `CreateObject`。

## 参与贡献 / Contributing

**EN** — The project needs help to survive: superset edge cases, translations,
built-ins, packaging, performance. Contributions use **DCO** (no CLA):
git commit -s. See [CONTRIBUTING.md](CONTRIBUTING.md) and
[dev-docs/good-first-issues.md](dev-docs/good-first-issues.md).

**中文** —— 这个项目需要社区帮忙才能走下去：超集边界情况、翻译、内置函数、
打包、性能。贡献采用 **DCO**（不需要签 CLA），提交时加 -s 即可。详见
[CONTRIBUTING.md](CONTRIBUTING.md) 与 [dev-docs/good-first-issues.md](dev-docs/good-first-issues.md)。

包管理器用法见 [doc/OVER.md](doc/OVER.md)。完整路线图与推迟项见 [dev-docs/ROADMAP.md](dev-docs/ROADMAP.md)。
与微软 VBScript 的法律关系自查见 [doc/LEGAL.md](doc/LEGAL.md)。

## 路线图

- **v0.2.x**：函数与 `Sub`、`Dim`/`ReDim`/`Const`、`Option Explicit`、数组、内置函数 —— 已完成
- **v0.3.x**：单行 `If`、`:`、`On Error`、`With`、`For Each`、按位逻辑与 `\ Mod ^`、VBS 超集检查 —— 已完成
- **v0.4.x**：对象模型（`Class`/`Property`/`Me`/`CreateObject`/`Dictionary`）、`Option Typing` 静态类型检查 —— 已完成
- **v0.5.x**：`Execute`/`Eval`、`FormatNumber` 系列、`Option Compare Text`、`GetObject`、
  默认成员、私有成员可见性、数组元素类型检查 —— 已完成；VBS 超集用例 **0 缺口**
- **v0.6.x**：对象方法 ByRef（统一 Ref 协议）、类继承 `Inherits`、`#python` 嵌入、
  `Option COP`、类型检查补强 —— 已完成
- **v0.7.x–v0.9.x**：over 包管理器 M1/M2（本地源、git 源、锁文件、Python 依赖、bundled 包）、
  `CreateObject` 回归 VBS 语义、真实 COM 支持（`com-support`）、网页 IDE、`--json` 诊断 —— 已完成
- **v0.9.20**：Python 互操作改成随包分发的 `python-embed` over 包，必须显式写
  `<script src="python-embed" />`（`Option COP` 已移除）；删除 `ast2str`；文档改为英文为主 + 中文对照，
  维护者笔记移到 `dev-docs/`
- **v0.9.15**：性能优化（启动懒加载、只读 `ByRef` 不再装箱、无 `On Error` 的过程不建错误帧、`For` 安全时生成 `range`）
- 下一步（方法重载、`Attribute` 其它属性、事件与 `Collection`）
- **v1.0**：稳定语言规范、官方运行时库 `ovpprt`、IDE / LSP 支持

## 许可

Apache License 2.0，见 [LICENSE](LICENSE)；再分发时请保留 [NOTICE](NOTICE)。

### 第三方与商标

- 仓库、wheel、sdist **不含**第三方代码。`pywin32` 只是 Windows 上可选的运行期依赖
  （`pip install "openvbspp[com]"`，或由 `over` 按 `com-support` 包的
  `python: pywin32>=306` 自动安装）。
- 发布到 GitHub Releases 的**单文件 exe** 为了开箱即用，会把 `pywin32`（PSF-2.0）
  的模块与 DLL 一并打包；许可声明见 [NOTICE](NOTICE)。想完全避开第三方代码，
  用 wheel + 自行安装 `pywin32`。
- VBScript 是 Microsoft Corporation 的商标。OpenVBS++ 是独立实现，与 Microsoft
  无关联，也不分发其代码或文档。

### 安全提示

- `<% %>` 内嵌 Python 块（以及 `over` 安装/`pip` 调用）意味着**任意代码执行**能力，
  与 `cscript.exe` 同级的信任模型：**不要运行来源不明的脚本或包**。
- 非 Windows 平台没有 COM，`CreateObject` 会明确报错而不是静默降级。
