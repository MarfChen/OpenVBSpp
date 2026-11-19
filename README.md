# OpenVBS++

![tests](https://img.shields.io/badge/tests-457%20passing-brightgreen)
![version](https://img.shields.io/badge/version-0.9.14-blue)
![license](https://img.shields.io/badge/license-Apache--2.0-blue)
![python](https://img.shields.io/badge/python-3.9%2B-yellow)
![deps](https://img.shields.io/badge/dependencies-none-success)

> **English** — OpenVBS++ compiles VBScript-flavoured source into the standard
> Python AST and runs it on CPython. It is a **superset of VBScript**: every legal
> VBS program must keep working, and extensions may only add syntax.
> New here? Read [CONTRIBUTING.md](CONTRIBUTING.md) and pick a
> [good first issue](doc/good-first-issues.md).
>
> **中文** —— OpenVBS++ 把 VBScript 风格源码编译成标准 Python AST 后在 CPython 上运行。
> 它**必须是 VBScript 的超集**：任何合法 VBS 程序都要继续能跑，扩展只能新增语法。
> 第一次来？看 [CONTRIBUTING.md](CONTRIBUTING.md)，挑一个
> [新手任务](doc/good-first-issues.md)。

>
> **免责声明 / Disclaimer** —— OpenVBS++ 是独立实现，与 Microsoft 无任何关联；
> VBScript 是 Microsoft Corporation 的商标。本仓库不包含 Microsoft 的任何代码或文档。
> OpenVBS++ is an independent reimplementation, not affiliated with or endorsed by
> Microsoft. VBScript is a trademark of Microsoft Corporation.

**OpenVBS++**（Python 包名 `OpenVBSPP`，当前版本 **0.9.14**）是一个以 **VBScript 语法风格**为基础的编译器。
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
语言规范见 [doc/OpenVBS spec.docx](doc/OpenVBS%20spec.docx)，设计初衷见 [doc/plan.docx](doc/plan.docx)。

## 快速开始

```bash
# 仓库自带虚拟环境（Python 3.9+）
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
```

| 参数 | 说明 |
|---|---|
| `script.vbs` | 要运行的文件；省略则运行内置示例 |
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
编辑器集成（VSCode/LSP）的方案见 [doc/IDE_LSP_DESIGN.md](doc/IDE_LSP_DESIGN.md)。

## 安装与打包

| 目标 | 命令 | 产物 |
|---|---|---|
| 本地开发 | `pip install -e ".[all]"` | 可编辑安装 + 全部开发依赖 |
| PyPI 包 | `.\tools\build.ps1 -Target wheel` | `dist/*.whl` + `dist/*.tar.gz` |
| Windows exe | `.\tools\build.ps1 -Target exe` | `dist\openvbs.exe`（单文件，无需 Python） |
| Windows COM 支持 | `pip install "openvbspp[com]"` | `pywin32`（也可以让 `over install` 自动装） |

发布流程（TestPyPI 验证 → PyPI → 打 exe）见 [doc/RELEASE.md](doc/RELEASE.md)。
项目已预留好：`[project.scripts]` 入口、`MANIFEST.in`、`packaging/openvbs.spec`
（PyInstaller 配置）、`tools/build.ps1`。

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
| Python 互操作 | `#python ... #endpython` 代码块、`Option COP` 公共对象池 |
| 动态执行 | `Execute` / `Eval`（复用编译器，在全局作用域运行） |
| 其它 | `Erase`、`Option Compare Text`、`FormatNumber` 系列、`vbCrLf` 等预定义常量、约 70 个内置函数 |

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
| 转换 | `CStr CInt CLng CDbl CBool CByte` |
| 数学 | `Abs Int Fix Sgn Sqr Round Rnd Randomize` |
| 数组 | `Array UBound LBound` |
| 判断 | `IsNumeric IsEmpty IsNull IsObject IsArray IsDate TypeName` |
| 日期 | `Now Date Time Timer Year Month Day Hour Minute Second Weekday DateAdd DateDiff` |
| 交互 | `MsgBox`（打印到标准输出）、`InputBox`（读标准输入） |

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
├── plugins/debug_tools.py  ast2str()：把生成的 Python AST 反编译成源码
└── vendor/com-support/ 随包分发的 COM 支持包（Windows 零配置可用）
examples/               示例脚本（控制流 / 过程 / 数组 / 语法糖 / 错误处理 / 类 / 类型 / FizzBuzz）
tests/                  回归测试（457 个用例）
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
- `tests/test_interop.py`：`#python` 代码块与 `Option COP` 对象池。
- `tests/test_vbs_superset.py`：**VBS 超集一致性**——100+ 条合法 VBS 片段必须跑通；
  v0.5 起已没有 expectedFailure（缺口全部补齐），新增缺口时按老规矩加。
- `tests/test_over.py` / `tests/test_over_registry.py`：over 的依赖图、锁文件、本地源与 registry。
- `tests/test_over_python.py`：over 声明 Python 依赖时的合并、钉版本、frozen 还原与卸载。
- `tests/test_com.py`：COM 桥的定位优先级、开关、错误翻译与适配层行为（不依赖真实 COM）。

## 与规范的差距 / 已知限制

**尚未实现**（`doc/VBS_SUPERSET.md` §4）：

- `Collection` 内置类（VBScript 本身没有，属于 VBA）
- `Declare`（声明外部 DLL 函数）

COM 宿主对象（`Excel.Application`、`WScript.Shell` 等）**已经支持**：Windows 上由
`com-support` 包真连，非 Windows 平台、`OPENVBS_COM=0`、或没装 `pywin32` 时会给出明确报错
（不再按"未实现"处理）。

其它已知限制：

- 数组下标不做负值检查（Python 语义下 `a(-1)` 会取最后一个元素）。
- `Exit For` 若嵌在 `Do` 循环里，只跳出最内层循环（Python 没有具名循环）。
- `#python` 代码块在全局作用域运行（不能直接改过程的局部变量）。
- `Class_Terminate` 依赖 Python 的引用计数/GC，时机不保证。

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
[doc/good-first-issues.md](doc/good-first-issues.md).

**中文** —— 这个项目需要社区帮忙才能走下去：超集边界情况、翻译、内置函数、
打包、性能。贡献采用 **DCO**（不需要签 CLA），提交时加 -s 即可。详见
[CONTRIBUTING.md](CONTRIBUTING.md) 与 [doc/good-first-issues.md](doc/good-first-issues.md)。

包管理器用法见 [doc/OVER.md](doc/OVER.md)。完整路线图与推迟项见 [doc/ROADMAP.md](doc/ROADMAP.md)。

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
- **v0.9.14**：下一步（方法重载、`Attribute` 其它属性、事件与 `Collection`）
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
