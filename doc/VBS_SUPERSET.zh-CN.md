# VBS 超集一致性

> **In English** — The superset matrix: every VBScript construct and whether
> OpenVBS++ supports it (✅), deviates semantically (⚠️) or declares it
> unsupported (❌). Sections marked "extension" are additions that VBScript
> itself would reject, so they cannot break existing scripts.


> 项目铁律：**OpenVBS++ 必须是 VBScript 的超集。**
> 任何合法的 VBScript 程序都必须能编译运行，除非它明确列在本文的"已知缺口"里。
> 扩展语法只能**新增**，绝不允许改变已有 VBS 语义。

## 1. 这条规则怎么被执行

| 手段 | 位置 |
|---|---|
| 可执行的用例矩阵：合法 VBS 片段必须跑出预期结果 | `tests/test_vbs_superset.py` → `VBS_FEATURES` |
| 已知缺口用 `expectedFailure` 标注；缺口补齐时该用例会变红，提醒同步文档 | `tests/test_vbs_superset.py` → `TestKnownGaps` |
| 文档矩阵（本文） | `doc/VBS_SUPERSET.md` |
| 改动检查清单 | `AGENTS.md` |

新增特性时：先在 `VBS_FEATURES` 里写用例，再实现，最后更新本文的状态列。

## 2. 支持矩阵

图例：✅ 支持 · ⚠️ 支持但有语义差异（见 §3） · ❌ 尚未实现（见 §4）

### 2.1 词法与语句

| VBS 特性 | 状态 | 说明 |
|---|---|---|
| 换行分隔语句 | ✅ | |
| `:` 分隔多条语句 | ✅ | `a = 1 : b = 2` |
| `_` 行续接 | ✅ | 行尾 `空格+_` |
| `'` 注释 / `rem` 注释 | ✅ | 整行与行尾 |
| CRLF / CR / LF 换行 | ✅ | 自动归一化 |
| 关键字大小写不敏感 | ✅ | 关键字与标识符都按小写归一 |
| UTF-8 / BOM | ✅ | CLI 用 `utf-8-sig` 读取 |

### 2.2 数据与表达式

| VBS 特性 | 状态 | 说明 |
|---|---|---|
| 变量无需声明 | ✅ | |
| `Dim` / `ReDim [Preserve]` / `Const` | ✅ | 数组下界 0、上界含 |
| `Option Explicit` | ✅ | 编译期检查，带行号 |
| 整数 / 浮点 / 字符串 / 布尔 | ✅ | |
| `Empty` / `Null` / `Nothing` 字面量 | ✅ | 三个哨兵 |
| `+ - * /` | ✅ | 按 VBS 规则强制转换：`+` 只在两侧都是文本时拼接，`/` 恒为 Double；`Option Native` 可关掉转换（§2.7.1） |
| `\` 整除、`Mod` | ✅ | 向零截断 |
| `^` 幂 | ✅ | 恒为 Double；左结合（与 VB 同级左到右一致） |
| `&` 字符串连接 | ✅ | `Null` 当空文本；`Null & Null` 仍是 `Null` |
| `= <> < > <= >=` | ✅ | `=` 在表达式里就是相等比较 |
| `==` 作为相等比较 | ⚠️ | 我们的扩展；VBS 没有，但接受它不影响 VBS 代码 |
| `Is` 对象身份比较 | ✅ | `x Is Nothing` |
| `And Or Not Xor Eqv Imp` | ⚠️ | 按位语义（与 VBS 一致），见 §3.1 |
| `True` / `False` | ✅ | 真 = -1、假 = 0 |
| `vbCrLf` 等预定义常量 | ✅ | 字符串常量（`vbCrLf` `vbCr` `vbLf` `vbTab` `vbNullString`）、`FormatDateTime` 命名格式、`MsgBox` 的按钮/图标/返回值常量（`vbOKOnly` … `vbNo`）、`vbObjectError`，以及 `VarType` 的结果常量（`vbString` `vbLong` `vbArray` …） |

### 2.3 控制流

| VBS 特性 | 状态 |
|---|---|
| `If ... Then ... End If`（块） | ✅ |
| 单行 `If cond Then stmt [Else stmt]` | ✅ |
| `ElseIf` 多次 | ✅ |
| `Select Case` / `Case 值, 区间 To, Is 比较` / `Case Else` | ✅ |
| `For i = a To b [Step s] ... Next` | ✅ |
| `For Each x In 集合 ... Next` | ✅（数组与内置集合） |
| `Do [While\|Until] ... Loop [While\|Until]` | ✅ |
| `While ... Wend` | ✅ |
| `Exit Do` / `Exit For` / `Exit Sub` / `Exit Function` | ✅ |
| 单独的 `End`（结束整个脚本） | ✅ | 立即结束脚本；`On Error Resume Next` 与 `Try` 都接不住 |
| `Stop`（调试断点） | ⚠️ | 接受但忽略 —— 这里没有调试器可以停下来 |

### 2.4 过程与作用域

| VBS 特性 | 状态 | 说明 |
|---|---|---|
| `Sub` / `Function` / `End Sub` / `End Function` | ✅ | 仅脚本级（VBS 也是这样） |
| 参数 `ByRef`（默认）/ `ByVal` / 数组参数 `a()` | ✅ | ByRef 用 `Ref` 盒子 + 调用点写回 |
| 函数返回值（给函数名赋值） | ✅ | 未赋值返回 `Empty` |
| 递归 | ✅ | |
| 三种调用写法 `Foo` / `Foo 1,2` / `Call Foo(1,2)` | ✅ | |
| `Let` / `Set` 语句 | ⚠️ | 语法已接受；`Set` 的对象语义随对象模型实现（见 §4） |
| 脚本级变量在过程内读写 | ✅ | 自动生成 `global` 声明 |

### 2.5 错误处理

| VBS 特性 | 状态 | 说明 |
|---|---|---|
| `On Error Resume Next` | ✅ | 逐语句 try/except + `_ovpp_err` 帧栈 |
| `On Error GoTo 0` | ✅ | |
| `Err.Number` / `.Description` / `.Source` | ✅ | |
| `Err.Clear` / `Err.Raise` | ✅ | |
| 错误号沿用 VBS 约定 | ⚠️ | 见 §3.3 |

### 2.6 对象与集合

| VBS 特性 | 状态 | 说明 |
|---|---|---|
| `With ... End With` | ✅ | 表达式只求值一次 |
| `.member` / `.Method()` 语法 | ✅ | 字段/属性直接取，方法调用走 `member_call` |
| `Class ... End Class` | ✅ | 编译成 `_ovpp_cls_<name>` |
| 字段 `Public` / `Private` / `Dim` | ✅ | 私有字段/方法会改名，外部读写都会被拒绝 |
| `Property Get/Let/Set` | ✅ | 映射到 Python `@property` / setter（`Set` 与 `Let` 合并，见 [OOP.md](OOP.md)） |
| `Me` / `Class_Initialize` / `Class_Terminate` | ✅ | `Me` → `self`；构造/析构自动调用 |
| `New Foo` | ✅ | |
| `Set x = obj` / `Set x = Nothing` / `x Is Nothing` | ✅ | |
| `CreateObject("脚本类")` | ✅ | 编译期解析或运行时注册表 |
| `CreateObject("Scripting.Dictionary")` | ✅ | `ovpprt/collections.py` |
| `Collection` | ❌ | VBScript 本身没有内建 Collection（属于 VBA） |
| `GetObject` | ✅ | `GetObject(path)`、`GetObject(path, class)`、`GetObject(, class)`；`winmgmts:` 这类 moniker 在 Windows 上由 `com-support` 包解析；连不上时退回创建该类 |
| `CreateObject` 创建 COM 对象 | ✅ | v0.9.14 起 Windows 上由 `com-support` 包真连（pywin32）；非 Windows 明确报错 |
| COM 错误号与 `TypeName` | ✅ | 找不到 ProgID 报 429（`ActiveX component can't create object`），连接失败报 432，`TypeName(obj)` 返回组件类名（如 `FileSystemObject`、`SWbemServicesEx`） |
| 默认成员自动取值 | ✅ | `Attribute X.VB_UserMemId = 0`；`x = obj` 取默认属性，`Set x = obj` 保持引用 |
| 对象方法的 `ByRef` 写回 | ✅ | 方法参数统一走 `Ref` 协议（见 §3.7） |
| 方法覆盖（虚方法） | ✅ | 默认即可覆盖，按 MRO 解析 |
| 写入未声明的成员 | ✅ | 抛 438（`__setattr__` 白名单） |
| `Class Child Inherits Parent` | ✅ | 超集扩展；方法覆盖、字段/属性继承、`_ovpp_attrs` 合并 |

### 2.7 类型系统（OpenVBS++ 扩展）

| 特性 | 状态 | 说明 |
|---|---|---|
| `Option Typing` | ✅ | 开启后编译期检查；不开启时注解被忽略 |
| `x -> Integer` / 参数 / 返回值注解 | ✅ | `->` 与冒号互不冲突 |
| 赋值 / 实参 / 返回值类型检查 | ✅ | 保守检查，`Variant` 通配 |
| 数组元素类型 `Dim a(5) -> Integer` | ✅ | 下标读写都检查元素类型 |
| 类字段与属性类型 | ✅ | 含继承链上的字段、`Property Let` 的参数类型 |

### 2.7.1 `Option Native`（OpenVBS++ 扩展）

`Option Native` 关掉 VBScript 的算术强制转换：文本操作数直接报**错误 13**，而不是被转换。
`Option Typing` 会隐含打开它 —— 带类型声明的代码同样不该悄悄转换。不写这个 Option 时行为
完全不变，所以它不可能破坏任何 VBS 代码。

| 写法 | 效果 |
|---|---|
| `Option Native` | 原生算术：只用原生数值类型（等价于 `Option Native On`） |
| `Option Native On` | 显式写法 |
| `Option Native Off` | 显式恢复 VBScript 的强制转换 |
| `Option Typing` | 开启类型检查**并且**开启原生算术 |

| 运算符 | 默认（VBScript） | `Option Native` |
|---|---|---|
| `+` | `"1" + "2"` → `"12"`、`1 + "2"` → `3` | 两者都报 13 |
| `- * / ^` | `"6" - "4"` → `2`、`"2" * 3` → `6` | 两者都报 13 |
| `\` `Mod` | `"6" \ 2` → `3`、`"6" Mod 4` → `2` | 两者都报 13 |
| 一元 `+` `-` | `-"3"` → `-3` | 报 13 |
| `&` | `"a" & 1` → `"a1"` | 不变（它本来就是文本运算） |
| 比较运算 | 见 §4 | 不变 |
| `Empty` | `Empty + 1` → `1` | 不变（Empty 是默认值，不是被转换的值） |
| `Null` | `Null + 1` → `Null` | 不变（是传播，不是转换） |

`Option` 语句描述整个脚本，所以必须写在所有其它语句之前 —— 否则迟到的
`Option Native Off` 会悄悄改变它上面代码的编译方式。VBScript 自己对 `Option Explicit`
就是这个规则；报错是 1002（`'Option Native' must appear before any other statement`）。

#### 为什么叫 `Native` 而不是 `Strict`

这个关键字在 v0.9.26 之前叫 `Option Strict`。值得解释一下改名：一个关键字要同时承担三件事
——说清它约束什么、说清它给你什么、以及不能误导任何人。

- **约束什么。** "Strict" 只说了"更严"，没说严在哪。"Native" 说得准确：参与算术的只能是
  原生数值类型。
- **给你什么。** 这个 Option 一半是为了规范，一半是为了性能。VBScript 语义逼着每个
  `+ - * /` 都进运行时函数检查两个操作数的类型（每次约 60 ns）。打开这个选项等于给出一个
  承诺 —— "两边都是数字" —— 编译器用**原生**运算回报你。`Native` 把这个交易写在名字里；
  `Strict` 只宣传了代价，于是大家就懒得打开。
- **不能误导什么。** 在 VB6 / VBA / VB.NET 里，`Option Strict On` 是**编译期**检查的固定
  拼法：禁止隐式窄化转换、禁止后期绑定、要求 `Dim x As T`。本项目那件事另有名字
  （`Option Typing`），这一个管的是运行时的运算符行为。拿熟悉的词去装不同的含义，正是
  迁移者最容易被坑的地方 —— 从 VB6 过来的人会打开 `Option Strict` 等 `Dim` 行报错，结果在
  `"1" + "2"` 处撞上错误 13。

这个名字还给这个选项留了扩展余地：万一 §4.1 的比较缺口将来补上，"原生比较"是这个词的
自然延伸；而叫 `Native Arithmetic` 就把关键字钉死在一个运算符组上了。

`Option Strict` 不作为别名保留：这个选项在发布之前就改名了，世上没有脚本依赖旧拼法。

### 2.8 结构化异常（OpenVBS++ 扩展）

| 特性 | 状态 | 说明 |
|---|---|---|
| Try ... End Try | ✅ | 块作用域，出错时块内异常交给 Catch |
| Catch e | ✅ | e 是错误快照（Number / Description / Source / Line） |
| Catch e When 条件 | ✅ | 按错误号过滤；多个 When 依次匹配，最后可跟通配 Catch |
| Finally | ✅ | 正常离开、Exit Sub、错误传播、重抛都会执行 |
| Err.Line | ✅ | 扩展字段：出错行号（VBS 不提供） |
| 与 On Error 共存 | ✅ | Try 优先；离开块后恢复块外的 On Error 状态 |

要点：Try / Catch / Finally / When 都是**上下文关键字**——
只有在语句开头（Try 还必须独占一行）才被识别，所以
Try = 1、Catch = 2 这类既有 VBS 代码仍然合法（有测试覆盖）。

### 2.8.1 多文件（OpenVBS++ 扩展）

| 写法 | 语义 |
|---|---|
| <script src="lib.vbs"> | **合并**：目标文件的过程/类进入当前全局作用域，直接调用 |
| <namespace src="lib.vbs" alias="u"> | **命名空间**：u.Double(21)；只导出 Public 过程 |
| `alias` 省略 | 用文件名（不含扩展名）当别名 |
| `nickyname` / `tar` | `alias` 的同义属性 |

规则：路径相对**声明所在文件**解析；同一个文件重复 <script> 只合并一次；
循环依赖在编译期报错并打印链条；每个文件独立编译（行号、文件名都正确）；
<namespace> 的文件里若又用了 <namespace>，需要先被引用方已经绑定（见限制）。

### 2.9 Python 互操作（OpenVBS++ 扩展）

| 特性 | 状态 | 说明 |
|---|---|---|
| `#python ... #endpython` | ✅ | 预处理抽出代码块，原地留占位语句，**行号不变** |
| `COP` 公共对象池 | ✅ | `COP` 是全局字典，Python 侧 `from OpenVBSPP.ovpprt.cop import COP` 拿到同一个对象 |

自 v0.9.20 起 Python 互操作必须显式导入：在脚本最前面写 `<script src="python-embed" />`。
`python-embed` 是随 OpenVBS++ 分发的 over 包；原来的开关 `Option COP` 已移除，
写它现在是一个语法错误，报错信息会指向这个导入。

### 2.10 内置函数

VBScript 的内置函数除 `LoadPicture` 外全部可用（`LoadPicture` 只用来给出明确报错，见 §4）；
内置语句（`Execute`、`ExecuteGlobal`）同样可用。

- 字符串：`Len Left Right Mid UCase LCase LTrim RTrim Trim StrReverse Space String Asc Chr InStr InStrRev Replace Split Join Filter StrComp Hex Oct`
- 转换：`CStr CInt CLng CDbl CBool CByte CCur CSng CDate`
- 数学：`Abs Int Fix Sgn Sqr Round Rnd Randomize Sin Cos Tan Atn Exp Log`
- 数组：`Array UBound LBound`
- 判断：`IsNumeric IsEmpty IsNull IsObject IsArray IsDate IsMissing TypeName VarType`
- 日期：`Now Date Time Timer Year Month Day Hour Minute Second Weekday DateAdd DateDiff DatePart DateSerial DateValue TimeSerial TimeValue MonthName WeekdayName`
- 格式化：`Format FormatNumber FormatCurrency FormatDateTime FormatPercent`
- 运行时信息：`RGB ScriptEngine ScriptEngineMajorVersion ScriptEngineMinorVersion ScriptEngineBuildVersion GetLocale SetLocale`
- 交互：`MsgBox InputBox`（交互式运行时弹真窗口，否则是控制台替身，见 §3.10）

自 v0.9.21 起还支持 `#date#` 字面量（例如 `#2020-05-06#`、`#1/2/2020 3:04:05 PM#`）
以及 `&H` / `&O` 数字字面量。

| 函数 | 状态 | 说明 |
|---|---|---|
| `Execute` / `ExecuteGlobal` | ✅ | 语句形式；代码在全局作用域执行 |
| `Eval` | ✅ | 表达式形式 |
| `GetRef` | ✅ | `f(21)` 与 `f()` 都能调用引用；直接写 `f` 读到的是对象本身 |
| `LoadPicture` | ⚠️ | 抛 `OVPPObjectNotSupportedError`：没有图形图像库 |

## 3. 已知语义差异

这些地方我们**有意**与 VBS 不同，或者无法完全一致。

### 3.1 逻辑运算符是按位的

VBScript 的 `And/Or/Not/Xor` 作用在整数上是位运算，作用在布尔上是逻辑运算。
True 就是 -1、False 就是 0，两套语义天然统一，所以：

```vb
echo 5 And 3        ' 1（位运算）
echo Not 0          ' -1（位运算）
echo (1 = 1) And (2 = 2)   ' True（两边都是布尔，保持布尔）
echo True And 5     ' 5（混合：按位）
```

我们**完全按这个规则实现**，因此 `echo -1 And 0` 输出 `0` 而不是 `False`。

### 3.2 文本比较区分大小写

VBS 默认 `Option Compare Binary`，`"A" = "a"` 为假——我们一致。
`Option Compare Text`（不区分大小写）尚未实现（见 §4）。

### 3.3 错误号

错误类沿用 VBS 的编号（11 除零、13 类型不匹配、35 过程未定义、94 Null 用法非法、438 对象不支持、
450 参数个数不对、1002 语法错误……），但"变量未定义"用的是 1001 而不是 VBS 的 500，
这是项目既有的约定，改动会破坏已有脚本的 `Err.Number` 判断。

这条错误只在 `Option Explicit` 下出现，与 VBS 一致。没有它时，读取一个从未声明、也从未赋值的名字
得到的是 `Empty`（`TypeName(z) = "Empty"`、`z & "x" = "x"`），而且编译器不会留下任何运行时检查 ——
这些名字在模块顶部被初始化一次，所以紧凑循环里没有额外开销。

### 3.4 数字显示

`Echo` 与 `&` 都按 VBS 习惯格式化：`3.0` → `3`、`-0.0` → `0`。
Python 原生的 `3.0` 不会被输出。

非整数的 Double 保留 Python 的全部有效位，而不是 VBS 的 15 位：
`echo 1 / 3` 输出 `0.3333333333333333`，VBS 输出 `0.333333333333333`。
内存里的值完全相同（都是 Double），只有文本末尾的位数不同。

### 3.5 数组下标不做负值检查

`a(-1)` 在 Python 里会取最后一个元素，VBS 会报"下标越界"。

### 3.6 `Exit For` 的嵌套语义

`Exit For` 如果嵌在 `Do` 循环里，只跳出最内层循环（Python 没有具名循环）。

### 3.7 对象方法的 ByRef 走统一的 Ref 协议

VBS 里 `obj.Method x` 的 ByRef 参数会写回 `x`。v0.6 起所有对象方法调用都编译成：

```python
_ref_1 = Ref(x)
_ovpp_obj.call_method(obj, "method", [_ref_1])
x = _ref_1.v            # 写回
```

方法体内部：ByRef 参数直接用 `参数.v`，ByVal 参数在函数体开头解包成局部变量。
内置对象（`Dictionary` 等）不是 VBS 类，`call_method` 会自动取 `.v` 按值调用，
所以同一套协议对两边都成立。

### 3.8 `Option Typing` 的类型系统是保守的

开启后只做"能确定就报错"的检查：赋值、实参、返回值。
`Variant`、`Empty`、`Null`、`Object` 会通配一切，所以不会误伤合法程序。
未开启时注解被忽略，行为与普通 VBS 完全一致。

### 3.9 `Resume` / `Resume Next` 需要错误处理段

`Resume`（重试失败语句）与 `Resume Next`（继续下一条）自 v0.9.5 起已实现（见 §6）：
含这两个语句的过程会编译成"语句下标状态机"。它们只在带 `On Error GoTo <label>`
处理段的过程里有意义，缺少处理段时是编译错误。

注意不要与 `On Error Resume Next` 混淆：后者是模式语句，任何地方都能用，不需要处理段。

### 3.10 `MsgBox` 与 `InputBox` 在能弹窗时弹真窗口

VBScript 的 `MsgBox` 会弹窗并返回用户点击的按钮，`InputBox` 会弹出输入框。
OpenVBS++ 通过自带的 `ui-support` 包（由 Tkinter 绘制）实现这两者，在不适合弹窗的
场合自动退回控制台替身。

只有同时满足以下条件才弹窗：

- 脚本由 `openvbs` 命令行启动（不是 IDE，也不是在 Python 程序里嵌入 `vbs_run`）；
- stdin 与 stdout 是终端，且没有设置 `CI`；
- `ui-support` 包能加载，且 Tkinter 能创建窗口。

弹窗时参数含义与 VBScript 一致：

- `MsgBox 文本 [, 按钮] [, 标题]` 按请求的按钮组合（`vbOKOnly` … `vbRetryCancel`）、
  图标、默认按钮与模态标志弹出消息框，返回 `1`–`7`（`vbOK`、`vbCancel`、`vbAbort`、
  `vbRetry`、`vbIgnore`、`vbYes`、`vbNo`）。
- `InputBox(提示 [, 标题] [, 默认值])` 弹出输入框并预填 `默认值`；点取消或关闭窗口
  返回 `""`。

其余场合走控制台行为：`MsgBox` 打印文本并返回 `1`，`InputBox` 打印提示并从 stdin
读一行。CI、管道、包缺失或 Tk 不可用时都是这种情况 —— 批处理永远不会卡在窗口上。

`OPENVBS_UI` 可以覆盖这个判断：`console`（或 `off`）始终用替身，`gui`（或 `on`）要求
弹窗，默认的 `auto` 按上面的规则决定。`python -m OpenVBSPP check` 之类的运行从不弹窗。

### 3.11 `Declare` 与原生代码

`Declare Function/Sub` 绑定的是 DLL 入口；OpenVBS++ 跑在 Python 虚拟机上，没有这种绑定，
因此 `Declare` 会明确报语法错误，而不是悄悄什么都不做。

**短期决定（v0.9.22）：暂不支持。** 已经调研过用 `ctypes` 实现：技术上可行，但它允许脚本
调用任意原生代码，而签名写错会让进程以访问违例直接退出 —— `On Error Resume Next` 与
`Try ... Catch` 都接不住。这个代价目前不值得承担，因此 `Declare` 不进语言，也没有开关能打开。
调研结论记录在 [../dev-docs/ROADMAP.zh-CN.md](../dev-docs/ROADMAP.zh-CN.md) 的 7c 项。

目前可用的原生代码路径是：

- COM：Windows 上用 `com-support` 包执行 `CreateObject("ProgID")`。
- 原生 Python：写一个 `over` 包，或用 `#python` 代码块（见 §2.9）。

### 3.12 `End` 与 `Stop`

- 单独一行的 `End` 立即结束整个脚本，与 VBScript 的 `End` 语句一致。它编译成一个
  继承自 `BaseException` 的 `ScriptEnd` 异常，因此 `On Error Resume Next` 与
  `Try ... Catch` 都接不住；运行器把它当作正常结束（进程退出码 0）。而 `End If`、
  `End Sub`、`End With`、`End Select`、`End Try`、`End Function`、`End Property`、
  `End Class` 仍然保持「结束当前块」的含义。
- `Stop` 在 VBScript 里是调试断点。这里没有调试器可停，因此接受该语句但什么也不做。
  它不算语法错误，因为真实脚本常故意留着 `Stop` 调用。

### 3.13 `WScript` 宿主对象

用 `openvbs` 启动的脚本会得到与 Windows Script Host 相同的全局 `WScript` 对象，因此按老习惯
写 `WScript.Echo` 的代码可以继续用：

| 成员 | 行为 |
|---|---|
| `WScript.Echo a, b` | 用空格分隔输出各值，与 `Echo` 语句一致 |
| `WScript.Arguments` | 脚本路径之后的参数；`Count`、`Arguments(0)`（从 0 开始）、`Item(n)`、`For Each` 都可用。下标越界报错误 9 |
| `WScript.Arguments.Named` | `/name:value` 与 `/name=value` 形式的具名参数：`Exists(name)`、`Item(name)`、`Count` |
| `WScript.ScriptName`、`.ScriptFullName` | 正在运行脚本的文件名与完整路径（`-c`/标准输入运行时代码为 `""`） |
| `WScript.StdOut`、`.StdErr`、`.StdIn` | `TextStream` 对象：`Write`、`WriteLine`、`WriteBlankLines`；`StdIn` 另有 `ReadLine`、`ReadAll`、`AtEndOfStream`、`Skip`、`SkipLine`（读到文件尾时 `ReadLine` 报错误 62） |
| `WScript.Name`、`.Version` | `"Windows Script Host"` 与 `"5.812"`，老脚本会用它们做判断 |
| `WScript.FullName`、`.Path` | 运行脚本的解释器及其所在目录（OpenVBS++ 是 Python 宿主，不是 `cscript.exe`） |
| `WScript.Interactive`、`.Timeout` | `Interactive` 表示 stdin/stdout 是否为终端；`Timeout` 只保存、不强制 |
| `WScript.Sleep ms` | 休眠；`WScript.Quit code` 以该退出码结束 |
| `WScript.CreateObject`、`.GetObject` | 与同名全局函数一致，错误码见 §2.6 |

`WScript` 本身是只读的：给它赋值会报语法错误。命令行参数从
`openvbs script.vbs a b` 形式传入；嵌入运行时可写 `vbs_run(..., args=[...])`。

## 4. 已知缺口

`tests/test_vbs_superset.py` 把剩下的缺口登记成 `expectedFailure`，所以补齐时测试会故意变红，
直到把用例移进 `VBS_FEATURES`。之前的缺口（对象模型、`Execute`/`Eval`、格式化函数、
`Option Compare Text`、`GetObject`、默认成员、算术强制转换、Null 传播）已全部补齐。

### 4.1 比较运算不做强制转换

VBScript 在数字与"数字文本"比较时会转换。我们按类型比较，也就是 Python 的规则：

| 表达式 | VBScript | OpenVBS++ |
|---|---|---|
| `1 = "1"` | `True`（按数值比） | `False` |
| `2.5 = "2.5"` | `True` | `False` |
| `"2" > 10` | `False` | 错误 13 |
| `"abc" = 5` | 错误 13 | `False` |
| `"abc" < "abd"` | 文本比较 | 文本比较 ✅ |
| `1 = 1`、`"A" = "a"` | 同 VBScript | ✅ |

为什么还没补：得让每次比较都走运行时函数，约 60 ns/次 —— 100 万次的循环会从约 92 ms
涨到约 150 ms。补它等于用速度换 VBScript 保真，而这个角落主要出现在"把 `InputBox`
的字符串和数字比"的脚本里。

### 4.2 不支持或主动不做

| 项 | 说明 |
|---|---|
| COM 宿主对象（`Excel.Application`、`WScript.Shell` 等） | 自 v0.9.14 起在 Windows 上由 `com-support` 包真连；其它平台抛 `OVPPObjectNotSupportedError`，不模拟 COM |
| `WScript` 宿主对象 | 自 v0.9.23 起提供（见 §3.13）：`Echo`、`Arguments`（从 0 开始，另有 `Named`）、`ScriptName`、`StdOut`/`StdErr`/`StdIn`、`Sleep`、`Timeout`、`Quit`、`CreateObject`、`GetObject`。它描述的是 OpenVBS++ 控制台宿主，所以 `FullName`/`Path` 指向 Python 解释器而不是 `cscript.exe`，`Interactive` 取决于终端 |
| `Collection` 内置类 | VBScript 本身没有（属于 VBA） |
| `Attribute` 的其它属性（`VB_Name` 等） | 只解析 `VB_UserMemId` |
| `MsgBox` / `InputBox` | 交互式运行时有 `ui-support` 包就弹真窗口（见 §3.10）；CI、管道与嵌入运行时退回控制台：`MsgBox` 打印文本并返回 1，`InputBox` 从 stdin 读一行 |
| `LoadPicture` | 函数存在，只为让老脚本得到明确报错：没有图形图像库（`OVPPObjectNotSupportedError`） |
| `ExecuteGlobal` 在运行期新定义的过程 | 可以修改全局变量，但无法按名字调用 —— 过程名在编译期就绑定了。需要间接调用时用 `GetRef` 引用已有过程，或把代码放进独立模块 |
| 向后 `GoTo` | 只编译向前跳转；跳到更早的标签会报 `Recursive GoTo ... is not supported` |
| `Declare`（声明 DLL） | 短期内有意不支持：运行时是 Python，不是 DLL 宿主，而且原生调用写错会让进程无法捕获地退出（见 §3.11）。连 COM 用 `CreateObject`，要原生代码用 over 包 |

## 5. 改动检查清单

改编译器之前/之后都过一遍：

1. 新语法是否让某段合法 VBS 改变了行为？若是，改实现或写进 §3。
2. `tests/test_vbs_superset.py` 是否新增了对应用例？
3. 有缺口被补齐吗？从 §4 表格里删掉、用例移进 `VBS_FEATURES`、更新本文的状态说明。
4. 全量测试（`python -m unittest discover -s tests`）是否仍然全绿？
5. README / 本文的"支持矩阵"是否同步？


> 完整的对象模型清单（含"不适用"的关键字说明）见 [OOP.md](OOP.md)。

## 6. 状态更新（v0.9.5 - v0.9.26）

Resume / Resume Next 已实现：含 Resume 的过程会编译成"语句下标分派"的状态机，
每条语句一个 _pc 分支，出错时记录失败下标并跳到 On Error GoTo 的处理段；
Resume 回到失败语句，Resume Next 跳到下一条。因此 Resume 在嵌套 If / 循环里同样可用。
不含 Resume 的过程保持原来的顺序编译，行为与性能都不受影响。

经过 v0.9.21 的审计，VBScript 原生语句形式已全部支持，只有一处有意例外：`Declare`。
自 v0.9.22 起 `MsgBox` / `InputBox` 在交互式命令行运行时会弹出真实的 Tk 对话框
（见 §3.10），自 v0.9.23 起提供 `WScript` 宿主对象（见 §3.13），因此剩下的缺口都属于
宿主细节而不是语法：`LoadPicture` 报告没有图形图像库、`ExecuteGlobal` 在运行期定义的
过程无法按名字调用（改用 `GetRef`）、只允许向前 `GoTo`。上面 §4 的表格就是完整清单。

### 6.1 v0.9.24 补齐的取值语义

v0.9.23 审计发现的、位于"取值层"的三处差异，已按 `cscript` 的实测结果修正 ——
算术、Null、未声明变量的行为现在与 VBScript 一致，而不是跟着 Python 走：

| 规则 | 修正前 | 现在 |
|---|---|---|
| `+ - * / ^` 强制转换 | 直接映射成 Python 运算符，于是 `"2" * 3` 得到 `"222"`、`1 + "2"` 报类型不匹配 | `+` 只在两侧都是文本时拼接；`- * / ^` 把文本转成数字；`/` 与 `^` 恒为 Double；非数字字符串报 13 |
| Null 经过内置函数 | `Len(Null)` 为 4、`TypeName(Null)` 为 `"NullType"`、`UCase(Null)` 为 `"NULL"` | VBS 规定传播 Null 的那一组（`Len`、`UCase`、`Abs`、`Trim`、日期分量……）返回 Null；转换与数学组（`CStr`、`CInt`、`Sgn`、`Sqr`……）报 94；`Format*` 组报 13。`&` 把 Null 当空文本，但 `Null & Null` 仍是 Null |
| 读取未声明变量 | 报"变量未定义（1001）" | 除非开了 `Option Explicit`，否则读作 `Empty`（见 §3.3）。编译器在模块顶部一次性初始化这些名字，运行期没有任何检查 |

### 6.2 v0.9.26 新增 `Option Native`

VBScript 的算术会把文本转成数字。这对 VBScript 是正确行为，但也让笔误变成静默结果
（v0.9.24 之前 `"2" * 3` 就是 `"222"`）。`Option Native` 是可选的答案：默认行为不变，
但项目可以要求"必须是数字"，于是 `1 + "2"` 报 13 而不是变成 `3`。`Option Typing` 会把它
一起打开 —— 类型注解与隐式转换本来就互相矛盾。完整的运算符对照表见 §2.7.1。

同一版还登记了一处尚未补齐的保真缺口：比较运算仍然按类型比较、不做强制转换（§4.1）。
