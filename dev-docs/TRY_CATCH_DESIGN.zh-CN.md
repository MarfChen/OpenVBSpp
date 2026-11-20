# Try / Catch 设计报告 · 带 VBS 味道的结构化异常

> **[English](TRY_CATCH_DESIGN.md)** | 中文

> **状态：v0.8.0 已实现**（P0–P2）。实现与设计一致，差异记录在本文件末尾 §7。
>
> **English abstract** — We *could* just add `Try … Catch … Finally … End Try`.
> But VBScript already owns a distinctive error model: `On Error Resume Next`
> plus a global `Err` object, and the "check `Err.Number` after the call"
> idiom is the language's fingerprint. This report proposes a design where
> structured exceptions and `On Error` are **the same machinery viewed from two
> sides**, so mixing them in one procedure stays predictable, and VBS programs
> keep working untouched.
>
> 中文正文在后。核心观点：**不要另造一套异常**，而是让 `Try/Catch` 成为
> `On Error` 的"局部、有作用域、可嵌套"的写法。

---

## 1. 问题 / Problem

### 现状

```vb
On Error Resume Next
DoWork
If Err.Number <> 0 Then
    echo "失败: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
```

这套机制能用，但有三处结构性缺陷（这也是所有 VBS 程序员都骂过的）：

| 缺陷 | 后果 |
|---|---|
| **`On Error` 是过程级开关**，没有块作用域 | 一个 `On Error Resume Next` 会一直生效到过程结束，中间任何一行都被"吞掉" |
| **无法回到"出错点"之后** | `Resume Next` 只能跳过，做不了"重试 / 回滚 / 换方案" |
| **嵌套时没有栈** | 内层过程设了 `GoTo 0`，外层就再也接不到错误 |

### 目标

1. 提供**块作用域**的异常处理（进块开启、出块自动恢复）；
2. 支持 **`Finally`**（资源清理必须有）；
3. 支持**重抛**与**异常链**（`InnerException` 式的因果链）；
4. **与 `On Error` 共存**，且在同一个过程里混用时行为可预测；
5. **纯超集**：不加任何新语义到已有 VBS 代码上，老脚本一个字节不改。

---

## 2. 设计：三层，共用一套运行时

我的提议不是"加一个 try 语句"，而是把错误处理拆成**三层**，三层共享同一个
运行时错误栈 `_ovpp_err`：

```
第 3 层  Try / Catch / Finally        块作用域、可嵌套、可重抛   ← 新增语法
第 2 层  On Error Resume Next / GoTo 0  过程级开关（VBS 原生）      ← 保持原样
第 1 层  Err 对象 + 错误栈              运行时公共设施             ← 扩展
```

关键点：**第 2 层和第 3 层编译到同一套 `try/except` 结构**，只是"谁来接住"不同。
这样混用时可预测，实现也不必维护两套语义。

### 2.1 语法（带 VBS 味道的选择）

```vb
Try
    DoWork
Catch e
    echo "出错了: " & e.Number & " " & e.Description
    echo "来源: " & e.Source
Catch e When Err.Number = 53        ' 可选的过滤子句（VBS 风格的关键字 When）
    echo "文件没找到"
Finally
    CloseAll
End Try
```

设计取舍（为什么不是 C#/Java 的样子）：

| 选择 | 理由 |
|---|---|
| `Catch e` 而不是 `Catch (e)` | VBS 调用过程本来就不用括号 |
| `When` 过滤子句 | VBS 的 `Err.Number` 是数字，用数字过滤比类型匹配自然；`When` 是 VBS 已有的词（`Select Case` 之外没用过，不冲突） |
| 捕获对象就是 `Err` 的一个**快照** | 不引入新的异常类型体系，`e.Number/.Description/.Source` 和 `Err` 完全一致，学习成本为零 |
| `End Try` 而不是 `EndTry` | 和 `End If` / `End Sub` 对齐 |
| 没有 `Throw` 关键字，用 `Err.Raise` | **刻意不加新关键字**：`Err.Raise` 已经是 VBS 的重抛方式，`Throw x` 只是语法糖 |

### 2.2 与 `On Error` 的关系（最重要的一节）

**规则一句话：`On Error` 设置的是"过程默认行为"，`Try` 是"块内覆盖"。**

| 场景 | 行为 |
|---|---|
| 过程内 `On Error Resume Next`，块外 | 照旧：错误被捕获、执行继续 |
| 同上，进入 `Try` 块 | `Try` 生效期间，错误**不再被 Resume Next 吞掉**，而是交给 `Catch` |
| 离开 `Try`（正常或异常路径） | 恢复进入前的 `On Error` 状态（块作用域） |
| `Catch` 里没有匹配的 `When` | 继续向上：先给外层 `Try`，再给过程的 `On Error`，最后抛给调用者 |
| `Catch` 内再抛（`Err.Raise`） | 走同一条链，`Finally` 仍然执行 |
| 过程内没有 `Try`，只有 `On Error` | 和今天**完全一样**（零行为变化） |

实现上就是一个"错误处理栈"：`Try` 入栈时压入一个 handler，出栈时弹出；
`_ovpp_err.dispatch()` 查栈顶，栈空再走 `On Error` 的旧逻辑。旧逻辑不吃新路径，
所以超集性质自动成立。

### 2.3 `Finally` 的语义

- **一定会执行**：正常离开、`Exit Sub/Function`、`Exit Do/For`、错误离开、重抛；
- `Finally` 里抛出的错误**覆盖**正在传播的错误（与 .NET 一致），旧错误挂在
  `Err.InnerException` 上；
- `Finally` 里 `Exit Sub` 允许，但会跳过 `Finally` 之后的语句。

### 2.4 错误对象与链 / Error object and chain

沿用 `Err` 的四个字段，再加两个**只读**字段：

| 字段 | 含义 |
|---|---|
| `Err.Number` | 错误号（负数表示 COM 风格，正数表示 VBS 风格，沿用现状） |
| `Err.Description` | 描述 |
| `Err.Source` | 抛出位置（过程名，自动填） |
| `Err.Line` | **新增**：出错行号（现在只在终端打印，不给程序） |
| `Err.InnerException` | **新增**：上一个被覆盖的错误（形成链） |

`Err.Line` 是纯增量：VBS 程序不会读它，但调试器/LSP 和 `Catch` 都能用。

### 2.5 编译产物（示意）

```vb
Try
    DoWork
Catch e
    Handle e
End Try
```

编译成（简化）：

```python
_handler = _ovpp_err.push_handler()
_ovpp_err.begin_try(_handler)
try:
    _ovpp_sub_dowork()
except BaseException as _exc:
    if not _ovpp_err.handle(_handler, _exc):
        raise
    e = _ovpp_err.current_snapshot()
    _ovpp_sub_handle(e)
finally:
    _ovpp_err.pop_handler(_handler)
```

要点：`handle()` 负责 `When` 过滤与"没人要就继续往上抛"，`pop_handler` 恢复
`On Error` 状态——所以**块作用域是天然的**，不需要编译器做复杂的状态分析。

---

## 3. 为什么不用别的方案

| 方案 | 否决理由 |
|---|---|
| `On Error Resume Next` + 立即检查（现状） | 无块作用域、无 `Finally`、无重试 |
| 照抄 C#：`try/catch/finally` | 丢掉了 VBS 的 `Err` 心智模型；两套错误对象并存会让新人困惑 |
| 照抄 Python：`raise/try/except` | 需要新关键字 `raise`/`except`，且和 `Err.Raise` 功能重复 |
| 用 `Resume` 做重试（VBA 风格） | VBA 的 `Resume label` 需要行号/标签，本语言没有行号语句，成本高于收益 |
| 只加 `Catch` 不加 `Finally` | 资源清理没有归宿，反而容易写出泄漏 |

---

## 4. 超集兼容性论证

| 可能的重叠 | 结论 |
|---|---|
| `Try` / `Catch` / `Finally` / `When` 是不是 VBS 关键字？ | **不是**。它们是"上下文关键字"，只在语句开头被识别；现有 VBS 里 `Try = 1` 这种变量名仍然合法（我们的词法器本来就不保留关键字） |
| 老脚本里有没有 `Try:` 这样的标签？ | 没有（VBS 没有用户标签） |
| `End Try` 会不会和 `End` 语句冲突？ | 不会，`End` 后面接 `Try` 才成立，现有 `End` 无参数形式照旧 |
| 未使用 `Try` 的代码 | 零变化（新语法不参与编译） |

再补一条硬性要求：**`tests/test_vbs_superset.py` 里所有用例必须在加上这个特性后
继续通过**，且新增的 `Try` 用例单独分组，禁止改动既有断言。

---

## 5. 实施顺序（建议）

| 阶段 | 内容 | 风险 |
|---|---|---|
| P0 | 运行时：`push_handler/pop_handler/handle/snapshot` + `Err.Line` | 低（纯新增） |
| P1 | 语法：`Try … Catch e [When expr] … Finally … End Try`，单层 | 中（parser 块结构） |
| P2 | 嵌套、重抛、`InnerException` 链 | 中 |
| P3 | `Exit Sub` 穿越 `Finally`、`On Error` 混用矩阵测试 | 高（要穷举组合） |
| P4 | LSP/调试器接入（`Err.Line` + 栈帧） | 低 |

P3 的测试矩阵建议直接列成表：`{块内正常, 块内出错, Catch 内出错, 块内 ExitSub}` ×
`{Finally 有/无}` × `{外层 On Error 开/关}`，共 16 种，全部要有断言。

---

## 6. 开放问题（需要你拍板）

1. **`Catch` 是否允许多个无 `When` 的分支？** 我倾向：只允许一个通配 `Catch`，
   且必须在最后（和 `Select Case Else` 的位置规则一致）。
2. **`Finally` 里 `Exit Sub` 是否允许？** 我倾向允许，但要文档写清它会跳过
   `Finally` 剩余语句。
3. **要不要一个 `Try … Catch` 的简写形式**（单行 `Try stmt Catch: ...`）？
   我倾向**不做**，异常处理应该显式。
4. **错误号空间**：`Err.Raise` 的负数编号是否要和 `_ERROR_MAP` 对齐？
   我倾向保持现状，只补文档。


---

## 7. 实现落地记录（v0.8.0）

| 设计项 | 实现 | 说明 |
|---|---|---|
| 错误处理栈 | ✅ ovpprt/err.py 的 push_handler/pop_handler/handler_active | 支持嵌套 |
| Try 覆盖 On Error Resume Next | ✅ capture() 发现栈非空直接重新抛出 | 一行代码解决语义优先级 |
| Catch e 快照 | ✅ _Snapshot（Number/Description/Source/**Line**） | 与 Err 字段一一对应 |
| When 过滤 | ✅ 编译成 if/elif 链，未命中则 raise 继续向上 | When 里可以直接用 e.Number |
| Finally | ✅ Python try/finally + pop_handler() | Exit Sub 自动穿越 |
| 重抛 | ✅ 继续用 Err.Raise（不新增 Throw） | 与设计一致 |
| Err.Line | ✅ 从 traceback 找最深的非包内帧 | 生成代码带 VBS 行号，所以直接可用 |

**与设计的偏差**：无功能性偏差。设计里的四个开放问题按建议默认值实现——
只允许最后一个 Catch 省略 When（否则解析期报错）、Finally 里允许 Exit Sub、
不做单行简写、错误号保持现状。

**新增测试**：tests/test_try_catch.py（14 条），包含两条上下文关键字用例，
确保 Try/Catch/Finally 作为普通变量名时仍然符合 VBS 语义。
