# IDE / VSCode 兼容设计报告

> **English abstract** — Editors do not need our source code; they need three
> stable machine-readable things: **diagnostics**, **symbols with positions**,
> and **a run interface**. This report maps what OpenVBS++ already has onto the
> Language Server Protocol (LSP) and the Debug Adapter Protocol (DAP), lists the
> concrete gaps (spans, symbol export, incremental parsing), and proposes a
> milestone plan where M1 is achievable in a day.

---

## 0. 现状（v0.9.20，实际交付了什么）

| 能力 | 状态 | 说明 |
|---|---|---|
| 内置网页 IDE | ✅ 已交付 | `openvbs ide`：编辑 / 运行 / 检查 / 保存，零第三方依赖 |
| 机器可读诊断 | ✅ 已交付 | `openvbs check --json` → `{ok, diagnostics:[{kind, number, message, line, column, source_line}]}` |
| 机器可读运行 | ✅ 已交付 | `openvbs run --json` → `{ok, output, error, duration_ms}` |
| capabilities 文档 | ✅ 已交付 | `OpenVBSPP/ide/service.py` 导出带版本号的能力声明（`"lsp": false` 是显式占位） |
| VS Code 语法高亮 | ✅ 已交付（语法 / 片段 / 运行 / 诊断） | [extras/vscode-openvbspp](../extras/vscode-openvbspp)：TextMate 语法、注释与缩进规则、snippets、`Ctrl+F5` 运行、保存时诊断；**没有补全/跳转** |
| 其它编辑器 | ⚠️ 靠 CLI | Neovim / Emacs / Sublime / JetBrains 用各自的 "外部命令" 或 LSP 客户端接 `openvbs check --json`；VBScript 高亮大多内置 |
| LSP 服务器 | ❌ 未实现 | 就是下面 §4 的 M1 |
| DAP 调试器 | ❌ 未实现 | M4 |

也就是说：**M0（语法高亮）与"诊断/运行导出"已完成**，剩下的是真正的 LSP 传输层
（`Content-Length` 头 + JSON-RPC over stdio）和补全/跳转的数据导出。

---

## 1. 结论先行

| 能力 | 现在 | 差距 | 目标里程碑 |
|---|---|---|---|
| 诊断（红波浪线） | ✅ 有行号 + 源码行 | 列号不完整、缺 `span` | **M1** |
| 运行 / 停止 | ✅ `vbs_run` | 缺机器的可读结果 | **M1** |
| 语法高亮 | ✅ 已有 TextMate 语法（`extras/vscode-openvbspp`） | — | M0 ✅ |
| 补全 / 悬停 / 跳转 | 底层数据都有 | 未导出为 API | **M2** |
| 重命名 / 大纲 / 折叠 | 部分 | 缺符号树 | M3 |
| 格式化 | ❌ | 需要基于 AST 的输出器 | M3 |
| 断点调试 | ❌ | 需要 DAP + 行号映射 | M4 |

**核心判断：我们不需要重写编译器，只需要把已有信息"导出成标准形状"。**
`parser` 已经有 `(kind, value, line, col)` 的 token，`transformer` 的 `_collect`
已经建好了过程签名表，`typecheck` 已经有作用域与推断类型——这些正是 LSP 要的东西。

---

## 2. 现在就能复用的资产

| 资产 | 位置 | LSP 用途 |
|---|---|---|
| token 带行列 | `lexer.py`（`(kind, value, line, col)`） | 语义高亮、精确诊断 |
| AST 节点 `.at(line, col)` | `ast_nodes/_base.py` | 跳转、选中范围 |
| 过程签名表 `_ProcInfo` | `transformer._collect` | 补全、签名帮助 |
| 类型环境 `_Scope` / `TypeInfo` | `typecheck.py` | 悬停类型、类型错误 |
| 类/字段/属性表 | `typecheck.ClassInfo` | 成员补全 |
| 结构化错误 | `exceptions.OVPPError` | 诊断 |
| 行号映射 | 生成代码带 `.lineno` | DAP 栈帧回转 |

**结论：M1/M2 是"导出层"工作，不是编译器改造。**

---

## 3. 需要补的三件事

### 3.1 位置信息：从 "line" 到 "span"

现在 AST 只保证 `line`（部分有 `col`）。LSP 的每一个请求都要 `range`：

```python
@dataclass
class Span:
    line: int        # 1-based
    col: int         # 1-based
    end_line: int
    end_col: int
```

改造顺序（按性价比）：`Var` / `Call` / `Member`（补全与跳转用）→ `Assign` →
`SubDef/FunctionDef/ClassDef`（大纲用）→ 其余。

### 3.2 符号导出：`analyze()`

新增一个只读 API，**不动编译器内部**：

```python
from OpenVBSPP.analysis import analyze

result = analyze(source, filename="demo.vbs")
result.diagnostics   # list[Diagnostic]
result.symbols       # list[Symbol]（过程/类/字段/属性/常量，含 span 与签名）
result.types         # {name: inferred_type}
result.scopes        # 作用域树（补全用）
```

`analyze()` 内部就是"parse → typecheck → 收集器"，不生成 Python 代码，因此
**比编译快一个数量级**，适合按键触发（didChange 后 100ms 内）。

### 3.3 机器可读的 CLI

```bash
openvbs --json run demo.vbs     # {"ok":true,"output":"...","error":null,"duration_ms":3.1}
openvbs --json check demo.vbs   # {"ok":false,"diagnostics":[{...}]}
```

这条**已实现**（v0.9.12）：`openvbs --json check|run <script>` 输出
`{ok, diagnostics}` / `{ok, output, error, duration_ms}`，退出码 0/1 可直接给 CI 用。

---

## 4. LSP 服务设计

### 4.1 进程模型

```
VSCode ──stdin/stdout──> openvbs-lsp ──> analysis.analyze()
                                  └────> runtime.vbs_run()（仅 run 命令）
```

- 单文件 `OpenVBSPP/lsp/server.py`，JSON-RPC over stdio；
- 不引入 `pygls`（LSP 的消息层就是一个 `Content-Length` 头 + JSON，手写 100 行）；
- 也可以做成可选依赖 `pip install openvbspp[lsp]`，用手写实现则不需要。

### 4.2 能力矩阵（initialize 返回）

| LSP 能力 | 实现方式 | 里程碑 |
|---|---|---|
| `textDocument/publishDiagnostics` | `analyze().diagnostics` | M1 |
| `textDocumentSync` | 先 `Full(1)`，之后 `Incremental(2)` | M1 |
| `completionProvider` | 关键字 + 作用域变量 + 类成员 | M2 |
| `hoverProvider` | 推断类型 + 签名 + 文档注释 | M2 |
| `definitionProvider` | 符号表反查 | M2 |
| `documentSymbolProvider` | 过程/类/字段树 | M3 |
| `documentFormattingProvider` | Python `ast.unparse` 式的 VBS 输出器 | M3 |
| `renameProvider` | 符号表 + 引用收集 | M3 |
| `semanticTokensProvider` | token 流 → 类型化高亮 | M3 |
| `signatureHelpProvider` | `_ProcInfo.params` | M2 |
| `foldingRangeProvider` | 块结构 | M3 |

### 4.3 增量解析（M2 之后的事）

现在每次 `didChange` 全量重解析。实测小脚本 < 5ms，可以直接全量；超过 2000 行
再考虑增量（行级缓存 + 只重解析改动的过程体）。**先不要做增量**，这是典型的过早优化。

---

## 5. 编辑器侧要做什么

### 5.1 VSCode 扩展（`vscode-openvbspp`）

| 部件 | 内容 | 依赖 |
|---|---|---|
| 语法高亮 | `syntaxes/openvbspp.tmLanguage.json`（TextMate） | 无 |
| 语言配置 | `language-configuration.json`（注释、括号、缩进） | 无 |
| LSP 客户端 | `vscode-languageclient`，启动 `openvbs-lsp` | LSP |
| 运行命令 | `openvbs --json run` → 输出面板 | 无 |
| 任务 | `tasks.json` 模板（run / check / build exe） | 无 |
| 片段 | `snippets/openvbspp.json`（If/For/Class/Sub 模板） | 无 |
| 调试 | DAP 客户端 → `openvbs-dap` | M4 |

**关键点：TextMate 语法和 `language-configuration.json` 是"一次编写，到处复用"的**
——Neovim、Emacs、Sublime、JetBrains 都用同一份，所以先做它，收益最大。

### 5.2 其它编辑器

| 编辑器 | 接入方式 |
|---|---|
| Neovim | 内置 LSP 客户端，`vim.lsp.start` 指向 `openvbs-lsp` |
| Emacs | `eglot` / `lsp-mode`（都是标准 LSP 客户端） |
| Sublime | LSP 插件 |
| JetBrains | 官方 LSP 插件 |
| Vim/其它 | 先用 `openvbs --json check` 做 `errorformat` |

---

## 6. DAP（断点调试）的设计要点

调试真正的难点不是协议，是**行号映射**：

- 生成代码的每个节点都带 `.lineno`（VBS 行号），已有；
- 断点：在目标 VBS 行的第一条生成语句前插入 `_ovpp_debug.hit(line)`，
  由它和调试器握手（`sys.settrace` 的 `line` 事件过滤到 `_ovpp_debug.hit`）；
- 栈帧：把 Python 帧映射回 `(过程名, 行号)`，过程名从函数名反查
  （`_ovpp_fn_foo` → `Foo`）；
- 变量面板：显示当前作用域的局部变量，`Ref` 类型要展开成 `.v`。

不需要改编译器，只需要在 `transformer` 里加一个"调试模式"开关插入 `hit()` 调用。

---

## 7. 里程碑与工作量

| 里程碑 | 内容 | 预估 | 依赖 |
|---|---|---|---|
| **M0** | TextMate 语法 + `openvbs --json check` | 1 天 | 无 |
| **M1** | `analysis.analyze()` + LSP 服务（诊断 + 同步） | 2–3 天 | 无 |
| **M2** | 补全 / 悬停 / 跳转 / 签名帮助 | 3–5 天 | M1 |
| **M3** | 大纲 / 重命名 / 语义高亮 / 格式化 | 1 周 | M2 |
| **M4** | DAP 调试 | 1 周 | M1 |

**建议顺序：M0 → M1 → M2。** M0 一天就能让编辑器"看起来像认得这门语言"，
是宣传素材里最容易出效果的一步。

---

## 8. 现在就要预留的接口

这些改动很小，但能避免以后返工：

1. `OVPPError` 增加 `column` 与 `end_column` 字段（现在只有 `line_number`）；
2. `ast_nodes` 的 `.at()` 统一存 `(line, col, end_line, end_col)`；
3. 新增 `OpenVBSPP/analysis.py`（`analyze()`），只读、不生成代码；
4. CLI 固化 `--json` 输出（schema 见 §3.3），并补进 `dev-docs/RELEASE.md`；
5. 版本化的 `capabilities()`（IDE 已经在用），未来 LSP 复用同一份声明。

第 3、4 条做完，**任何编辑器都能用"外部命令 + JSON"接进来**，不必等 LSP。
