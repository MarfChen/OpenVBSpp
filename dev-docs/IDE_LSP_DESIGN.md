# IDE / VS Code compatibility design report

> English | **[中文](IDE_LSP_DESIGN.zh-CN.md)**

> **Abstract** — Editors do not need our source code; they need three stable
> machine-readable things: **diagnostics**, **symbols with positions**, and
> **a run interface**. This report maps what OpenVBS++ already has onto the
> Language Server Protocol (LSP) and the Debug Adapter Protocol (DAP), lists the
> concrete gaps (spans, symbol export, incremental parsing), and proposes a
> milestone plan where M1 is achievable in a day.

---

## 0. Status (v0.9.20 — what actually shipped)

| Capability | Status | Notes |
|---|---|---|
| Built-in web IDE | ✅ shipped | `openvbs ide`: edit / run / check / save, zero third-party dependencies |
| Machine-readable diagnostics | ✅ shipped | `openvbs check --json` → `{ok, diagnostics:[{kind, number, message, line, column, source_line}]}` |
| Machine-readable run | ✅ shipped | `openvbs run --json` → `{ok, output, error, duration_ms}` |
| Capabilities document | ✅ shipped | `OpenVBSPP/ide/service.py` exports a versioned capability declaration (`"lsp": false` is an explicit placeholder) |
| VS Code highlighting | ✅ shipped (syntax / snippets / run / diagnostics) | [extras/vscode-openvbspp](../extras/vscode-openvbspp): TextMate grammar, comment and indent rules, snippets, `Ctrl+F5` to run, diagnostics on save; **no completion or go-to-definition** |
| Other editors | ⚠️ via the CLI | Neovim / Emacs / Sublime / JetBrains plug `openvbs check --json` into their "external command" or LSP client; VBScript highlighting is usually built in |
| LSP server | ❌ not implemented | this is M1 in §4 |
| DAP debugger | ❌ not implemented | M4 |

In other words: **M0 (highlighting) and the diagnostic/run export are done**; what
remains is the actual LSP transport (`Content-Length` header plus JSON-RPC over
stdio) and exporting the completion / go-to-definition data.

---

## 1. Conclusions first

| Capability | Today | Gap | Target milestone |
|---|---|---|---|
| Diagnostics (red squiggles) | ✅ line numbers plus the source line | column numbers incomplete, no `span` | **M1** |
| Run / stop | ✅ `vbs_run` | no machine-readable result | **M1** |
| Syntax highlighting | ✅ TextMate grammar already (`extras/vscode-openvbspp`) | — | M0 ✅ |
| Completion / hover / go-to | the underlying data exists | not exported as an API | **M2** |
| Rename / outline / folding | partial | no symbol tree | M3 |
| Formatting | ❌ | needs an AST-based emitter | M3 |
| Breakpoint debugging | ❌ | needs DAP plus a line map | M4 |

**Core judgement: we do not need to rewrite the compiler, only to export the
information we already have in a standard shape.** `parser` already produces
tokens as `(kind, value, line, col)`, `transformer._collect` already builds the
procedure signature table, and `typecheck` already has scopes and inferred
types — exactly what LSP wants.

---

## 2. Assets we can reuse right now

| Asset | Location | LSP use |
|---|---|---|
| Tokens with line/column | `lexer.py` (`(kind, value, line, col)`) | semantic highlighting, precise diagnostics |
| AST nodes with `.at(line, col)` | `ast_nodes/_base.py` | go-to, selection ranges |
| Procedure signatures `_ProcInfo` | `transformer._collect` | completion, signature help |
| Type environment `_Scope` / `TypeInfo` | `typecheck.py` | hover types, type errors |
| Class / field / property tables | `typecheck.ClassInfo` | member completion |
| Structured errors | `exceptions.OVPPError` | diagnostics |
| Line mapping | generated code carries `.lineno` | DAP stack frames |

**Conclusion: M1/M2 are an export layer, not a compiler rewrite.**

---

## 3. Three things to add

### 3.1 Position information: from "line" to "span"

The AST currently guarantees only `line` (some nodes have `col`). Every LSP
request needs a `range`:

```python
@dataclass
class Span:
    line: int        # 1-based
    col: int         # 1-based
    end_line: int
    end_col: int
```

Order of work (by value for effort): `Var` / `Call` / `Member` (completion and
go-to) → `Assign` → `SubDef/FunctionDef/ClassDef` (outline) → the rest.

### 3.2 Symbol export: `analyze()`

A new read-only API that **does not touch the compiler internals**:

```python
from OpenVBSPP.analysis import analyze

result = analyze(source, filename="demo.vbs")
result.diagnostics   # list[Diagnostic]
result.symbols       # list[Symbol] (procedures/classes/fields/properties/constants, with span and signature)
result.types         # {name: inferred_type}
result.scopes        # scope tree (for completion)
```

Internally `analyze()` is just "parse → typecheck → collect"; it does not emit
Python code, so it is **an order of magnitude faster than compiling** and fits a
per-keystroke trigger (within 100 ms of didChange).

### 3.3 Machine-readable CLI

```bash
openvbs --json run demo.vbs     # {"ok":true,"output":"...","error":null,"duration_ms":3.1}
openvbs --json check demo.vbs   # {"ok":false,"diagnostics":[{...}]}
```

This one is **already implemented** (v0.9.12): `openvbs --json check|run <script>`
emits `{ok, diagnostics}` / `{ok, output, error, duration_ms}` and the 0/1 exit
code is directly usable in CI.

---

## 4. LSP service design

### 4.1 Process model

```
VSCode ──stdin/stdout──> openvbs-lsp ──> analysis.analyze()
                                  └────> runtime.vbs_run()  (run command only)
```

- A single file, `OpenVBSPP/lsp/server.py`, JSON-RPC over stdio;
- no `pygls` (the LSP message layer is a `Content-Length` header plus JSON,
  about 100 lines by hand);
- it can also be an optional dependency, `pip install openvbspp[lsp]`; a
  hand-written implementation needs nothing.

### 4.2 Capability matrix (what `initialize` returns)

| LSP capability | Implementation | Milestone |
|---|---|---|
| `textDocument/publishDiagnostics` | `analyze().diagnostics` | M1 |
| `textDocumentSync` | `Full(1)` first, `Incremental(2)` later | M1 |
| `completionProvider` | keywords + scope variables + class members | M2 |
| `hoverProvider` | inferred type + signature + doc comment | M2 |
| `definitionProvider` | symbol table reverse lookup | M2 |
| `documentSymbolProvider` | procedure / class / field tree | M3 |
| `documentFormattingProvider` | a `ast.unparse`-style VBS emitter | M3 |
| `renameProvider` | symbol table + reference collection | M3 |
| `semanticTokensProvider` | token stream → typed highlighting | M3 |
| `signatureHelpProvider` | `_ProcInfo.params` | M2 |
| `foldingRangeProvider` | block structure | M3 |

### 4.3 Incremental parsing (after M2)

Every `didChange` currently re-parses the whole file. Measured: small scripts
take < 5 ms, so a full re-parse is fine; consider incremental only above 2000
lines (line cache plus re-parsing just the changed procedure bodies). **Do not
do incremental yet** — that is a classic premature optimisation.

---

## 5. What the editor side needs

### 5.1 VS Code extension (`vscode-openvbspp`)

| Part | Content | Depends on |
|---|---|---|
| Syntax highlighting | `syntaxes/openvbspp.tmLanguage.json` (TextMate) | nothing |
| Language configuration | `language-configuration.json` (comments, brackets, indentation) | nothing |
| LSP client | `vscode-languageclient`, launching `openvbs-lsp` | LSP |
| Run command | `openvbs --json run` → output panel | nothing |
| Tasks | `tasks.json` templates (run / check / build exe) | nothing |
| Snippets | `snippets/openvbspp.json` (If/For/Class/Sub templates) | nothing |
| Debugging | DAP client → `openvbs-dap` | M4 |

**Key point: the TextMate grammar and `language-configuration.json` are write
once, reuse everywhere** — Neovim, Emacs, Sublime and JetBrains all consume the
same file, so doing them first pays off the most.

### 5.2 Other editors

| Editor | How to connect |
|---|---|
| Neovim | built-in LSP client, `vim.lsp.start` pointing at `openvbs-lsp` |
| Emacs | `eglot` / `lsp-mode` (both standard LSP clients) |
| Sublime | LSP package |
| JetBrains | the official LSP plugin |
| Vim / others | use `openvbs --json check` with `errorformat` first |

---

## 6. DAP (breakpoint debugging) design notes

The hard part of debugging is not the protocol, it is the **line mapping**:

- every generated node carries `.lineno` (the VBS line) already;
- breakpoints: insert `_ovpp_debug.hit(line)` before the first generated
  statement of a target VBS line, and let it handshake with the debugger (filter
  `sys.settrace`'s `line` events down to `_ovpp_debug.hit`);
- stack frames: map Python frames back to `(procedure name, line)`, reverse
  looking the name up from the function name (`_ovpp_fn_foo` → `Foo`);
- variables panel: show the locals of the current scope, expanding `Ref` values
  to `.v`.

No compiler change: add a "debug mode" switch in `transformer` that inserts
`hit()` calls.

---

## 7. Milestones and effort

| Milestone | Content | Estimate | Depends on |
|---|---|---|---|
| **M0** | TextMate grammar + `openvbs --json check` | 1 day | nothing |
| **M1** | `analysis.analyze()` + LSP service (diagnostics + sync) | 2–3 days | nothing |
| **M2** | completion / hover / go-to / signature help | 3–5 days | M1 |
| **M3** | outline / rename / semantic highlighting / formatting | 1 week | M2 |
| **M4** | DAP debugging | 1 week | M1 |

**Suggested order: M0 → M1 → M2.** M0 takes a day and already makes an editor
"look like it knows the language", which is the cheapest visible win for
promotion material.

---

## 8. Interfaces to reserve now

These are small changes that avoid rework later:

1. Add `column` and `end_column` to `OVPPError` (today it only has `line_number`);
2. make `ast_nodes`' `.at()` store `(line, col, end_line, end_col)` uniformly;
3. add `OpenVBSPP/analysis.py` (`analyze()`): read-only, no code generation;
4. pin the CLI `--json` output (schema in §3.3) and document it in `dev-docs/RELEASE.md`;
5. versioned `capabilities()` (the IDE already consumes it); LSP will reuse the
   same declaration later.

With items 3 and 4 done, **any editor can plug in through "external command +
JSON"** without waiting for LSP.
