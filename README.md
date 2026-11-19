# OpenVBS++

![tests](https://img.shields.io/badge/tests-485%20passing-brightgreen)
![version](https://img.shields.io/badge/version-0.9.15-blue)
![license](https://img.shields.io/badge/license-Apache--2.0-blue)
![python](https://img.shields.io/badge/python-3.9%2B-yellow)
![deps](https://img.shields.io/badge/dependencies-none-success)
[![CI](https://github.com/MarfChen/OpenVBSpp/actions/workflows/ci.yml/badge.svg)](https://github.com/MarfChen/OpenVBSpp/actions/workflows/ci.yml)

**Run your existing `.vbs` scripts on a modern compiler — and get classes, try/catch,
packages and real COM while you are at it.**

OpenVBS++ is a source-compatible VBScript engine written in Python. It compiles VBScript
into standard Python bytecode and executes it with CPython, so existing scripts keep
running as-is while you can opt into a superset of modern syntax: `Class ... Inherits`,
`Try / Catch / Finally`, string interpolation, optional chaining, a dependency manager
(`over`) and real Windows COM automation through `pywin32`.

> 中文文档：[README.zh-CN.md](README.zh-CN.md)

> **Disclaimer.** OpenVBS++ is an independent implementation. It is not affiliated with,
> endorsed by, or sponsored by Microsoft. VBScript is a trademark of Microsoft
> Corporation. This repository contains no Microsoft source code or documentation; the
> language grammar was implemented from publicly documented behaviour. See
> [doc/LEGAL.md](doc/LEGAL.md) for the full self-audit.

---

## Table of contents

- [Install](#install)
- [Quick start](#quick-start)
- [Command line](#command-line)
- [Web IDE](#web-ide)
- [Editor and IDE support](#editor-and-ide-support)
- [What you get beyond plain VBScript](#what-you-get-beyond-plain-vbscript)
- [Python interop](#python-interop)
- [Performance](#performance)
- [Why not X?](#why-not-x)
- [Project structure](#project-structure)
- [Tests](#tests)
- [Status and known limitations](#status-and-known-limitations)
- [Contributing](#contributing)
- [Roadmap](#roadmap)
- [License, third parties, trademarks](#license-third-parties-trademarks)
- [Security model](#security-model)

---

## Install

| Goal | Command | Result |
|---|---|---|
| From PyPI | `pip install openvbspp` | CLI `openvbs` + `python -m OpenVBSPP` |
| Editable dev install | `pip install -e ".[all]"` | editable install + dev tooling |
| Windows COM support | `pip install "openvbspp[com]"` | adds `pywin32` |
| Single-file exe | `.\tools\build.ps1 -Target exe` | `dist\openvbs.exe`, no Python required |
| Wheel + sdist | `.\tools\build.ps1 -Target wheel` | `dist/*.whl`, `dist/*.tar.gz` |

The runtime has **no required third-party dependency**. `pywin32` is only needed when you
want COM objects, and the shipped `openvbs.exe` already bundles it.
Release process (TestPyPI → PyPI → exe) lives in [doc/RELEASE.md](doc/RELEASE.md).

## Quick start

```bash
.\venv\Scripts\python.exe cli.py examples\procedures.vbs   # from a source checkout
openvbs examples\fizzbuzz.vbs                              # after pip install
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
x = 1 : y = 2                 ' colon separates statements
PrintLine "factorial"
echo Fact(5)                  ' 120
```

## Command line

```
openvbs [script.vbs] [-v] [--save] [--save-dir DIR]
python -m OpenVBSPP <same arguments>     # equivalent
```

| Argument | Meaning |
|---|---|
| `script.vbs` | file to run; omitted → built-in demo |
| `-v, --verbose` | print token stream, AST and generated Python |
| `--save` | also write `.pyc` and `.ast.txt` |
| `--save-dir` | output directory, default `./__pycache__` |
| `check <script>` | compile only; `--json` for machine-readable diagnostics |
| `run <script>` | run; `--json` returns `{ok, output, error, duration_ms}` |
| `over <subcommand>` | package manager: `install` / `list` / `info <name>` / `remove <name> [--python]` / `com` … see [doc/OVER.md](doc/OVER.md) |
| `ide [--port N]` | start the web IDE (default `http://127.0.0.1:7860`) |

## Web IDE

```bash
openvbs ide                                   # http://127.0.0.1:7860
openvbs ide --port 8000 --workspace ./scripts
```

Edit, run (Ctrl+Enter), inspect output and diagnostics, save/load scripts — all in the
browser, with zero third-party dependencies. The package manager is real (`over`); the
IDE's own package endpoints are still a stub. See [doc/WEB_IDE.md](doc/WEB_IDE.md).

## Editor and IDE support

| Editor / environment | Today | How |
|---|---|---|
| Built-in web IDE | ✅ full | `openvbs ide` |
| Any editor (tasks, plugins, CI) | ✅ full | `openvbs check --json`, `openvbs run --json` |
| VS Code | ⚠️ syntax + snippets only | [extras/vscode-openvbspp](extras/vscode-openvbspp) (TextMate grammar, comment/bracket config, snippets); no IntelliSense yet |
| Neovim / Emacs / Sublime / JetBrains | ⚠️ VBScript highlighting + CLI diagnostics | point the editor at `openvbs check --json`; a language server is planned |
| Language server (LSP) | ❌ not implemented | design + milestone plan in [doc/IDE_LSP_DESIGN.md](doc/IDE_LSP_DESIGN.md) |
| Debugger (DAP) | ❌ not implemented | `Err.Line`, Try/Catch frames are the groundwork |

Machine-readable contract used by every integration:

```bash
openvbs check --json script.vbs   # {"ok": true|false, "diagnostics": [{line, column, message, source_line, ...}]}
openvbs run --json script.vbs     # {"ok": true|false, "output": "...", "error": "...", "duration_ms": 12.3}
```

The IDE service also exposes a versioned capabilities document
([OpenVBSPP/ide/service.py](OpenVBSPP/ide/service.py)), which a future LSP server will reuse.

## What you get beyond plain VBScript

| Area | Support |
|---|---|
| Statement separators | newline, `:`, trailing `_` continuation; `rem` / `'` comments; CRLF / CR / LF / BOM tolerant |
| Identifiers | case-insensitive |
| Declarations | `Dim a, b(10), c()`; `ReDim [Preserve]`; `Const`; `Option Explicit`; `Let` / `Set` |
| Data types | integers, floats, strings, booleans, `Empty` / `Null` / `Nothing`, 1-D and multi-dimensional arrays |
| Operators | `+ - * / \ Mod ^`, `&`, `= <> < > <= >=` (also `==`), `Is`, `And Or Not Xor Eqv Imp` |
| Conditionals | `If ... Then ... End If`, **single-line `If ... Then ... [Else ...]`**, `ElseIf`, `Select Case` (with `Is`, `To`, comma lists) |
| Loops | `For ... Next`, `For Each ... Next`, `Do [While\|Until] ... Loop [While\|Until]`, `While ... Wend`, `Exit Do/For` |
| Procedures | `Sub` / `Function`, `ByRef` (default) / `ByVal`, recursion, `Exit Sub/Function`, `Call` |
| Packages | `over`: local directory / `.ovppkg` / git+ sources, lock file, conflict reporting; declared Python dependencies installed through pip |
| Multi-file | `<script src="lib.vbs">`, `<namespace src="lib.vbs" alias="u">` |
| First-class procedures | `f = AddressOf Foo`; procedures become values |
| Regex / JSON | `New RegExp` (`Test`/`Replace`/`Execute`), `Json.Parse` / `Json.Stringify` |
| Sugar | interpolation `$"..."`, optional chaining `?.`, null-coalescing `??`, compound assignment `+= &=` |
| Structured exceptions | `Try / Catch e [When cond] / Finally / End Try`, coexisting with `On Error` |
| Error handling | `On Error Resume Next`, `On Error GoTo 0`, `Err.Number/.Description/.Source`, `Err.Clear`, `Err.Raise` |
| Objects | `Class ... End Class`, fields, `Sub`/`Function` methods, `Property Get/Let/Set`, `Me`, `Class_Initialize`/`Class_Terminate`, `New`, `CreateObject`, `Dictionary`, `With ... End With` |
| Type system | `Option Typing` + `->` annotations (off by default; zero behaviour change when off) |
| Inheritance | `Class Child Inherits Parent` (superset extension: overrides, inherited fields/properties) |
| Python interop | `#python ... #endpython` blocks, `Option COP` shared object pool |
| Dynamic execution | `Execute` / `Eval` (reuses the compiler, runs in global scope) |
| Misc | `Erase`, `Option Compare Text`, `FormatNumber` family, `vbCrLf`-style constants, ~70 built-ins |

### Procedures: `Sub` / `Function`

```vb
Sub Hello()          ' no parameters
    echo "hi"
End Sub

Function Add(a, b)   ' return value is carried by the function name
    Add = a + b
End Function

Hello                ' three call forms are accepted
Call Add(1, 2)
x = Add(1, 2)
```

- Procedures may be called before their definition — signatures are registered in a first pass.
- Parameters default to `ByRef` and are written back; `ByVal` is not.
- Since v0.9.15 by-reference boxes are only created for parameters a procedure actually
  assigns (or passes on to another procedure); read-only `ByRef` parameters cost nothing.

### Syntax sugar

```vb
a = 1 : b = 2 : echo a + b            ' colon-separated statements
If b = 2 Then echo "two"              ' single-line If
If b = 2 Then echo "a" : echo "b" Else echo "c"
total = 1 + _
        2                             ' line continuation
```

### Variables, arrays, `Option Explicit`

```vb
Option Explicit
Dim total, i               ' scalars, initial value Empty
Dim scores(4)              ' array: lower bound 0, upper bound 4 (5 elements)
Dim grid(1, 2)             ' 2-D array, access as grid(1, 2)
ReDim Preserve scores(5)   ' grow, keep existing values
Erase scores               ' reset elements to Empty
Const NAME = "openvbs"
```

`Empty` behaves like VBS: prints as an empty line, is `0` in arithmetic and `""` in `&`.
`Null` propagates (`Null + 1` and `Null = 0` are `Null`) — test with `IsNull()`.
`Option Explicit` is enforced **at compile time** with line number and source line.

### Classes and objects

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

`New Foo` and `CreateObject("Foo")` both construct script classes.
`CreateObject("Scripting.Dictionary")` provides `Add/Exists/Item/Remove/RemoveAll/Count/Keys/Items`,
and `d("k")`, `d("k") = v`, `For Each k In d` all work.
On Windows, COM objects (`WScript.Shell`, `Excel.Application`, WMI, …) connect for real —
properties, methods, collections and `For Each` included, with no `import` in your script
(provided by the `com-support` package; see [extras/com-support](extras/com-support)).
Details and limits: [doc/OBJECT_MODEL.md](doc/OBJECT_MODEL.md), [doc/VBS_SUPERSET.md](doc/VBS_SUPERSET.md).

### `Option Typing`: compile-time type checking

Off by default — when off, annotations are ignored and behaviour is plain VBScript.
When on, `->` annotations (not `:`, which statement separation already uses) are checked:

```vb
Option Typing

Dim count -> Integer
Function Twice(n -> Integer) -> Integer
    Twice = n * 2
End Function

echo Twice(count)
```

Assignment, arguments and return values are checked, plus class fields/properties when the
receiver type is statically known. `Variant`, `Empty`, `Null` and `Object` match anything,
so untyped code is never rejected:

```
Type mismatch in assignment to 'count': expected Integer, got String
Location: Line 12
Source: count = "oops"
```

### Error handling

```vb
On Error Resume Next
value = 10 / 0
If Err.Number <> 0 Then
    echo "failed: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
```

Each procedure gets its own error frame, so a callee's `On Error` never leaks to its
caller; `Err.Raise number, source, description` raises from script. Errors are reported
with line number and source line:

```
Syntax error (1002): Missing 'End If': unexpected end of file while reading the If ... End If block
Location: Line 2, Column 1
Source: echo 1
```

### Built-in functions

| Category | Functions |
|---|---|
| Strings | `Len Left Right Mid UCase LCase LTrim RTrim Trim StrReverse Space String Asc Chr InStr InStrRev Replace Split Join Filter StrComp Hex Oct` |
| Conversion | `CStr CInt CLng CDbl CBool CByte` |
| Math | `Abs Int Fix Sgn Sqr Round Rnd Randomize` |
| Arrays | `Array UBound LBound` |
| Predicates | `IsNumeric IsEmpty IsNull IsObject IsArray IsDate TypeName` |
| Date/time | `Now Date Time Timer Year Month Day Hour Minute Second Weekday DateAdd DateDiff` |
| Interaction | `MsgBox` (prints to stdout), `InputBox` (reads stdin) |

Constants: `vbCrLf vbCr vbLf vbTab vbNullString vbNullChar vbBack vbFormFeed vbVerticalTab vbTrue vbFalse`.

### Numbers and booleans

- Comparisons return **VBS booleans**: true is `-1`, false is `0`; `echo` prints `True` / `False`.
- Logical operators are **bitwise**, as in VBS: `5 And 3` → `1`, `Not 0` → `-1`,
  `(1 = 1) And (2 = 2)` → `True` (booleans stay booleans).
- `echo` and `&` format numbers VBS-style: `3.0` prints as `3`, `7 / 2` prints as `3.5`.
- `\` and `Mod` truncate toward zero: `-7 Mod 2` → `-1`; `^` is left-associative: `2 ^ 3 ^ 2` → `64`.

## Python interop

```python
from OpenVBSPP.runtime import vbs_run
vbs_run('''
Function Add(a, b)
    Add = a + b
End Function
echo Add(2, 3)
''', filename="demo.vbs")

from OpenVBSPP.compiler import vbs_compile
code_obj = vbs_compile("echo 1\n", filename="demo.vbs")
exec(code_obj, {})
```

`vbs_run` returns captured output/errors; `vbs_compile` returns a plain `types.CodeType`.
Inside a script, `#python ... #endpython` embeds Python and `Option COP` exposes a shared
object pool. See [doc/VBS_SUPERSET.md](doc/VBS_SUPERSET.md) §4.

## Performance

Short version: **2.5–3× faster end-to-end than native `cscript.exe`, with ~1.5× faster
startup; loops, arrays and function calls now run within 1.0–1.4× of hand-written
CPython; `Dictionary` member access and deep recursion are the remaining weak spots.**
Good for batch jobs, glue and ops automation; not for million-entry dictionary workloads
or deep recursion in a hot loop.

Reproduce (no arguments, no hard-coded paths):

```bash
python tools/bench.py          # in-process: compile once, run many
python tools/bench.py --e2e    # adds end-to-end runs, auto-detects cscript.exe on Windows
```

### End-to-end (includes interpreter startup — what users actually feel)

| Script | OpenVBS++ | cscript.exe | CPython |
|---|---:|---:|---:|
| `echo` one line | **114 ms** | 169 ms | 70 ms |
| 1M-iteration arithmetic loop | **142 ms** | 438 ms | 101 ms |
| `fib(24)` | **150 ms** | 394 ms | 76 ms |

### In-process (pure runtime cost, median)

| Case | OpenVBS++ | CPython | Ratio |
|---|---:|---:|---:|
| 1M loop, inside a procedure | 29.7 ms | 30.1 ms | **0.99×** |
| 1M loop, module level | 69.7 ms | 67.0 ms | 1.04× |
| 100k array read/write | 12.8 ms | 12.7 ms | 1.00× |
| 200k function calls | 23.1 ms | 17.1 ms | 1.35× |
| 20k string concatenations | 6.3 ms | 4.5 ms | 1.39× |
| `fib(24)` | 37.9 ms | 5.5 ms | 6.85× |
| 50k `Dictionary` inserts | 58.4 ms | 4.5 ms | 13.1× |

### What v0.9.15 optimized

| Change | Effect |
|---|---|
| Lazy stdlib imports at startup | cold start 27 ms → 13 ms |
| Read-only `ByRef` parameters no longer boxed | 200k calls 118 ms → 23 ms |
| No error frame for procedures without `On Error` | `fib(24)` 67 ms → 38 ms |
| `For` compiles to `range` when provably safe | 1M loop 50 ms → 29.7 ms |

Remaining, deliberately deferred: `Dictionary` dispatch, call inlining, native lists for `Dim a(n)`.

## Why not X?

Honest answers — this tool is not the right choice for everything.

**Why not plain VBScript / `cscript.exe`?**
If your scripts are stable, never need speed and you cannot install anything, stay on
`cscript` — it is preinstalled and is the reference implementation. Choose OpenVBS++ when
you want classes/try-catch/packages, 2–3× faster batch runs, or better diagnostics and
JSON output for tooling. Trade-off: an extra dependency (Python or the 8 MB exe) and an
implementation that is not the vendor runtime.

**Why not Python?**
For new code, Python is usually the better answer: bigger ecosystem, cross-platform,
better tooling. OpenVBS++ wins when you must keep `.vbs` sources and entry points
(scheduled tasks, logon scripts, customer installations, old build systems), or when the
people maintaining the scripts only know VBScript. It is a migration and execution layer,
not a Python replacement.

**Why not PowerShell?**
PowerShell is the modern Windows admin shell and has a far larger module ecosystem.
But many environments still restrict or sign PowerShell, `.ps1` chains break on
ExecutionPolicy, and enterprises have mountains of working `.vbs`. Rewriting them is risk
for no business value. If you are writing new automation for machines you control,
PowerShell may be the better pick; OpenVBS++ exists to run what you already have.

**Why not JavaScript / Node?**
You would add a runtime and an npm supply chain for a job that needs file, registry and
COM access — TS/JS gives little there. Node is the right call when you also need a web UI
or that ecosystem's libraries.

**Why not C# / .NET?**
Better for long-lived, typed, compiled applications and real services. It is far heavier
for glue scripts: project files, build chain, runtime footprint. Use .NET when the script
is turning into a product.

**Why not AutoHotkey / AutoIt?**
Those excel at UI automation, hotkeys and window control. OpenVBS++ targets file,
registry, WMI and COM automation with VBScript syntax and semantics.

**When you should not use OpenVBS++ at all**
Cross-platform requirements, CPU-bound heavy computation, long-running services, or
environments where installing anything is forbidden and the vendor runtime must be used.
Edge cases of VBScript semantics also exist — see [Status and known limitations](#status-and-known-limitations).

## Project structure

```
OpenVBSPP/
  lexer.py parser.py transformer.py compiler.py   # VBScript → Python AST → bytecode
  typecheck.py                                    # Option Typing checks
  runtime.py                                       # vbs_run / vbs_compile entry points
  over.py                                          # package manager (over.json / over.lock)
  cli.py ide/                                      # CLI and web IDE
  ovpprt/                                          # runtime library: ref, err, string, array,
                                                   # dictionary-compatible object model, COM bridge
  vendor/com-support/                              # Windows COM adapter, shipped with the package
  ast_nodes/ plugins/
examples/            # 23 runnable .vbs examples
tests/               # 485 unittest cases
tools/               # bench.py, build.ps1
extras/              # over package examples (web-ide, com-support), VS Code grammar
doc/                 # language spec, superset notes, object model, LSP design, legal audit
```

## Tests

```bash
.\venv\Scripts\python.exe -m unittest discover -s tests -v
python -m unittest discover -s tests          # 485 tests, ~2.2 s
```

| File | Covers |
|---|---|
| `tests/test_openvbspp.py` | example output, operators, boolean semantics, control flow, lexer, error locations, compiler and CLI |
| `tests/test_procedures.py` | `Sub`/`Function`, `ByRef`, recursion, `Dim`/arrays, `Option Explicit`, built-ins, `Empty` |
| `tests/test_objects.py` | `Class`, properties, `Me`, `Class_Initialize`, `Dictionary`, private members, default members |
| `tests/test_option_typing.py` | `Option Typing` checks (including array element types and inheritance chains) |
| `tests/test_interop.py` | `#python` blocks and the `Option COP` pool |
| `tests/test_vbs_superset.py` | **superset conformance**: 100+ legal VBScript snippets must run |
| `tests/test_over.py`, `test_over_registry.py` | `over` dependency graph, lock file, local sources, registry |
| `tests/test_over_python.py` | Python dependency merging, pinning, frozen restore, uninstall |
| `tests/test_com.py` | COM bridge lookup order, kill switch, error translation, adapter behaviour (no real COM needed) |
| `tests/test_perf_semantics.py` | semantics pinned by the v0.9.15 optimizations (`For`→`range`, `ByRef`, error frames, lazy imports) |

## Status and known limitations

Not implemented yet ([doc/VBS_SUPERSET.md](doc/VBS_SUPERSET.md) §4):

- the `Collection` built-in class (that is VBA, not VBScript)
- `Declare` (binding external DLL functions)

COM host objects **are** supported: real connections on Windows through `com-support`;
on other platforms, with `OPENVBS_COM=0`, or without `pywin32` you get an explicit error.

Other known limits:

- array indices are not range-checked (`a(-1)` behaves like Python's last element)
- `Exit For` inside a `Do` loop only exits the innermost loop
- `#python` blocks run in global scope and cannot touch procedure locals directly
- `Class_Terminate` depends on reference counting/GC; timing is not guaranteed

The full object-model contract is in [doc/OOP.md](doc/OOP.md).

## Contributing

The project needs help to survive: superset edge cases, translations, built-ins,
packaging, performance. Contributions use **DCO** (no CLA) — `git commit -s`.
See [CONTRIBUTING.md](CONTRIBUTING.md) and [doc/good-first-issues.md](doc/good-first-issues.md).

Package-manager docs: [doc/OVER.md](doc/OVER.md). Roadmap and deferred items:
[doc/ROADMAP.md](doc/ROADMAP.md). Legal self-audit: [doc/LEGAL.md](doc/LEGAL.md).

## Roadmap

- **v0.2.x** — procedures, `Dim`/`ReDim`/`Const`, `Option Explicit`, arrays, built-ins ✅
- **v0.3.x** — single-line `If`, `:`, `On Error`, `With`, `For Each`, bitwise logic, superset check ✅
- **v0.4.x** — object model, `Option Typing` static checks ✅
- **v0.5.x** — `Execute`/`Eval`, `FormatNumber`, `Option Compare Text`, `GetObject`, default members ✅
- **v0.6.x** — object-method `ByRef`, `Inherits`, `#python`, `Option COP` ✅
- **v0.7.x–v0.9.x** — `over` package manager (M1/M2), real COM support, web IDE, `--json` diagnostics ✅
- **v0.9.15** — performance: lazy startup, no boxing for read-only `ByRef`, error frames on demand, `For` → `range` ✅
- **Next** — method overloading, remaining `Attribute` properties, events, `Collection`
- **v1.0** — stable language spec, official `ovpprt` runtime library, IDE / LSP support

## License, third parties, trademarks

Apache License 2.0 — see [LICENSE](LICENSE); keep [NOTICE](NOTICE) on redistribution.

- The repository, wheel and sdist contain **no third-party code**. `pywin32` is only an
  optional Windows runtime dependency (`pip install "openvbspp[com]"`, or installed
  automatically by `over` from the `com-support` package's `python: pywin32>=306`).
- The single-file **exe** published in GitHub Releases bundles `pywin32` (PSF-2.0) modules
  and DLLs so it works out of the box; its license statement is in [NOTICE](NOTICE).
  Use the wheel plus your own `pywin32` install if you want to avoid third-party code.
- VBScript is a trademark of Microsoft Corporation. OpenVBS++ is an independent
  implementation, not affiliated with Microsoft, and distributes neither its code nor its
  documentation.

## Security model

- `<% %>` embedded Python blocks (and `over` installs / `pip` calls) mean **arbitrary code
  execution** — the same trust model as `cscript.exe`: never run scripts or packages of
  unknown origin.
- On non-Windows platforms there is no COM: `CreateObject` fails explicitly instead of
  silently degrading.
- Report vulnerabilities per [SECURITY.md](SECURITY.md).
