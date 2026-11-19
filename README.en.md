# OpenVBS++

[![tests](https://img.shields.io/badge/tests-457%20passing-brightgreen)](tests)
[![version](https://img.shields.io/badge/version-0.9.14-blue)](doc/RELEASE.md)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue)](LICENSE)
[![python](https://img.shields.io/badge/python-3.9%2B-yellow)](pyproject.toml)
[![CI](https://github.com/MarfChen/OpenVBSpp/actions/workflows/ci.yml/badge.svg)](https://github.com/MarfChen/OpenVBSpp/actions/workflows/ci.yml)

**Run your legacy `.vbs` files — and write new scripts with a modern syntax, without leaving VBScript.**

OpenVBS++ compiles VBScript-flavoured source into the standard Python AST and executes it on CPython.
The one rule the project never breaks: **it is a superset of VBScript.** Every legal VBS program
keeps working; extensions may only *add* syntax.

> 中文文档见 [README.md](README.md)（更完整）。This file is a condensed English overview.

Why it exists: Microsoft is retiring VBScript (from Windows 11 24H2 it ships on demand), but millions of
scripts still run in production. OpenVBS++ keeps them runnable **and** gives you a path forward — real
classes, `Try/Catch`, a package manager, a type checker, and genuine COM interop.

## Install

```bash
pip install openvbspp            # Python 3.9+
pip install "openvbspp[com]"     # + pywin32, for CreateObject on Windows
```

Or grab the single-file `openvbs.exe` from Releases — it bundles the runtime *and* pywin32, so COM works
on a machine that has never seen Python.

## Quick start

```vbscript
Option Explicit

Class Counter
    Public Total
    Private Sub Class_Initialize()
        Total = 0
    End Sub
    Public Sub Add(n)
        Total = Total + n
    End Sub
End Class

Dim c
Set c = New Counter
c.Add 40
c.Add 2
echo c.Total          ' 42
```

```bash
openvbs script.vbs              # compile + run
openvbs check script.vbs        # static check only (lint)
openvbs check --json script.vbs # machine-readable diagnostics
openvbs run --json script.vbs   # {ok, output, duration_ms}
openvbs ide                     # local web IDE on 127.0.0.1
```

## What you get beyond plain VBScript

| | |
|---|---|
| **Classes** | `Class` / `New` / `Me`, `Class_Initialize`, `Property Get/Let/Set` |
| **Modern control flow** | `Try/Catch/Finally`, `Continue For`, `Exit Sub`, ternary-free but readable |
| **Type checker** | `Option Typing` — compile-time checks for `Dim x As Integer` |
| **Package manager** | `over` — local dirs, `.ovppkg`, git, plus **automatic `pip` resolution** for Python deps |
| **Real COM** | `CreateObject("WScript.Shell")`, Excel, WMI — no `import` needed, works from the exe |
| **Web IDE** | `openvbs ide` → editor with run/check endpoints |
| **Modules** | `<script src="lib.vbs">` / `<namespace src="lib.vbs" alias="u">` |

There is no `import` statement in the language, by design — `.vbs` uses the HTML-style tags above, and
`over` package names can be used directly: `<namespace src="text-utils" alias="tu">`.

## COM interop, zero configuration

On Windows, `CreateObject` just works:

```vbscript
Dim sh, env
Set sh = CreateObject("WScript.Shell")
env = sh.ExpandEnvironmentStrings("%OS%")
echo env                      ' Windows_NT
```

COM support ships as an `over` package (`com-support`) that is **bundled with the compiler**, so there is
nothing to install and nothing to import. A project can pin its own copy under `packages/com-support/`;
`openvbs over list` shows where it comes from, `openvbs over remove com-support` falls back to the bundled
copy, and `OPENVBS_COM=0` disables it entirely. On non-Windows platforms you get an explicit error, never a
silent downgrade.

## Packages

```bash
openvbs over init my-app            # create over.json
openvbs over add text-utils file:../libs/text-utils
openvbs over install                # resolves the whole graph; Python deps go to pip in one shot
openvbs over list                   # over packages + python packages + where COM support comes from
openvbs over info text-utils
openvbs over install --frozen       # reproducible install from over.lock
```

`over.json` may declare `"python": {"pywin32": ">=306"}`. The installer collects Python requirements
across the **entire** dependency graph, merges constraints, runs `pip install` once, and records the
*actual* resolved versions (transitive deps included) in `over.lock` v2.

## Performance

Measured with `python tools/bench.py` (reproducible, no hard-coded paths):

- **End to end it beats native `cscript.exe` by 2–2.6×** (startup 118 ms vs 174 ms; a 1M-iteration loop
  170 ms vs 443 ms; `fib(24)` 186 ms vs 396 ms).
- In-process, arithmetic loops are **1.7–2.1×** slower than equivalent hand-written Python;
  call-heavy and recursive code is **6–12×** slower, because every call honours `ByRef` parameters and
  `On Error Resume Next` scopes.
- Verdict: great for batch/glue/ops automation, **not** for CPU-bound number crunching. See the
  [performance section in README.md](README.md#性能) for the full table and the reasons.

## Status and limits

- 457 tests pass on Windows (Python 3.9–3.13); CI covers Windows and Linux.
- Not implemented: `Declare` (P/Invoke) and the `Collection` built-in class.
- `Option Explicit`, error codes, `Err`, number/boolean coercion, arrays, `With`, `Property`, classes,
  and COM all follow VBScript semantics, including the odd ones.

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) first — the superset rule is not negotiable. Beginner-friendly
tasks live in [doc/good-first-issues.md](doc/good-first-issues.md). Bug reports without a minimal
reproducible script are hard to act on, so please include one.

## License

Apache-2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE) for third-party components.

OpenVBS++ is an independent implementation. It is not affiliated with Microsoft; VBScript is a
Microsoft trademark. This repository contains no Microsoft code or documentation.

> **Security**: OpenVBS++ executes scripts with your privileges, exactly like `cscript.exe`. `<% %>`
> blocks embed Python, and `over install` invokes `pip` — both are arbitrary code execution. Never run
> scripts or `over.json` files from untrusted sources.
