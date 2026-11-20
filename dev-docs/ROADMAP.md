# Roadmap

> English | **[中文](ROADMAP.zh-CN.md)**

> Where the project stands and what comes next. Everything the maintainer has
> explicitly **deferred** (remote package sources, the IDE/LSP work) is collected
> at the very end, so it is never mistaken for a near-term commitment. COM
> support used to be deferred; since v0.9.14 it ships as an optional `over`
> package (`com-support`), bundled on Windows.

---

## 1. Done

| Version | Contents |
|---|---|
| 0.1–0.3 | Lexer, parser, AST mapping, control flow, procedures and ByRef, `On Error`, `With`, `Select Case`, VBS superset tests |
| 0.4 | Object model (`Class` / `Property` / `Me` / `Dictionary`), `Option Typing` static checking |
| 0.5 | `Execute` / `Eval`, `Format*`, `Option Compare`, `GetObject`, default members, private visibility |
| 0.6 | Method/procedure overloading, class inheritance via `Inherits`, `#python` blocks, `Option COP`, method ByRef |
| 0.7–0.8 | Typed array elements, `Try/Catch/Finally`, labels with `On Error GoTo`, labelled error handling |
| 0.9.0–0.9.5 | `Format()` patterns, date functions, `FileSystemObject`, interpolation plus `?.` / `??` / compound assignment, `AddressOf`, `RegExp`, `JSON`, multi-file modules, `Resume` state machine |
| 0.9.6–0.9.7 | `Shared`, `Implements`, `Property Set` vs `Let`, constructor arguments |
| 0.9.8–0.9.10 | `over` M1: install from a local source, lock file, package-name resolution; `CreateObject` back to VBS semantics |
| 0.9.11–0.9.13 | Reproducible lock file (`--frozen`, `commit` field), `over pack` / Git sources / local registry, `--json` diagnostics, web IDE |
| 0.9.14 | `over` M5/M6: `python` dependencies handed to pip (the real closure is recorded in the lock), `platforms` conditions, bundled `com-support` (real COM on Windows, no `import` in scripts) |
| 0.9.15 | Performance: lazy imports at startup, read-only `ByRef` no longer boxed (calls 13.7× → 1.35×), no error frame for procedures without `On Error` (recursion 12× → 6.9×), `For` compiles to `range` when safe (loops 1.65× → 0.99×) |
| 0.9.20 | Python interop becomes the `python-embed` over package with an explicit import (`Option COP` removed), `ast2str` deleted, docs rewritten English-first with `dev-docs/` split out |
| 0.9.21 | VBScript completeness pass: `&H` / `&O` / `#date#` literals, the 19 remaining built-ins (`Sin` … `GetRef`), `ExecuteGlobal`, the `TimeValue` / `Time()` date bug, and a `doc/VBS_SUPERSET.md` corrected against the implementation |

**Current status**: every native VBScript statement form is in place (only `Declare`
is declined by design); 557 tests green;
in-process loops, arrays and calls are now on par with hand-written Python
(0.99×–1.35×). On Windows COM is real through the `com-support` package — and
on non-Windows platforms, with `OPENVBS_COM=0`, or without `pywin32`, you get an
explicit error instead. See the note at the end of §3.

---

## 2. Next steps (suggested order)

### P0 · required before going public (about half a day each)

| # | Task | Why now |
|---|---|---|
| 1 | `.github/workflows/ci.yml` (3.10–3.13 × win/linux, with coverage) | ✅ done in v0.9.15 |
| 2 | Issue / PR templates (bug, feature, superset violation) | lowers the cost of outside contributions |
| 3 | `over pack` (turn a package directory into `.ovppkg`) | without it `over` can only install from directories, so nothing can be distributed |
| 4 | ~~`openvbs --json check/run`~~ | ✅ done in v0.9.12 |

### P1 · closing the semantics (1–2 days)

| # | Task | Notes |
|---|---|---|
| 5 | Connect `Implements` with `Option Typing` | `Dim s -> IShape` should accept any implementing class (today it reports a type mismatch) |
| 6 | `Option Strict` (optional) | forbids implicit conversions, a stricter tier for the type system |
| 7 | `over` M2: Git sources, conflict reporting, `over update` | from "local directory" to "shared by a team" |

### P2 · experience (as needed)

| # | Task | Notes |
|---|---|---|
| 8 | Compile cache (content hash) | saves parsing time on repeated runs and prepares incremental LSP |
| 9 | REPL (`openvbs repl`) | teaching scenarios want this badly |
| 10 | Formatter (built on the AST dump) | keeps style consistent across community contributions |

---

## 3. Explicitly deferred (maintainer's decision)

> The items below are **not planned for the near future**. They are listed so
> nobody mistakes them for a commitment.

| Item | Why deferred | Precondition |
|---|---|---|
| **IDE / LSP** | engineering work on the language itself is not finished | the P0 `--json` contract lands first |
| **Remote package registry / server** | the maintainer has no server; `over` stays offline-first | real distribution demand appears in the community |
| **Promotion / public launch** | the maintainer needs time to prepare | timing is the maintainer's call |

**Moved out of the deferred list**:

* **COM support** — landed in v0.9.14 as an `over` package (`extras/com-support`,
  wrapping `pywin32`). It did not break "no mandatory third-party dependencies,
  cross-platform": `pywin32` is an **optional** `python` dependency, needed only
  when a script actually talks to COM. Non-Windows platforms install nothing and
  `CreateObject` fails with a clear message.
* **A pure-Python `WScript.Shell` stand-in** — no longer worth writing: going
  through real COM is more accurate.

---

## 4. Acceptance rules (every new feature must pass)

1. **Superset**: a valid VBScript program must never change behaviour; extensions
   may only add syntax, and each needs a case in `tests/test_vbs_superset.py`.
2. **Tested**: at least three tests per feature (positive, boundary, error).
3. **Documented**: the matching section under `doc/` plus one row in the README
   capability table.
4. **Explain the alternative**: if a feature has a "we could live without it"
   argument, write it into §3 of this document.

---

## 5. Long-term goal

| Goal | Notes | Trigger condition |
|---|---|---|
| **Self-hosting** | compile OpenVBS++ with an OpenVBS++ compiler written in OpenVBS++, recursively, like a C compiler (the fixed point of a T-diagram) | language features mostly frozen + healthy community response + maintainer has spare capacity |

**Not doing it now, but writing the path down** so it does not become an empty
promise. When the time comes, three steps:

1. **Fix the bootstrap subset**: carve out of the current language the smallest
   set that can express a compiler, write it into `doc/BOOTSTRAP.md`, and agree
   that this subset **only grows, never changes** — that is what makes
   self-hosting converge. Smaller is better.
2. **Prototype compiler (written in `.vbs`)**: lexer, parser and Python-AST
   emission only; being able to compile half of `examples/` counts as success.
   Run it with the existing Python OpenVBS++ — this is `T1` in the T-diagram.
3. **Fixed-point check**: let the `.vbs` compiler compile itself into `stage1`,
   then compile `stage1` into `stage2`, and compare artefacts and behaviour
   (running the `tests/` subset with zero differences). Only then is
   self-hosting real.

**Why it stays long-term**: self-hosting turns "add one syntax feature" from
"edit the transformer" into "edit three implementations", which is negative
value while the language is still moving. Re-evaluate after the §2 engineering
list converges and the coverage matrix in `doc/VBS_SUPERSET.md` is stable.
