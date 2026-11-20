# Changelog

All notable changes to **OpenVBS++** are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Release-by-release engineering notes (how to cut a release) live in
[`dev-docs/RELEASE.md`](dev-docs/RELEASE.md); the roadmap is in [`dev-docs/ROADMAP.md`](dev-docs/ROADMAP.md).

## [0.9.27] — 2026-11-21

### Fixed

- **`Scripting.FileSystemObject` path helpers keep Windows semantics on every
  host**: `GetFileName`, `GetBaseName` and `GetExtensionName` used the host's own
  path flavour, so `"C:\data\r.txt"` came back as the whole string on Linux.
  They now pick `ntpath` for Windows-style paths and `os.path` otherwise, and
  `GetParentFolderName` follows the same rule.
- **`over list` / `over info com-support` report the bundled directory off
  Windows** instead of an empty path, while still telling you the bridge cannot
  load there — the platform gate now only blocks loading, not bookkeeping.
- **The COM suite only runs where COM exists**: provider-dependent cases are
  skipped on non-Windows runners, and the incompatible-platform dependency test
  picks the platform the host is not running on, so the Linux CI gate is honest.

## [0.9.26] — 2026-11-20

### Added

- **`Option Native`**: arithmetic stops converting text and reports `Type mismatch (13)`
  instead. `"1" + "2"`, `"2" * 3`, `"6" \ 2`, `"6" Mod 4` and `-"3"` all raise `13`, while
  `1 + 2`, `Empty + 1` (Empty still counts as the default `0`) and `Null + 1` (Null still
  propagates) behave as before. `&` is untouched — it is explicit text operation, not a
  conversion. `Option Typing` implies `Option Native`.
- `examples/option_native.vbs` and `tests/test_option_native.py` (22 cases).

### Changed

- **`Option` statements must now precede every other statement.** `_collect` scans options
  before compiling, so a late `Option Native Off` used to change how the code above it was
  compiled. Writing one after the first statement is now a syntax error (`1002`), matching
  what VBScript already requires of `Option Explicit`.
- **The keyword is `Native`, not `Strict`.** "Strict" says only that something is stricter,
  never what it buys; this option exists for speed as much as for discipline, and the point
  is that the compiler may emit *native* operations instead of a runtime type dispatch on
  every `+ - * /`. The old spelling also collides with VB6/VB.NET's `Option Strict On`,
  which means compile-time narrowing checks — that role belongs to `Option Typing` here. No
  alias is kept: the option was renamed before it was ever released. See
  `doc/VBS_SUPERSET.md` §2.7.1 for the full rationale.

### Docs

- `doc/VBS_SUPERSET.md` §2.7.1 (`Option Native`, including "Why `Native` and not `Strict`"),
  §4.1 (comparisons do not coerce — a registered gap with three `expectedFailure` cases) and
  §6.2; the Chinese edition follows.

## [0.9.24] — 2026-11-20

### Fixed

- **Arithmetic follows VBScript coercion again**: `+ - * / ^` used to compile straight onto
  Python's operators, so `"2" * 3` produced `"222"` and `1 + "2"` raised `Type mismatch`
  instead of returning `3`. The five operators now go through the runtime: `+` concatenates
  only when both sides are text, the others convert numeric strings, `/` and `^` always
  yield a Double (as in VBS), and a non-numeric string raises `13`. Checked against
  `cscript.exe`.
- **Null propagates through built-ins**: `Len(Null)` returned `4`, `UCase(Null)` returned
  `"NULL"` and `TypeName(Null)` returned `"NullType"`. The functions VBScript propagates
  Null through (`Len`, `Left`, `Mid`, `UCase`, `Abs`, `Int`, the date-part functions, …)
  now return Null; the conversion and math group (`CStr`, `CInt`, `Sgn`, `Sqr`, …) raises
  `Invalid use of Null (94)` — a new `OVPPInvalidUseOfNullError` — and the `Format*` group
  raises `13`. `&` treats Null as empty text, except `Null & Null`, which stays Null.
- **An undeclared variable reads as `Empty`**: without `Option Explicit`, reading a name
  that was never declared or assigned raised `Variable is undefined (1001)`, something
  VBScript only does under `Option Explicit`. The compiler now initialises such names once
  at the top of the module (and of each procedure), so the read costs nothing at run time.
  `Err.Line` and the reported source line keep pointing at user code when the failing
  statement sits inside a runtime helper.

### Added

- 24 conformance cases in `tests/test_vbs_superset.py` covering the three rules above.

## [0.9.23] — 2026-11-20

### Added

- **The `WScript` host object**: scripts started by `openvbs` now get the global object
  Windows Script Host provides, so legacy `cscript.exe` code runs unchanged.
  `WScript.Echo` prints its values separated by spaces; `WScript.Arguments` is indexed
  from zero and also exposes the `/name:value` form through `.Named`; `ScriptName` and
  `ScriptFullName` name the running file; `StdOut` / `StdErr` / `StdIn` are `TextStream`
  objects with `Write`, `WriteLine`, `WriteBlankLines`, and on input also `ReadLine`,
  `ReadAll`, `AtEndOfStream`, `Skip` and `SkipLine`; `Name`, `Version`, `FullName`,
  `Path`, `Interactive`, `Timeout`, `Sleep`, `Quit`, `CreateObject` and `GetObject`
  round out the surface. `WScript` is read-only (assigning to it is a syntax error).
  Documented in `doc/VBS_SUPERSET.md` §3.13.
- Script arguments reach the host: `openvbs script.vbs a b` puts `a` and `b` into
  `WScript.Arguments`, and `vbs_run(..., args=[...])` does the same when the runtime is
  embedded.
- `tests/test_host.py`: 27 cases over the identity properties, arguments, streams,
  `Quit`/`Sleep`/`Timeout` and the read-only rule.

### Fixed

- **COM fidelity**: reading a property whose value is another COM object
  (`folder.Files.Count`, `service.Count`) no longer calls the value's default member and
  fails with `Type mismatch (13)`. `GetObject` now accepts monikers (`winmgmts:`), a
  file/class pair, and the `GetObject(, "class")` form that attaches to a running object.
  Failure codes match VBScript: a missing ProgID reports `429` and a failed attach reports
  `432`. `TypeName(obj)` reports the coclass name (`FileSystemObject`,
  `SWbemServicesEx`) instead of `Object`.
- **Omitted arguments**: `GetObject(, "WScript.Shell")` and statement calls such as
  `P , 2` produce an `Empty` argument slot instead of failing with
  `Bad expression token: ','`.
- `Scripting.FileSystemObject` folder collections (`Files`, `SubFolders`) are real 1-based
  collections: `Count`, `Item(n)` and `For Each` work, and `TypeName` reports
  `Files` / `Folder` instead of `Variant()`.

### Changed

- `Declare ... Lib` is documented as **deliberately unsupported for the short term**: a
  `ctypes` binding is feasible but a mistyped native call terminates the process with an
  access violation that no `On Error` / `Try Catch` can intercept
  (`doc/VBS_SUPERSET.md` §3.11, `dev-docs/ROADMAP.md` 7c).
- The copyright year is now 2026 in `LICENSE`, `NOTICE` and every source header.

## [0.9.22] — 2026-11-20

### Added

- **Real `MsgBox` / `InputBox` dialogs**: the bundled `ui-support` package draws them with
  `tkinter`/Tk. The button set (`vbOKOnly` … `vbRetryCancel`), the icon, the default button
  and the modal flag are honoured, and the call returns the VBScript codes `1`–`7`;
  `InputBox` pre-fills the default text and returns `""` on cancel. Dialogs appear only for
  interactive `openvbs` command-line runs: stdin/stdout must be a terminal, `CI` must be
  unset, and Tk must be importable. Everything else (pipes, embedded `vbs_run()`, the IDE,
  `-c` code runs) keeps the console behaviour, so batch jobs never block.
- `OPENVBS_UI` (`auto` by default, `console`/`off`, `gui`/`on`) and `OVPP_UI_PATH` switch
  the behaviour; `python -m OpenVBSPP` reports the resolved mode via `ui.describe()`.
- `tests/test_ui_dialogs.py`: 12 cases over the button/icon mapping, activation rules and
  the console fallback.

### Added (docs and licensing)

- `dev-docs/GUI_DIALOGS_DESIGN.md` and its Chinese pair: feasibility, package shape,
  packaging cost, Tcl/Tk licensing and the risk table.
- **`THIRD_PARTY_LICENSES.md`**: verbatim CPython (PSF-2.0) and Tcl/Tk (`TCL`) license
  texts plus the libtommath / zlib / X11 notes, so an exe can be redistributed lawfully.
- `NOTICE` gained the Python and Tcl/Tk entries (they only reach an artifact through the
  single-file exe).
- `packaging/openvbs.spec`: `OpenVBSPP.ovpprt.ui` is declared, `tkinter` is no longer
  excluded, `vendor/ui-support` and the license files ship as data.

## [0.9.21] — 2026-11-20

### Added

- **VBScript literals**: hexadecimal `&H...`, octal `&O...` and date `#...#` literals, so
  `echo &HFF` (255), `echo &O17` (15) and `echo Year(#2020-05-06#)` (2020) now work. Date
  literals accept `yyyy-mm-dd`, `mm/dd/yyyy`, `d Month yyyy` and an `h:mm:ss AM/PM` suffix.
- **19 built-in functions** that VBScript has always had: `Sin Cos Tan Atn Exp Log`,
  `CCur CSng CDate`, `VarType` (with the `VT_*` constants), `RGB`, `ScriptEngine`,
  `ScriptEngineMajorVersion`, `ScriptEngineMinorVersion`, `ScriptEngineBuildVersion`,
  `GetLocale`, `SetLocale`, `IsMissing`, `GetRef`, and `LoadPicture` (which raises
  `OVPPObjectNotSupportedError`: there is no picture library).
- **`ExecuteGlobal`**: the statement form runs its code in the global scope; `Execute`
  accepts the same form.
- **A bare `End`** stops the whole script, as in VBScript. It raises a `ScriptEnd`
  (`BaseException`), so neither `On Error Resume Next` nor `Try ... Catch` can swallow it;
  the runner treats it as a normal exit. The block closers (`End If`, `End Sub`, ...) are
  unchanged.
- **`Stop`** is accepted and does nothing: it is a debugger breakpoint and there is no
  debugger here.
- **Predefined constants** that were still missing: the `MsgBox` button, icon and return
  values (`vbOKOnly`, `vbYesNo`, `vbQuestion`, `vbOK`, `vbYes`, `vbNo`, …),
  `vbObjectError`, and every `VarType` result (`vbEmpty`, `vbString`, `vbLong`,
  `vbDouble`, `vbBoolean`, `vbArray`, …), so `If VarType(x) = vbString Then` works.
- `tests/test_vbs_coverage.py`: 48 regression cases pinning every item in this release.

### Fixed

- `TimeValue("13:45")`, `TimeValue("1:45:00 PM")` and `Hour(Time())` no longer fail with
  `Cannot convert '...' to a date`. Time-only values use the VBScript serial-date origin
  (1899-12-30), and `_to_datetime` now accepts `datetime` / `date` / `time` objects as well
  as numeric serial dates.
- `GetRef` returns a usable procedure reference: both `f(21)` and `f()` invoke it, and the
  name lookup is case-insensitive.

### Changed

- `doc/VBS_SUPERSET.md` and its Chinese counterpart were corrected: the built-in table, the
  known-gap list and the "all native syntax is supported" paragraph were stale. The gap list
  now names the `MsgBox` / `InputBox` console stand-ins, `LoadPicture`, procedures added at
  run time by `ExecuteGlobal`, backward `GoTo` and `Declare`; the built-in table lists what
  actually exists.

## [0.9.20] — 2026-11-20

### Breaking

- **`Option COP` was removed.** Python interop must now be imported explicitly with
  `<script src="python-embed" />` at the top of the script. Writing `Option COP` is a
  syntax error whose message points at that import, and a `#python` block without the
  import fails at compile time with a message naming `python-embed`.
- **`OpenVBSPP.plugins.debug_tools` (`ast2str`) was removed.** `--save` still writes
  `<name>.ast.txt`, now produced with `ast.unparse`, and `--verbose` prints the generated
  Python the same way.

### Added

- **`python-embed` over package**: shipped inside the distribution
  (`OpenVBSPP/vendor/python-embed/`) so every release carries it, and mirrored
  byte-for-byte in `extras/python-embed/` for project-local installs (`over add python-embed`).
- `over list` ends with the packages bundled with the release (`com-support`,
  `python-embed`), and `over add <name>` may omit the source when the package is bundled.
- `over.resolve_source` falls back to the bundled copy, so `<script src="python-embed" />`
  works in a project that installed nothing.

### Changed

- User documentation is English-first, with a `<page>.zh-CN.md` counterpart for every page;
  maintainer notes (release process, roadmap, IDE/LSP design, try/catch design, object-model
  design, good first issues) moved to `dev-docs/` and out of the published docs site.
- **Every code comment and docstring is now bilingual and user-facing**: English first,
  then a `中文：…` line/paragraph. They explain what a module, class or function gives a
  script author, how to use it and what the limits are; implementation notes belong in
  `dev-docs/`. The convention is written down in `AGENTS.md` (文档约定).

- Maintainer documents in `dev-docs/` also ship as English | Chinese pairs
  (`X.md` + `X.zh-CN.md`), matching `doc/`.
### Fixed

- `over` no longer leaves a half-cloned git cache behind: an incomplete cache directory is
  dropped and fetched again instead of failing with `over.json not found`.

## [0.9.15] — 2026-11-19

Performance release: the four hot paths that made generated code slower than
hand-written Python are gone, and the package is now clean to publish.

### Added

- **CLI inline execution**: `openvbs -c "echo 1+1"`, `openvbs --code CODE`,
  plus `openvbs check|run -c CODE` and `openvbs [-c|-] -` for standard input.
  `-c` and a script path are mutually exclusive (exit code 2).
- **`CHANGELOG.md`** (this file) and **documentation site** sources
  (`mkdocs.yml`, `doc/index.md`).
- **Coverage reporting** in CI (`coverage run -m unittest …` →
  `coverage.xml` → Codecov, optional `CODECOV_TOKEN`).
- **ARM64 single-file build** lane in CI (`windows-11-arm` runner, tag or
  manual dispatch only).
- `docs` and `coverage` entries in the optional dependency groups.

### Changed

- **`requires-python` is now `>=3.10`**; Python 3.9 reached end of life in
  October 2025. `0.9.14` remains the last release that installs on 3.9 — the
  code itself is still 3.9-compatible, only the metadata moved.
- **Linux is no longer "experimental" in CI**: Windows and Linux lanes are both
  hard gates (Ubuntu keeps the cheapest matrix: 3.10 and 3.13).
- Licensing metadata moved to PEP 639 (`license = "Apache-2.0"` +
  `license-files = ["LICENSE", "NOTICE"]`), `setuptools>=77` required; the
  deprecated `License :: OSI Approved` classifier is gone.

### Fixed

- **Startup 27 ms → 13 ms** (lazy `import`): `re`, `json`, `datetime`,
  `random`, `math`, the filesystem helpers and `OpenVBSPP.runtime` are only
  imported when a script actually touches them.
- **Read-only `ByRef` parameters are no longer boxed**: only parameters the
  procedure (or a callee) can write get a `Ref`; the callee unpacks a stray
  `Ref` inline for cross-module and `AddressOf` calls.
- **Error frames are emitted on demand**: procedures without `On Error` /
  `Err` / `Resume` no longer pay two calls per invocation.
- **`For … Next` compiles to `range()`** when the step is a non-zero constant
  and the body does not assign the counter, with VBS semantics preserved
  (counter after the loop is `end + step`, `Exit For` skips the step).
- **`Scripting.Dictionary` member lookup is O(1)** instead of scanning every
  key (50 000 inserts: 103 s → 58 ms), and `&` concatenation takes a fast
  path for two strings.
- **Packaging**: `python -m build` is warning-free; the sdist now ships
  `packaging/openvbs.spec`, `tools/bench.py` and `doc/*.docx`, so an sdist
  consumer can rebuild the exe and run the benchmarks.

### Measured (Windows 11, CPython 3.13, `tools/bench.py`)

| Workload | vs. hand-written Python | before 0.9.15 |
|---|---|---|
| 1M arithmetic loop (in procedure) | **0.99×** | 1.65× |
| 1M arithmetic loop (module level) | **1.04×** | 4.19× |
| 100k array read/write | **1.00×** | 4.4× |
| 200k function calls | **1.35×** | 13.7× |
| 20k string concatenation | 1.39× | 1.8× |
| `fib(24)` recursion | 6.85× | 10.3× |
| 50k `Dictionary` inserts | 13.1× | ~2000× |

End-to-end the interpreter is 2.5–3× faster than `cscript.exe` on the same
scripts (startup included).

## [0.9.14] — 2026-11-19

### Added

- **Real COM support** as an over package (`com-support`): wraps `pywin32`,
  ships inside the wheel under `OpenVBSPP/vendor/com-support`, works on Windows
  with zero configuration and **no import in scripts** —
  `CreateObject("WScript.Shell")` just works. A copy in
  `<project>/packages/com-support` overrides the bundled one; install it
  explicitly with `openvbs over install --with-com`; remove it again with
  `openvbs over remove com-support`. `OPENVBS_COM=0` disables COM entirely.
- **Python dependencies in `over.json`**: `over install` collects the whole
  dependency graph, merges constraints, installs them with `pip` in one pass,
  and records the resolved versions (including transitive ones) in
  `over.lock` (lock format v2). `--frozen` restores exactly the locked set.
- **Package manager commands**: `over list`, `over info <name>`,
  `over remove <name> [--python]`, `over com`; bare `openvbs over` prints usage
  with exit code 0.
- **Web IDE** (`openvbs ide`) with `POST /api/run` and `POST /api/check`, and
  `check|run --json` on the CLI for editor integration.
- **Legal self-audit** in [`doc/LEGAL.md`](doc/LEGAL.md): independent
  implementation, no Microsoft sources or documentation distributed,
  nominative use of the VBScript trademark.

### Fixed

- `Scripting.Dictionary` member access raised `NameError` in some branches.
- Frozen installs rewrote `over.lock` byte-for-byte differently.

## [0.9.13] and earlier

See [`dev-docs/RELEASE.md`](dev-docs/RELEASE.md) for the pre-0.9.14 engineering notes and
[`dev-docs/ROADMAP.md`](dev-docs/ROADMAP.md) for what each milestone covered:

| Version | Headline |
|---|---|
| 0.1–0.3 | Lexer, parser, AST mapping; procedures and functions; `ByRef`, `On Error`, `With`, `Select Case`, core VBScript statements |
| 0.4 | Object model: `Class` / `Property` / `Me` / `Dictionary`; `Option Typing` static type checking |
| 0.5 | `Execute` / `Eval`, `Format*`, `Option Compare`, `GetObject`, default members |
| 0.6 | Modules and namespaces, `Inherits`, `#python` embedding, `Option COP`, deep `ByRef` |
| 0.7–0.8 | Typed arrays, `Try/Catch/Finally`, labeled `On Error GoTo`, label-style procedures |
| 0.9.0–0.9.5 | `Format()`, `FileSystemObject`, value semantics, `?.` / `??`, compound assignment, `AddressOf`, `RegExp`, `JSON`, file modules, `Resume` state |
| 0.9.6–0.9.7 | `Shared`, `Implements`, `Property Set` / `Let` |
| 0.9.8–0.9.10 | `over` M1: local-source install, packing, registry, `CreateObject` parity with VBScript |
| 0.9.11–0.9.13 | Module split, `--frozen`, `commit` field, `over pack` / Git sources / local registry, `check`/`run --json`, web IDE |

