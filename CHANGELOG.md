# Changelog

All notable changes to **OpenVBS++** are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Release-by-release engineering notes (how to cut a release) live in
[`doc/RELEASE.md`](doc/RELEASE.md); the roadmap is in [`doc/ROADMAP.md`](doc/ROADMAP.md).

## [Unreleased]

Nothing yet.

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

See [`doc/RELEASE.md`](doc/RELEASE.md) for the pre-0.9.14 engineering notes and
[`doc/ROADMAP.md`](doc/ROADMAP.md) for what each milestone covered:

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

