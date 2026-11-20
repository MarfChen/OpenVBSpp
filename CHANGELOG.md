# Changelog

All notable changes to **OpenVBS++** are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Release-by-release engineering notes (how to cut a release) live in
[`dev-docs/RELEASE.md`](dev-docs/RELEASE.md); the roadmap is in [`dev-docs/ROADMAP.md`](dev-docs/ROADMAP.md).

## [Unreleased]

### Removed

- `cli_demo.py` — with no arguments `python -m OpenVBSPP` already runs the
  built-in `DEMO`, so this file duplicated reachable behaviour; runnable
  samples live in `examples/*.vbs`.
- `requirements.txt` — it only held comments (the runtime has no third-party
  dependencies). Dependencies and extras are declared in `pyproject.toml`
  (`[project.dependencies]` / `[project.optional-dependencies]`).

### Changed

- `MANIFEST.in` narrows the `doc/` packaging rules: the language spec
  `doc/OpenVBS spec.docx` still ships in the sdist (the README links to it),
  while the internal planning essay `doc/plan.docx` is explicitly excluded and
  stays repository-only.

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

