# OpenVBS++ build and release

> English | **[中文](RELEASE.zh-CN.md)**

This document covers both release paths:

1. **The PyPI package** (`pip install openvbspp`) — for Python users;
2. **The single-file exe** (`openvbs.exe`) — for Windows users without Python.

---

## 0. One-time setup

```powershell
# packaging tools (build wheel / upload / build exe)
.\venv\Scripts\python.exe -m pip install -U build twine pyinstaller
```

PyPI credentials belong in `%USERPROFILE%\.pypirc`:

```ini
[distutils]
index-servers =
    pypi
    testpypi

[pypi]
username = __token__
password = pypi-<your API token>

[testpypi]
repository = https://test.pypi.org/legacy/
username = __token__
password = pypi-<TestPyPI token>
```

---

## 1. Pre-release checks

```powershell
# 1) the full test suite must be green
.\venv\Scripts\python.exe -m unittest discover -s tests

# 2) the version number agrees in five places
Select-String -Path OpenVBSPP\__init__.py,OpenVBSPP\ast_nodes\__init__.py,`
  OpenVBSPP\ovpprt\__init__.py,OpenVBSPP\plugins\__init__.py,`
  pyproject.toml -Pattern "0\.\d+\.\d+"

# 3) clean build (remove old artifacts first)
Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue
```

Version rule: `OpenVBSPP/__init__.py`, `ast_nodes/__init__.py`, `ovpprt/__init__.py`,
`plugins/__init__.py` and `pyproject.toml` **must agree** (a test enforces it too).

---

## 2. Build the wheel + sdist

```powershell
.\tools\build.ps1 -Target wheel
```

which is equivalent to:

```powershell
.\venv\Scripts\python.exe -m build
```

Artifacts:

```
dist/
├── openvbspp-0.9.20-py3-none-any.whl     # pure Python, no platform dependency
└── openvbspp-0.9.20.tar.gz               # sdist (includes examples/tests/doc)
```

Local self-check:

```powershell
# is the metadata valid?
.\venv\Scripts\python.exe -m twine check dist\*

# try installing into a throwaway environment
python -m venv .tmp-venv
.tmp-venv\Scripts\python.exe -m pip install dist\openvbspp-0.9.20-py3-none-any.whl
.tmp-venv\Scripts\python.exe -m OpenVBSPP examples\control1.vbs
.tmp-venv\Scripts\python.exe -m OpenVBSPP --version
```

---

## 3. Upload to PyPI

Go to TestPyPI first, then the real index:

```powershell
# 3.1 TestPyPI
.\venv\Scripts\python.exe -m twine upload --repository testpypi dist\*

# 3.2 install from TestPyPI to verify (note --index-url and --extra-index-url)
python -m pip install --index-url https://test.pypi.org/simple/ `
  --extra-index-url https://pypi.org/simple/ openvbspp

# 3.3 real PyPI
.\venv\Scripts\python.exe -m twine upload dist\*
```

After publishing, check:

- <https://pypi.org/project/openvbspp/> — page, README rendering, classifiers;
- `pip install openvbspp` then `openvbs --version` works;
- the same version **cannot be uploaded twice** (bump the version or use `--skip-existing`).

---

## 4. Building the exe

```powershell
.\tools\build.ps1 -Target exe
```

Artifact: `dist\openvbs.exe` (single file, double-clickable or run from a terminal):

```
openvbs.exe examples\control1.vbs
openvbs.exe --version
openvbs.exe            # no arguments: run the built-in demo
```

Implementation notes (`packaging/openvbs.spec`):

- the entry point is `packaging/entry.py`, which only calls `OpenVBSPP.cli:main`;
- generated code imports `OpenVBSPP.ovpprt.*` at run time, invisible to static
  analysis, so the spec lists those modules in `hiddenimports` (**remember to add
  new runtime modules**);
- `com-support` is a **data file** (loaded at run time with `spec_from_file_location`),
  also invisible to static analysis, so the spec must ship
  `OpenVBSPP/vendor/com-support` through `datas` — otherwise
  `CreateObject("WScript.Shell")` inside the exe degrades to "COM unavailable";
- `import win32com` inside `com_support.py` is equally invisible: if pywin32 is
  installed at build time the spec adds `pythoncom` / `pywintypes` / `win32com.*` to
  `hiddenimports` and packs `pywin32_system32/*.dll` as `binaries` — so **the exe
  carries COM support** and the target machine needs no pywin32; building without
  pywin32 still works, COM is just unavailable;
- `excludes` drops `tkinter` / `astpretty` / `unittest` to shrink the binary;
- install `upx` for a smaller file; `upx=True` is already set in the spec.

---

## 5. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `twine` says `File already exists` | that version is published; bump the version and rebuild |
| the README does not render on PyPI | `README.md` uses syntax PyPI does not support (some HTML comments) |
| the exe raises `ModuleNotFoundError: OpenVBSPP.ovpprt.xxx` | a new runtime module was not added to `hiddenimports` |
| antivirus flags the exe | common with PyInstaller one-file builds; switch to `--onedir` (edit `EXE`/`COLLECT` in the spec) |
| the wheel has no examples | expected: `MANIFEST.in` governs the sdist, the wheel only ships the package |

---

## 6. Release notes

The text pasted into PyPI / GitHub Releases is taken from here; keep it aligned
with the README changelog and the version table in `dev-docs/ROADMAP.md`.

### v0.9.20 — explicit Python interop import, English-first docs

- Python interop (`#python ... #endpython` and `COP`) is no longer enabled by
  `Option COP`: it requires `<script src="python-embed" />` (an over package
  shipped with the installation; `over add python-embed` vendors it into the
  project). `Option COP` is now a syntax error with a migration hint.
- `plugins/debug_tools.py` (`ast2str`) is removed; `--save` still writes
  `.ast.txt`, now via the standard-library `ast.unparse`.
- Documentation is English-first with a `.zh-CN.md` twin for every page;
  maintainer documents moved to `dev-docs/` and are no longer part of the docs site.
- `over list` now lists the packages shipped with the installation; `over add
  <name>` may omit the source; a half-written git cache self-heals.

### v0.9.15 — performance

A pure performance release: **only the generated Python code changed, no language
semantics**, with every boundary pinned by `tests/test_perf_semantics.py`.

- **Lazy imports at startup**: `re`, `json`, `datetime`, `math` and `random` in
  `ovpprt/regexp.py`, `jsonobj.py`, `object.py` and `builtins.py` are imported on
  first use; `cli.py` imports `OpenVBSPP.runtime` only when it actually runs a
  script. Cold start 27 ms → 13 ms.
- **Read-only `ByRef` parameters are no longer boxed**: the compiler decides
  statically whether a parameter is really written (assignment, `ReDim`, `For`,
  or passing it further down) and passes read-only ones by value; callees unpack
  inline with `if type(n) is Ref: n = n.v`, which keeps cross-module `namespace`
  calls and `AddressOf` delegates working. 200k function calls 118 ms → 23 ms
  (13.7× → 1.35× hand-written Python).
- **Procedures without `On Error` build no error frame**: when the body has no
  `On Error ...` / `Resume` / `Err`, no `enter()/try/finally/leave()` is generated.
  Recursive fib(24) 67 ms → 38 ms (12.0× → 6.85×).
- **`For` compiles to `range` when safe**: constant step and no assignment to the
  counter inside the body produce `for i in for_seq(...)` plus
  `else: i = i + step`, reproducing the three VBScript rules (counter = last value
  + step after a normal exit, no increment after `Exit For`, empty range keeps the
  start value). 1M loop 50 → 29.7 ms (0.99× hand-written Python).

Also: 457 → 485 test cases; `tools/bench.py` notes and the README performance
section were updated together.

### v0.9.14 — Python dependencies and COM support for `over`

- `over.json` gained a `python` field: the Python dependencies of the whole graph
  go to pip in one pass, and `over.lock` (v2) records the real closure plus
  `required_by`; `over list` / `info` / `remove --python` inspect and uninstall.
- `over.json` gained a `platforms` field; a non-matching platform is skipped, not
  an error.
- COM support became the over package `com-support` (wrapping `pywin32`): shipped
  with the installation (zero config on Windows), the project copy wins when
  present, `over install --with-com` vendors it, `OPENVBS_COM=0` disables it; the
  script still writes `CreateObject("WScript.Shell")` with **no import**.
- Fixed three uses of an undefined variable in the `Dictionary` member dispatch
  paths in `object.py`.

---

## 7. CI (GitHub Actions)

Two workflows: `.github/workflows/ci.yml` (tests + exe) and
`.github/workflows/docs.yml` (docs site, see §7.1).

**`test` (every push / PR / manual trigger)**

| Dimension | Value |
|---|---|
| Windows (hard gate) | Python 3.10 / 3.11 / 3.12 / 3.13 |
| Linux (hard gate) | Python 3.10 / 3.13; `continue-on-error` was removed in v0.9.15 |
| Steps | toolchain → `pip install -e ".[dev]" --no-build-isolation` → pywin32 (Windows only) → `coverage run -m unittest discover -s tests` → `coverage report` + `coverage xml` → upload to Codecov (Windows 3.13 with a token) → entry point self-check (including `-c` inline) → COM smoke (Windows only) → `python -m build` → clean-venv wheel smoke (asserting `vendor/com-support` exists plus a `-c` smoke) → upload `dist` artifacts |

Windows is a hard gate because COM, WScript semantics and path behaviour are all
validated there. Linux is a **hard gate** since v0.9.15 as well (3.10 and 3.13);
the matrix no longer has an `experimental` dimension — if a Linux version becomes
flaky, add `continue-on-error: true` back for that entry only.

**`exe` (runs only on tags or manual trigger)**

PyInstaller builds `dist/openvbs.exe`, copies it to a temp directory for a smoke
test (`--version` plus a real `CreateObject("WScript.Shell")`), and uploads it as
the `openvbs-exe-x64` / `openvbs-exe-arm64` artifact. The matrix is
`[{runner: windows-latest, arch: x64}, {runner: windows-11-arm, arch: arm64}]`;
note that the `windows-11-arm` runner is **only available to public repositories**
and fails outright on private ones.

### What you must do yourself to finish CI

1. **Create the repository and push** (without this there is nowhere for Actions
   to run and the badge stays grey):
   ```bash
   git remote add origin https://github.com/<you>/OpenVBSpp.git
   git push -u origin main --tags
   ```
2. Repository Settings → Actions → General: confirm workflows are allowed (the
   default for new repositories; nothing to change).
3. **No secret is needed** to run the tests. Only these two extras need one:
   - publishing to PyPI from CI: add a `PYPI_API_TOKEN` secret and a
     `twine upload` job;
   - a coverage badge: add a `CODECOV_TOKEN` secret (the upload step already
     exists; without the secret its `if` skips it, so the badge stays grey
     instead of failing).
4. The first push produces results in about 3–5 minutes: 6 test matrix jobs
   (4 Windows + 2 Linux) plus 2 manual exe jobs. Once everything is green the CI
   badge turns green by itself — **no README change needed**.
5. To bring the docs site online: Settings → Pages → Build and deployment →
   Source = **GitHub Actions** (the default is "Deploy from a branch"; without
   this switch `docs.yml` fails).

---

## 8. Badges (how they turn green)

The README hero has 7 badges in two categories:

| Badge | Type | What keeps it green |
|---|---|---|
| `tests-562 passing` | static (a hand-written shields.io string) | always green; **changing the case count means changing both `README.md` and `README.zh-CN.md`**, otherwise it is fake green |
| `version-0.9.20` | static | bump together with the five version locations on release (two README files plus two body mentions) |
| `license-Apache--2.0` | static | always green |
| `python-3.10+` | static | always green; update when `requires-python` changes |
| `dependencies-none` | static | always green (`pywin32` is an optional extra, not a runtime dependency) |
| **CI** | **dynamic** | after the push, one successful workflow run on `main` turns it green; `no status` means the default branch has no runs yet |
| **coverage** | **dynamic** | served by Codecov; grey without `CODECOV_TOKEN`, green after the first successful upload |

Static badges are the project's own claim, dynamic CI badges are machine-measured;
keep both, but reviewers trust the latter more.

Two more dynamic badges can be added once the package is on PyPI (adding them now
would show a red "package not found"):

```markdown
![pypi](https://img.shields.io/pypi/v/openvbspp)
![pyversions](https://img.shields.io/pypi/pyversions/openvbspp)
```

Coverage collection has been in CI since v0.9.15 (`coverage run` → `coverage xml`
→ `codecov-action@v5`) and the README already shows the Codecov badge; **configure
`CODECOV_TOKEN` in the repository and it turns green after the first successful
run**. If you would rather not use Codecov, delete the badge line from the README
and the upload step from `ci.yml`.

---

## 9. Python versions, artifacts and ARM

### 9.1 How old a Python is supported

| Item | Current state |
|---|---|
| Declaration | `requires-python = ">=3.10"`, classifiers list 3.10 – 3.13 |
| Static check | `vermin --eval-annotations --target=3.9 --violations OpenVBSPP cli.py tests` → **zero violations**, minimum 3.9 (no `match`, no `X | Y` annotations, no `tomllib`, no `except*`) |
| Runtime measurement | Windows embeddable **3.9.13** and the dev venv **3.13.15** each ran 485 cases **all green** (3.9 is a compatibility measurement, outside the declared range) |
| CI | Windows 3.10 / 3.11 / 3.12 / 3.13 hard gate; Linux 3.10 / 3.13 hard gate |

**The change already made (v0.9.15)**: 3.9 reached end of security support in
October 2025, so the declared floor moved to **3.10**. **0.9.14 is the last version
that declared support for 3.9**; the runtime code itself is still 3.9 compatible
(the vermin verdict did not change, we simply no longer promise it), so anyone
willing can still run it on 3.9. Raising the floor later only needs three places:
`requires-python` and the classifiers in `pyproject.toml`, plus the Python badge in
both READMEs.

### 9.2 What to upload to PyPI

Only two files, and **both are pure Python**:

```powershell
python -m build                        # dist\openvbspp-0.9.20-py3-none-any.whl + dist\openvbspp-0.9.20.tar.gz
python -m twine check dist/*           # metadata + README rendering
python -m twine upload --repository testpypi dist/*   # TestPyPI first
python -m twine upload dist/*                          # then the real PyPI
```

- **Do not** build per Python version: there are no C extensions and no ABI
  bindings, so one universal wheel (`py3-none-any`) covers every Python version and
  every CPU architecture.
- **Do not** upload the exe to PyPI: it is a self-contained executable per OS /
  architecture (about 8 MB) and belongs to GitHub Releases, not to PyPI.
- Prefer [Trusted Publishing](https://docs.pypi.org/trusted-publishers/) (OIDC)
  over an API token; if you use a token, put `PYPI_API_TOKEN` in the repository
  secrets, never in code or a workflow in plain text.

### 9.3 ARM (Windows ARM64 / Linux arm64)

Two separate questions with completely different answers:

1. **`pip install` on an ARM machine: nothing to do.** A universal wheel installs
   on any architecture; on Windows ARM64 the real risk, `pywin32`, already ships a
   `win_arm64` wheel (measured: pywin32 **312** provides
   `pywin32-312-cpXX-cpXX-win_arm64.whl` for cp39–cp315), so
   `pip install "openvbspp[com]"` works on ARM64 Windows.
2. **The single-file exe: needs an ARM machine to build.** PyInstaller **cannot
   cross-compile**; the artifact architecture equals the builder architecture. Two
   routes:
   - **GitHub-hosted ARM runners**: public repositories just set `runs-on:
     windows-11-arm` (public preview since April 2025, generally available for
     public repositories since August 2025; **private repositories do not support
     this label** and fail outright). Add a matrix line to the `exe` job in
     `ci.yml`:
     ```yaml
     exe:
       strategy:
         matrix:
           os: [windows-latest, windows-11-arm]
       runs-on: ${{ matrix.os }}
     ```
     name the artifact with the architecture, e.g.
     `openvbs-${{ matrix.os == 'windows-11-arm' && 'arm64' || 'x64' }}.exe`.
   - **A local ARM device** (Surface / Snapdragon X / Dev Kit): run
     `python -m PyInstaller packaging/openvbs.spec --noconfirm` and the artifact is
     arm64.
   Linux arm64 is the same story with the `ubuntu-24.04-arm` runner; the exe itself
   only targets Windows, so Linux distribution is just `pip install`.

### 7.1 Docs site (GitHub Pages)

The docs site has been publishable since v0.9.15: the repository-root
`mkdocs.yml` points at `doc/` as the site root, `doc/index.md` is the landing page,
and `.github/workflows/docs.yml` builds and deploys to GitHub Pages whenever
`doc/**` or `mkdocs.yml` changes (pull requests build without deploying).

```powershell
pip install "mkdocs-material>=9.5"     # or pip install -e ".[docs]"
mkdocs serve                          # local preview at http://127.0.0.1:8000
mkdocs build --site-dir site          # generate the static site (site/ is gitignored)
```

`.docx` files in `doc/` are copied into the site verbatim; the `.md` structure is
decided by `nav`. When adding a document, remember to add it to `nav` in
`mkdocs.yml`, otherwise it is only reachable through search.

### 9.4 Version / architecture release self-check

```powershell
python -m unittest discover -s tests                                    # 562 OK
python -m twine check dist/*
pip download --only-binary=:all: --no-deps openvbspp==0.9.20 -d tmp    # must yield py3-none-any.whl
```

If you get `py3-none-any.whl`, the package has no extra requirement on either the
Python version or the CPU architecture.

### 9.5 Packaging metadata hygiene (done in v0.9.15)

`python -m build` is now **warning-free**; the changes:

- licence metadata moved to PEP 639 style: `license = "Apache-2.0"` plus
  `license-files = ["LICENSE", "NOTICE"]`, `build-system.requires` raised to
  `setuptools>=77`; the deprecated `License :: OSI Approved` classifier removed.
  The wheel's `dist-info/licenses/` now carries `LICENSE` and `NOTICE`, satisfying
  Apache-2.0's requirement to keep the notices with distributed files.
- `MANIFEST.in` dropped the Chinese pattern line that Windows reads as cp936
  mojibake (its explanation lives in `.gitignore` and this document) and added
  `packaging/*` and `tools/*.py`, so the exe (`packaging/openvbs.spec`) can be
  rebuilt and the benchmark (`tools/bench.py`) run from the sdist as well. The
  `doc/` rule is "only `.md` plus explicitly listed specifications": the
  `doc/OpenVBS spec.docx` referenced by the README is in the sdist, while the
  internal planning note `doc/plan.docx` is excluded and stays in the repository.
- the `packages` list gained `OpenVBSPP.ide.static` and `OpenVBSPP.vendor`,
  silencing setuptools' "absent from the `packages` configuration" notice.

Self-check:

```powershell
python -m build --outdir dist      # should print no warnings
python -m twine check dist/*
```
