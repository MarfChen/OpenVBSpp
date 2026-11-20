# Built-in web IDE

> English | **[中文](WEB_IDE.zh-CN.md)**

`openvbs ide` starts a local, zero-dependency web IDE: edit, run, check, save and load VBS
scripts in the browser. It is an **optional extension component** — the compiler never
imports it, and it has no third-party dependencies. A package-manager slot exists but is
disabled for now.

## 1. Start

```bash
openvbs ide                       # default http://127.0.0.1:7860
openvbs ide --port 8000
openvbs ide --workspace ./scripts --no-browser
python -m OpenVBSPP.ide           # equivalent
```

The UI has the editor on the left (line numbers, Tab indent) and three tabs on the right:

| Tab | Contents |
|---|---|
| Output | stdout + runtime errors (with line number and source line) + duration |
| Diagnostics | compile only, no execution; `Option Typing` type errors appear here too |
| Extensions | capability matrix: run / check / package manager / language server |

Shortcuts: `Ctrl+Enter` runs, `Ctrl+S` saves, `Tab` indents.

## 2. Why not Gradio / Streamlit

Gradio, Streamlit and NiceGUI were evaluated; the choice was **stdlib HTTP + a single-page
front end**:

| Option | Size | What fits this project | Why rejected |
|---|---|---|---|
| **Built-in (chosen)** | 0 MB | Zero dependencies, works offline, testable, fully controlled UI | Limited editor (no highlighting) |
| Gradio | ~300 MB (pulls pandas etc.) | `gr.Code` out of the box | Too heavy; contradicts "zero runtime dependencies" |
| Streamlit | ~150 MB | Fast layouts | Re-runs the whole page per interaction — wrong for an IDE |
| NiceGUI | ~40 MB | Pretty Quasar UI | Still an optional dependency, and needs uvicorn |

The key judgement: **an IDE's core is the service layer, not the front-end framework.** So
`service.py` came first (pure stdlib, unit-testable) and the front end is a thin shell. A
Gradio front end can be added later by calling the same functions; front-end code stays out
of the core package.

> If syntax highlighting is genuinely needed, CodeMirror's static files can be added later
> (still zero Python dependencies, just one more static asset).

## 3. Architecture

```
OpenVBSPP/ide/
├── service.py      back-end service layer (run/check/files/examples/capabilities)  ← the part that needs tests
├── packages.py     package manager contract (reserved, contract only)
├── server.py       stdlib http.server, routing and JSON encoding only
├── static/index.html   single-page front end (inline CSS/JS, no build step)
├── __init__.py     public exports
└── __main__.py     python -m OpenVBSPP.ide
```

The front end talks to the back end **only through the JSON API**, so any front end can be
swapped in without touching the back end:

| Method | Path | Description |
|---|---|---|
| GET | `/api/capabilities` | version, capability matrix, package manager contract |
| GET | `/api/examples` | built-in examples |
| GET | `/api/scripts` | scripts in the working directory |
| POST | `/api/run` | `{source, filename}` → `{ok, output, error, duration_ms}` |
| POST | `/api/check` | `{source}` → `{ok, diagnostics}` |
| POST | `/api/save` | `{name, source}` |
| POST | `/api/load` | `{name}` → `{source}` |
| POST | `/api/packages/search` | **501** (reserved) |

## 4. Open / save: native system dialogs

Browsers **cannot** show a system "save as" dialog for a web page, so the local server calls
Tk's `filedialog` instead — the IDE listens on 127.0.0.1 only and the dialog appears on your
own desktop, so it feels like a local program.

| Button | Behaviour |
|---|---|
| Open… | `filedialog.askopenfilename` (native dialog on Windows/Linux/macOS) |
| Save | writes in place when a path is known, otherwise save-as |
| Save as… | `filedialog.asksaveasfilename`, default extension `.vbs` |

Three levels of degradation keep it usable anywhere:

1. **Native dialog** (tkinter plus a graphical environment);
2. **Browser file picker** (no tkinter, e.g. a minimal Linux container): `<input type=file>` + Blob;
3. **Working directory** (no dialog but a save is requested): writes to `~/.openvbspp/scripts`.

`GET /api/dialogs` tells the front end which level is active; the Extensions tab shows it too.

Related endpoints:

| Method | Path | Description |
|---|---|---|
| GET | `/api/dialogs` | `{platform, native_dialogs, newline}` |
| POST | `/api/dialog/open` | shows an open dialog → `{ok, path, source}` |
| POST | `/api/dialog/save` | shows a save-as dialog → `{ok, path}` |
| POST | `/api/file/open` | reads an absolute path |
| POST | `/api/file/save` | writes an absolute path |

## 5. Security boundary

- Listens on `127.0.0.1` only (default) and is not exposed to the network;
- File access is confined to the working directory; `../` escapes are rejected (covered by tests);
- Request bodies are capped at 2 MB;
- **Note**: a VBS script can execute arbitrary code (`#python` blocks, `Execute`), so the
  IDE is for local use only — do not deploy it as a public service.

## 6. Package manager interface (reserved)

`packages.py` currently defines only the contract, not an installer:

```python
class PackageProvider:          # package source protocol
    def search(self, keyword=""): ...
    def install(self, name, version=None): ...
    def uninstall(self, name): ...
    def installed(self): ...

class PackageInfo:              # what one package looks like
    name, version, description, provides, requires, homepage
```

Hooking it up (**no front-end change required**):

```python
from OpenVBSPP.ide.packages import set_provider, LocalPackageStore

class MyStore(LocalPackageStore):
    def search(self, keyword=""): ...      # later: PyPI / a local directory / the over index
    def install(self, name, version=None): ...

set_provider(MyStore())
```

`/api/capabilities` reflects the `enabled` state live, and the IDE's Extensions tab shows
"reserved" or "enabled" accordingly. The real package management entry point on the CLI is
`openvbs over ...` (see [OVER.md](OVER.md)); the IDE's `/api/packages/*` is not wired to
`over` yet.

## 7. Tests

```bash
python -m unittest tests.test_ide -v
```

Coverage: the service layer (run / error shape / check / every example runs), the working
directory (save / read / path escape), and an HTTP smoke test (index, capabilities, run,
reserved endpoint returning 501).
