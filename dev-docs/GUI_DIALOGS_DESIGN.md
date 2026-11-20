# GUI dialogs for MsgBox / InputBox (Tkinter) — design and licensing

> English | **[中文](GUI_DIALOGS_DESIGN.zh-CN.md)**

> **Abstract** — Today `MsgBox` prints its text to stdout and returns `1`, and
> `InputBox` reads one line from stdin. That runs, but it silently breaks the
> one thing VBScript code actually depends on: **the return value**. A script
> that asks `If MsgBox("Save?", vbYesNo) = vbYes Then` can never take the "yes"
> branch. This report checks whether Python's `tkinter` can close that gap,
> proposes a bundled over package that needs no import, and states the licensing
> obligations for the two cases that matter (pure-Python wheel vs. bundled exe).

---

## 1. What "UI controls" means in VBScript

VBScript itself has exactly **two** GUI entry points:

| Statement | What it does |
| --- | --- |
| `MsgBox(prompt, buttons, title)` | modal message box, returns `1`–`7` |
| `InputBox(prompt, title, default)` | modal single-line text prompt, returns a string |

Everything else that looks like a "UI control" in old scripts is not VBScript —
it is COM automation: `Shell.Application.BrowseForFolder`,
`UserAccounts.CommonDialog`, `InternetExplorer.Application` for HTA-style forms,
`WScript.Shell.Popup` for timeouts. Those keep working through COM
(`com-support`, v0.9.14+); Tkinter is not a substitute for them.

So the compatibility target is narrow and well defined: **make the two built-ins
behave like their Windows counterparts, including the return codes.**

---

## 2. Feasibility on this machine (measured)

| Fact | Value |
| --- | --- |
| Tcl / Tk version | 8.6 / 8.6, patchlevel `8.6.13` |
| Windowing system | `win32` |
| Tk root creatable | yes (hidden root, 443 font families enumerated) |
| Screen | 1536×864, session `Console` |

`tkinter.messagebox` exposes `showinfo`, `showwarning`, `showerror`,
`askquestion`, `askokcancel`, `askyesno`, `askyesnocancel`, `askretrycancel`.
Its private `_show(title, message, _icon, _type, **options)` forwards `type`,
`icon`, `default` and `parent` straight to the Tcl command `tk_messageBox`,
which supports **all six** VBScript button sets:

| VBScript buttons | value | `tk_messageBox -type` |
| --- | --- | --- |
| `vbOKOnly` | 0 | `ok` |
| `vbOKCancel` | 1 | `okcancel` |
| `vbAbortRetryIgnore` | 2 | `abortretryignore` |
| `vbYesNoCancel` | 3 | `yesnocancel` |
| `vbYesNo` | 4 | `yesno` |
| `vbRetryCancel` | 5 | `retrycancel` |

Icons map one-to-one (`vbCritical`→`error`, `vbQuestion`→`question`,
`vbExclamation`→`warning`, `vbInformation`→`info`), the default button maps to
`-default`, and the returned string maps back to the VBScript constants:

```
ok→1  cancel→2  abort→3  retry→4  ignore→5  yes→6  no→7
```

`InputBox` maps to `tkinter.simpledialog.askstring(title, prompt,
initialvalue=default)`, which returns `None` on Cancel — VBScript returns `""`,
so the wrapper translates it. Everything needed is present; no third-party
dependency is required.

**Conclusion: yes, Tkinter can implement both built-ins faithfully.** The
remaining work is activation policy, not capability.

---

## 3. Package shape: `ui-support`, bundled, no import

Following the `com-support` precedent, the runtime ships the code and the script
author does not import anything:

```
OpenVBSPP/vendor/ui-support/over.json
OpenVBSPP/vendor/ui-support/ui_support.py     # the real implementation
OpenVBSPP/ovpprt/ui.py                        # thin dispatcher used by builtins
extras/ui-support/…                           # byte-identical mirror
```

`over.json`:

```json
{
  "name": "ui-support",
  "version": "1.0.0",
  "kind": "feature",
  "entry": "",
  "description": "GUI dialogs for MsgBox and InputBox (Tkinter).",
  "capabilities": ["ui"],
  "dependencies": {}
}
```

Resolution order is the same as COM: project `packages/ui-support` (pinned
version) → bundled `OpenVBSPP/vendor/ui-support/`. Loading is cached by file
mtime. Removing the package or setting `OPENVBS_UI=console` restores the
current console behaviour.

### Activation rules

`MsgBox` / `InputBox` stay console-based unless all of the following hold:

1. `OPENVBS_UI` is unset (default `auto`) or explicitly `gui`;
2. the process was started through the **CLI entry point**
   (`python -m OpenVBSPP file.vbs`, `openvbs file.vbs`, `openvbs.exe`), not
   through the embedding API `vbs_run()` — this keeps library callers and the
   test suite non-interactive by construction;
3. `sys.stdin.isatty()`, and `CI` is not set in the environment;
4. `tkinter` imports and a hidden `Tk()` root can be created.

If any check fails, or if Tk raises `TclError` at dialog time (service session,
no display), the call falls back to today's console behaviour and the decision
is remembered for the rest of the process. A failed `gui` request prints one
warning to stderr and then degrades quietly.

`OPENVBS_UI=console` (alias `off`) forces the console path; `OPENVBS_UI=gui`
forces GUI even without a TTY (useful under `pythonw.exe`).

### Semantics to preserve

- `MsgBox` returns `1`–`7` exactly as above; unknown/combined button values
  degrade to `vbOKOnly`; `vbSystemModal` makes the dialog topmost.
- Title defaults to `OpenVBS++` when omitted; the prompt is the message text.
- `InputBox` returns the typed text, the pre-filled `default` on OK, and `""`
  on Cancel — never `None`.
- The dialog blocks like the real one; there is no timeout (VBScript has none).

---

## 4. Packaging impact

| Artifact | Cost | Notes |
| --- | --- | --- |
| sdist / wheel | **0 bytes** | `ui-support` is ~5 KB of Python; Tcl/Tk come from the user's interpreter |
| Windows exe | **+8.9 MB uncompressed** | measured on this machine: `tcl86t.dll` 1.82 MB, `tk86t.dll` 1.59 MB, `lib\tcl8.6` 3.14 MB, `lib\tk8.6` 1.56 MB, `Lib\tkinter` 0.72 MB, `_tkinter.pyd` 0.06 MB; UPX/PyInstaller compress this to roughly half |
| Linux install | — | document `apt install python3-tk`; without it the console fallback engages |

`packaging/openvbs.spec` currently excludes `tkinter`; shipping GUI dialogs in
the exe means removing that exclude and adding the Tcl/Tk license file to
`datas` (see §5). Keeping a console-only exe and a separate GUI exe is possible
but doubles the release matrix for little gain.

---

## 5. Licensing

### What is actually distributed

| Case | Redistributes Tcl/Tk? | New obligations |
| --- | --- | --- |
| `pip install openvbspp` (wheel/sdist) | no — the interpreter provides it | none |
| Windows single-file exe | **yes** (`tcl86t.dll`, `tk86t.dll`, script libraries) | Tcl/Tk license text must ship |
| Our own source | our code only | Apache-2.0 (unchanged) |

### Components and their licenses

| Component | License | SPDX | OSI-approved |
| --- | --- | --- | --- |
| CPython, `tkinter`, `_tkinter` | Python Software Foundation License 2.0 | `PSF-2.0` | yes |
| Tcl/Tk 8.6 runtime | Tcl/Tk License (BSD-style) | `TCL` | **no** |
| libtommath (inside `tcl86t.dll`) | public domain | — | n/a |
| zlib (inside Tcl/Tk) | zlib License | `Zlib` | yes |
| pywin32 (COM exe) | PSF-2.0 | `PSF-2.0` | yes |
| X11/fontconfig (Linux only) | MIT / BSD-style | `X11` | yes |

The Tcl/Tk license text lives at `tcl8.6/tk8.6/license.terms` (2,267 bytes,
40 lines in this environment). It is the standard wording: copyright held by
the Regents of the University of California, Sun Microsystems, Scriptics,
ActiveState, Apple and others; permission to "use, copy, modify, distribute, and
license this software and its documentation for any purpose" **provided that**
the copyright notice and the permission notice "appear in all copies" and that
"the name of the copyright holders may not be used to endorse or promote
products derived from this software without specific prior written permission".

Key properties, checked against the local copy and SPDX:

- **permissive, not copyleft** — proprietary redistribution is allowed;
- **no royalty, no copyleft, no per-file bookkeeping** — the only duty is
  notice retention;
- **not OSI-approved** (SPDX `isOsiApproved: false`). This is a *policy*
  question, not a legal-risk question; Fedora classifies it as a good license
  and it is FSF-free and GPL-compatible;
- the trailing paragraph is the usual US-Government "Restricted Rights" clause;
- `LICENSE_PYTHON.txt` from the interpreter **does not mention Tcl/Tk**, so the
  two notices are independent and both must be present in an exe.

### Obligations, concretely

1. Ship the Tcl/Tk license text verbatim inside the exe's data directory and
   reference it from `NOTICE`.
2. Add a `THIRD_PARTY_LICENSES.md` (or `LICENSES/`) with the two texts
   (PSF-2.0 for the embedded CPython, TCL for Tcl/Tk) and link it from
   `README`'s Licensing section.
3. Keep `NOTICE`'s existing Apache-2.0 and pywin32 entries.
4. Do not name the product after Tk ("OpenVBS Tk Edition") and do not imply
   endorsement by the copyright holders; a factual sentence such as "optional
   GUI dialogs use Python's `tkinter` (Tcl/Tk 8.6)" is nominative use and fine.
5. State the Linux dependency (`python3-tk`) in the install docs so users do
   not file "MsgBox does nothing" bugs.

### Risk assessment

| Risk | Level | Why |
| --- | --- | --- |
| License violation in wheel/sdist | **none** | no Tcl/Tk bytes are redistributed |
| License violation in exe | **low** | notice retention only; satisfied by §5.1–5.2 |
| Enterprise "OSI-only" policy block | **low–medium** | `TCL` is not OSI-approved; document the console fallback and that Tkinter is optional |
| Trademark | **low** | nominative use only; no Tk branding |
| Patent | **low** | no patent grant in the text, but no known patents covering message boxes |
| Export control | **none** | no cryptography; libtommath is arithmetic |
| GPL incompatibility | **none** | TCL is GPL-compatible |

Bottom line: **shipping the Python wheel costs nothing legally; shipping the exe
costs one license file and a NOTICE line.** The only friction is
procurement-style "OSI-approved allowlist" policies, and the console fallback
already covers them.

---

## 6. Alternatives considered

| Option | Verdict |
| --- | --- |
| `ctypes` + `user32.MessageBoxW` | Windows-only, but zero third-party bytes and native look. No `InputBox` in `user32`, so a custom dialog would still be needed. Good **second backend** for the exe, not a replacement. |
| COM `WScript.Shell.Popup` | already available via `com-support` on Windows; no button sets, no `InputBox` default, no free-standing dialog. Keep as-is. |
| HTA / `InternetExplorer.Application` | deprecated by Microsoft, unavailable on modern Windows; not a target. |
| Tkinter (this proposal) | works on Windows, macOS and Linux, no new dependency for pip users, covers both built-ins faithfully. **Chosen.** |

---

## 7. Test plan

- Unit tests keep the console path: they call `vbs_run()` directly, which is
  rule 2 of §3, and additionally pin `OPENVBS_UI=console` in
  `tests/test_vbs_coverage.py` so a future refactor cannot make CI block on a
  dialog.
- A GUI smoke test guarded by `OPENVBS_UI_TEST=1` creates a hidden root,
  asserts the button-set/icon/default mapping table, and closes the dialog with
  `after(0, …)`-style automation; it never runs in CI.
- Manual matrix: interactive console (dialogs appear), `pythonw` (dialogs
  appear), service/session-0 (falls back), Linux without `python3-tk` (falls
  back), `CI=1` (falls back).

---

## 8. Roadmap

| Version | Scope |
| --- | --- |
| 0.9.22 | `ui-support` package, `MsgBox`/`InputBox` dialogs, activation rules, tests, docs |
| 0.9.23 | optional `ui` API for file/folder pickers (a clearly-marked superset, not VBScript) |
| later | `ctypes` Windows backend if exe size becomes a problem |
