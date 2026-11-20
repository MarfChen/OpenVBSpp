> English | **[中文](COOKBOOK.zh-CN.md)**

# Cookbook

Practical recipes, each one a complete script you can save as `recipe.vbs` and run with
`openvbs recipe.vbs` (or `python -m OpenVBSPP recipe.vbs`). Recipes marked **Windows** call
into COM and only work on Windows; everything else runs anywhere CPython runs.

If you only read one section, read [Migrate an old script](#migrate-an-old-script) — it lists
the handful of differences that actually bite people.

---

## Run something, fast

```powershell
# a file
openvbs hello.vbs

# one line, no file
openvbs -c "echo ""hello"" & "" from the command line"""

# from stdin (handy in pipelines)
Get-Content hello.vbs | openvbs -

# parse only, JSON diagnostics (exit code 1 when there are errors)
openvbs check broken.vbs --json
```

`check` compiles and stops before running, `run` compiles and runs — both are what editors and
CI call. `openvbs --version` prints the version.

Tabs and ANSI colours are fine: output is plain text, so `| Out-File log.txt` and `> log.txt`
both work.

---

## Arguments, exit codes, and the host object

```vbscript
' args.vbs — run with: openvbs args.vbs alpha beta /mode:fast
WScript.Echo "script: " & WScript.ScriptName
WScript.Echo "args: " & WScript.Arguments.Count

Dim i
For i = 0 To WScript.Arguments.Count - 1
    WScript.Echo "  [" & i & "] " & WScript.Arguments(i)
Next

If WScript.Arguments.Named.Exists("mode") Then
    WScript.Echo "mode: " & WScript.Arguments.Named("mode")
End If

WScript.Quit 3          ' exit code 3, nothing after this line runs
```

`WScript.Arguments` is 0-based and `Named` understands both `/name:value` and `/name=value`.
`Quit` takes an optional exit code (default `0`); it is the only way to set a non-zero exit
status. `WScript.Echo` joins its arguments with a space, exactly like `cscript`.

---

## Read and write text files

```vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.BuildPath(fso.GetSpecialFolder(2), "openvbs-demo.txt")   ' 2 = temp folder

Set out = fso.CreateTextFile(path, True)      ' True = overwrite
out.WriteLine "first"
out.WriteLine "second"
out.Close

Set src = fso.OpenTextFile(path, 1)           ' 1 = reading
Do While Not src.AtEndOfStream
    line = src.ReadLine
    WScript.Echo "line: " & line
Loop
src.Close

WScript.Echo "exists: " & fso.FileExists(path)
WScript.Echo "size: " & fso.GetFile(path).Size
fso.DeleteFile path
```

Only plain text is implemented (no binary streams, no `ADODB.Stream`). `.Count` on
`folder.Files` / `folder.SubFolders` works, `.Item(i)` is 1-based, and `For Each` iterates.

---

## Parse a CSV into a Dictionary

```vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set totals = CreateObject("Scripting.Dictionary")   ' keys are case-insensitive

Set data = fso.OpenTextFile("sales.csv", 1)         ' region,amount
Do While Not data.AtEndOfStream
    line = data.ReadLine
    parts = Split(line, ",")
    If UBound(parts) = 1 Then
        region = Trim(parts(0))
        amount = CDbl(Trim(parts(1)))
        If totals.Exists(region) Then
            totals(region) = totals(region) + amount
        Else
            totals.Add region, amount
        End If
    End If
Loop
data.Close

Dim key
For Each key In totals.Keys
    WScript.Echo key & ": " & FormatNumber(totals(key), 2)
Next
```

`Scripting.Dictionary` is built in (no COM needed) and case-insensitive, like the real one.

---

## Error handling: two styles

`On Error` is the VBScript way and keeps working:

```vbscript
On Error Resume Next
Set shell = CreateObject("No.Such.Class.Anyway")
If Err.Number <> 0 Then
    WScript.Echo "failed with " & Err.Number        ' 429 on Windows
    Err.Clear
End If
On Error GoTo 0
```

`Try` / `Catch` / `Finally` is the concise way — the handler receives an object with
`Number`, `Description` and `Message`:

```vbscript
Try
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile("nope.txt", 1)
Catch e When e.Number = 53
    WScript.Echo "missing file: " & e.Description
Catch e
    WScript.Echo "error " & e.Number & ": " & e.Description
Finally
    WScript.Echo "cleanup"
End Try
```

`Finally` always runs, including when the `Try` block hits a bare `End`.
A bare `End` stops the whole script with exit code 0 and cannot be caught — use `Quit 1`
when you want a failure exit code.

---

## Automate Windows with COM **Windows**

```vbscript
' env vars, a shell command, and a WMI query
Set shell = CreateObject("WScript.Shell")
WScript.Echo "OS: " & shell.ExpandEnvironmentStrings("%OS%")

' 0 = hidden window, True = wait, result is the exit code
code = shell.Run("cmd /c exit 5", 0, True)
WScript.Echo "child exit: " & code

Set wmi = GetObject("winmgmts:")
Set rows = wmi.ExecQuery("SELECT Caption, Version FROM Win32_OperatingSystem")
For Each row In rows
    WScript.Echo row.Caption & " (" & row.Version & ")"
Next
```

Object-valued properties chain like real COM (`folder.Files.Count`), `TypeName` returns the
type-library name (`IWshShell3`, `FileSystemObject`, `SWbemServicesEx`), and failures report
VBScript codes: `429` for "cannot create", `432` for "file/class not found". `GetObject`
accepts a moniker (`"winmgmts:"`), a file path, or `GetObject(, "WScript.Shell")` to attach to
a running instance. `WScript` itself is always available, even without COM.

---

## Ask the user **any platform, real dialogs on Windows**

```vbscript
answer = MsgBox("Save the report?", vbYesNoCancel + vbQuestion, "OpenVBS++")
WScript.Echo "picked: " & answer          ' 1 ok, 2 cancel, 6 yes, 7 no

name = InputBox("Report name:", "OpenVBS++", "quarterly-report")
WScript.Echo "name: " & name
```

On a desktop that satisfies the enable rules (started from the CLI, an interactive terminal,
not CI, Tk available) the bundled `ui-support` package shows real Tk dialogs and returns the
VBScript codes. Use `OPENVBS_UI=gui` to force them (`OPENVBS_UI=console` to force the
text-only behaviour), `OVPP_UI_PATH` to point at your own copy, and
`over add ui-support` to vendor it into a project.

From the web IDE or any non-terminal host the console fallback is used: `MsgBox` prints the
text and returns `1`, `InputBox` reads a line from stdin.

---

## Split a legacy dump into modules

Two forms, and the difference matters:

```vbscript
' merge: the file's procedures become part of this script
<script src="lib/format.vbs" />
WScript.Echo Money(12.5)

' isolate: everything lives under the alias, nothing leaks
<namespace src="lib/strings.vbs" alias="s" />
WScript.Echo s.Shout("hello")
```

`<script src>` is textual-ish inclusion for code you own; `<namespace src alias>` behaves like
an import: the module's globals stay inside the namespace, so two modules may define the same
helper name. There is no `import` keyword.

---

## Call Python for the hard parts

```vbscript
<script src="python-embed" />          ' opt in to the Python bridge

' Python does the work; results come back through the COP pool.
' 中文：让 Python 完成工作，结果经 COP 池取回。
#python
import json
from OpenVBSPP.ovpprt.cop import COP

payload = '{"b": 1, "a": 2}'
data = json.loads(payload)
COP["keys"] = ", ".join(sorted(data))
COP["total"] = sum(data.values())
#endpython

WScript.Echo "keys: " & COP("keys")     ' a, b
WScript.Echo "total: " & COP("total")   ' 3
```

`#python ... #endpython` runs real Python with the script's globals; `COP` is one dictionary
shared by both sides (`from OpenVBSPP.ovpprt.cop import COP` in Python, `COP("name")` in
VBScript), so it is the natural place to hand values across. This is the escape hatch for
JSON, regexes, HTTP, binary formats and anything else the VBScript surface does not cover.
`<script src="python-embed">` must be present; the block does nothing without it.

---

## Package reusable code with `over`

`over.json` at the project root describes dependencies:

```json
{
  "name": "reports",
  "version": "1.0.0",
  "entry": "main.vbs",
  "dependencies": {"format-utils": "file:../packages/format-utils"}
}
```

```powershell
openvbs over install            # resolve, write over.lock, install python deps
openvbs over install --frozen   # CI: fail unless the lock is satisfied exactly
openvbs over list               # over packages, bundled packages, python packages
openvbs over add ui-support     # vendored packages can be added by name
```

For quick experiments, drop a directory under `packages/<name>/` — a package there shadows the
bundled copy, no install step needed. `over pack` / `over publish` make `.ovppkg` archives; the
full reference is in [OVER.md](OVER.md).

---

## Turn on typing and native arithmetic

```vbscript
Option Explicit        ' every variable must be Dim'd (VBScript behaviour)
Option Typing          ' static checks: declarations and parameters may carry -> types
Option Native          ' + - * / stop converting text: "1" + "2" raises 13

Dim count -> Integer
count = 2

Function Twice(n -> Integer) -> Integer
    Twice = n * 2
End Function

WScript.Echo Twice(count)
```

`Option Native` is opt-in because it changes semantics; `Option Typing` implies it. Prefer it
for new code that should fail loudly instead of silently coercing.

---

## Migrate an old script

1. **Run it first.** `openvbs old.vbs` is usually enough — the VBScript core (control flow,
   classes, `Property`, `With`, `On Error`, `ReDim`, `Erase`, `Select Case`, `#date#`
   literals, `&H`/`&O`, `GetRef`, `ExecuteGlobal`, 95+ built-ins, `WScript`) is implemented.
2. **Replace file paths only if they were 8.3 or UNC.** `C:\data\x.txt` works as written; the
   path helpers keep Windows semantics even on Linux.
3. **`Declare` is not supported** (on purpose). Replace the native call with COM, a
   `#python` block, or an `over` package.
4. **`LoadPicture` raises**, there is no GUI picture type; keep images outside the script.
5. **`MsgBox`/`InputBox` become console calls** unless the dialog enable rules are met.
6. **Backward `GoTo`** is rejected; restructure with `Do ... Loop` (forward `GoTo` works).
7. **`Option Explicit` before everything else** — options cannot appear mid-script.
8. **Exit codes**: `WScript.Quit n` and bare `End` (exit 0) are supported; `Err.Raise`
   numbers surface as `OVPPRuntimeError.number`.

Auto-run the suite with `openvbs check old.vbs --json` in CI: it reports errors with line and
column numbers.

---

## Where to look next

- [VBS_SUPERSET.md](VBS_SUPERSET.md) — the language reference and the known-gap table.
- [OBJECT_MODEL.md](OBJECT_MODEL.md) / [OOP.md](OOP.md) — classes, `Property`, `Inherits`.
- [MODERN_FEATURES.md](MODERN_FEATURES.md) — `Option Typing`, `Try/Catch`, `#python`, sugar.
- [OVER.md](OVER.md) and [WEB_IDE.md](WEB_IDE.md) — packaging and the browser IDE.
- `examples/` in the repository — the same recipes as runnable files, plus `examples/demo/`
  for the end-to-end tour.
