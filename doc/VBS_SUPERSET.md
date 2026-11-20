# VBS superset conformance

> English | **[中文](VBS_SUPERSET.zh-CN.md)**

The superset matrix: every VBScript construct and whether OpenVBS++ supports it (✅),
deviates semantically (⚠️) or declares it unsupported (❌). Sections marked "extension" are
additions that VBScript itself would reject, so they cannot break existing scripts.

> Project rule: **OpenVBS++ must be a superset of VBScript.**
> Every legal VBScript program must compile and run, unless it is explicitly listed under
> "known gaps" below. Extension syntax may only **add**; it must never change existing VBS
> semantics.

## 1. How the rule is enforced

| Mechanism | Location |
|---|---|
| Executable case matrix: legal VBS snippets must produce the expected result | `tests/test_vbs_superset.py` → `VBS_FEATURES` |
| Known gaps are marked `expectedFailure`; when a gap is closed the case turns red and reminds you to update the docs | `tests/test_vbs_superset.py` → `TestKnownGaps` |
| Documentation matrix (this file) | `doc/VBS_SUPERSET.md` |
| Change checklist | `AGENTS.md` |

When adding a feature: write the case in `VBS_FEATURES` first, then implement it, then
update the status column here.

## 2. Support matrix

Legend: ✅ supported · ⚠️ supported with semantic differences (see §3) · ❌ not implemented (see §4)

### 2.1 Lexing and statements

| VBS feature | Status | Notes |
|---|---|---|
| Newline separates statements | ✅ | |
| `:` separates several statements | ✅ | `a = 1 : b = 2` |
| `_` line continuation | ✅ | `space + _` at end of line |
| `'` comments / `rem` comments | ✅ | Whole-line and trailing |
| CRLF / CR / LF line endings | ✅ | Normalised automatically |
| Keywords are case-insensitive | ✅ | Keywords and identifiers are normalised to lower case |
| UTF-8 / BOM | ✅ | The CLI reads with `utf-8-sig` |

### 2.2 Data and expressions

| VBS feature | Status | Notes |
|---|---|---|
| Variables need no declaration | ✅ | |
| `Dim` / `ReDim [Preserve]` / `Const` | ✅ | Array lower bound 0, upper bound inclusive |
| `Option Explicit` | ✅ | Compile-time check with line numbers |
| Integers / floats / strings / booleans | ✅ | |
| `Empty` / `Null` / `Nothing` literals | ✅ | Three sentinels |
| `+ - * /` | ✅ | VBS coercion: `+` concatenates only when both sides are text, `/` always yields a Double; `Option Native` turns the coercion off (§2.7.1) |
| `\` integer division, `Mod` | ✅ | Truncates toward zero |
| `^` exponentiation | ✅ | Always a Double; left-associative (same as VB) |
| `&` string concatenation | ✅ | Null counts as empty text; `Null & Null` stays Null |
| `= <> < > <= >=` | ✅ | `=` in an expression is equality |
| `==` as equality | ⚠️ | Our extension; VBS lacks it, and accepting it cannot break VBS code |
| `Is` object identity | ✅ | `x Is Nothing` |
| `And Or Not Xor Eqv Imp` | ⚠️ | Bitwise semantics (as in VBS), see §3.1 |
| `True` / `False` | ✅ | True = -1, False = 0 |
| Predefined constants such as `vbCrLf` | ✅ | String constants (`vbCrLf` `vbCr` `vbLf` `vbTab` `vbNullString`), `FormatDateTime` names, the `MsgBox` button/icon/return constants (`vbOKOnly` … `vbNo`), `vbObjectError` and the `VarType` results (`vbString` `vbLong` `vbArray` …) |

### 2.3 Control flow

| VBS feature | Status |
|---|---|
| `If ... Then ... End If` (block) | ✅ |
| One-line `If cond Then stmt [Else stmt]` | ✅ |
| `ElseIf` repeated | ✅ |
| `Select Case` / `Case value, range To, Is comparison` / `Case Else` | ✅ |
| `For i = a To b [Step s] ... Next` | ✅ |
| `For Each x In collection ... Next` | ✅ (arrays and built-in collections) |
| `Do [While\|Until] ... Loop [While\|Until]` | ✅ |
| `While ... Wend` | ✅ |
| `Exit Do` / `Exit For` / `Exit Sub` / `Exit Function` | ✅ |
| A bare `End` (stop the whole script) | ✅ | Ends the script at once; neither `On Error Resume Next` nor `Try` can catch it |
| `Stop` (debugger breakpoint) | ⚠️ | Accepted and ignored — there is no debugger to break into |

### 2.4 Procedures and scope

| VBS feature | Status | Notes |
|---|---|---|
| `Sub` / `Function` / `End Sub` / `End Function` | ✅ | Script level only (VBS is the same) |
| Parameters `ByRef` (default) / `ByVal` / array parameter `a()` | ✅ | ByRef uses a `Ref` box written back at the call site |
| Function return value (assigning to the function name) | ✅ | Returns `Empty` when never assigned |
| Recursion | ✅ | |
| Three call forms `Foo` / `Foo 1,2` / `Call Foo(1,2)` | ✅ | |
| `Let` / `Set` statements | ⚠️ | Syntax accepted; `Set` object semantics follow the object model (see §4) |
| Script-level variables read/written inside procedures | ✅ | `global` declarations generated automatically |

### 2.5 Error handling

| VBS feature | Status | Notes |
|---|---|---|
| `On Error Resume Next` | ✅ | Per-statement try/except plus the `_ovpp_err` frame stack |
| `On Error GoTo 0` | ✅ | |
| `Err.Number` / `.Description` / `.Source` | ✅ | |
| `Err.Clear` / `Err.Raise` | ✅ | |
| Error numbers follow the VBS convention | ⚠️ | See §3.3 |

### 2.6 Objects and collections

| VBS feature | Status | Notes |
|---|---|---|
| `With ... End With` | ✅ | The expression is evaluated once |
| `.member` / `.Method()` syntax | ✅ | Fields/properties are fetched directly; method calls go through `member_call` |
| `Class ... End Class` | ✅ | Compiled to `_ovpp_cls_<name>` |
| Fields `Public` / `Private` / `Dim` | ✅ | Private fields/methods are renamed; outside access is rejected |
| `Property Get/Let/Set` | ✅ | Mapped to Python `@property` / setter (`Set` and `Let` merge, see [OOP.md](OOP.md)) |
| `Me` / `Class_Initialize` / `Class_Terminate` | ✅ | `Me` → `self`; constructor/destructor called automatically |
| `New Foo` | ✅ | |
| `Set x = obj` / `Set x = Nothing` / `x Is Nothing` | ✅ | |
| `CreateObject("script class")` | ✅ | Resolved at compile time or via the runtime registry |
| `CreateObject("Scripting.Dictionary")` | ✅ | `ovpprt/collections.py` |
| `Collection` | ❌ | VBScript has no built-in Collection (that is VBA) |
| `GetObject` | ✅ | `GetObject(path)`, `GetObject(path, class)` and `GetObject(, class)`; monikers such as `winmgmts:` resolve through the `com-support` package on Windows; attaches to a running instance and falls back to creating the class |
| `CreateObject` for COM objects | ✅ | Since v0.9.14 on Windows this is real COM via the `com-support` package (pywin32); other platforms report a clear error |
| COM error numbers and `TypeName` | ✅ | A missing ProgID reports 429 (`ActiveX component can't create object`), a failed attach reports 432, and `TypeName(obj)` returns the coclass name (`FileSystemObject`, `SWbemServicesEx`) |
| Automatic default member | ✅ | `Attribute X.VB_UserMemId = 0`; `x = obj` reads the default property, `Set x = obj` keeps the reference |
| `ByRef` write-back for object methods | ✅ | Method parameters go through the `Ref` protocol (see §3.7) |
| Method overriding (virtual methods) | ✅ | Overridable by default, resolved through the MRO |
| Writing an undeclared member | ✅ | Raises 438 (the `__setattr__` allowlist) |
| `Class Child Inherits Parent` | ✅ | Superset extension; method overriding, field/property inheritance, `_ovpp_attrs` merge |

### 2.7 Type system (OpenVBS++ extension)

| Feature | Status | Notes |
|---|---|---|
| `Option Typing` | ✅ | Compile-time checks when enabled; annotations are ignored otherwise |
| `x -> Integer` / parameter / return annotations | ✅ | `->` does not conflict with the colon |
| Assignment / argument / return type checks | ✅ | Conservative; `Variant` matches anything |
| Array element type `Dim a(5) -> Integer` | ✅ | Element type checked on read and write |
| Class field and property types | ✅ | Includes inherited fields and `Property Let` parameter types |

### 2.7.1 `Option Native` (OpenVBS++ extension)

`Option Native` turns off VBScript's arithmetic coercion: a text operand raises **error 13**
instead of being converted. `Option Typing` implies it, because typed code should not convert
silently either. Without the option nothing changes, so this can never break VBScript code.

| Form | Effect |
|---|---|
| `Option Native` | Native arithmetic (same as `Option Native On`) |
| `Option Native On` | Explicit form |
| `Option Native Off` | Explicitly back to VBScript coercion |
| `Option Typing` | Enables type checking **and** native arithmetic |

| Operator | Default (VBScript) | `Option Native` |
|---|---|---|
| `+` | `"1" + "2"` → `"12"`, `1 + "2"` → `3` | both raise 13 |
| `- * / ^` | `"6" - "4"` → `2`, `"2" * 3` → `6` | both raise 13 |
| `\` `Mod` | `"6" \ 2` → `3`, `"6" Mod 4` → `2` | both raise 13 |
| unary `+` `-` | `-"3"` → `-3` | raises 13 |
| `&` | `"a" & 1` → `"a1"` | unchanged (it *is* a text operation) |
| comparisons | see §4 | unchanged |
| `Empty` | `Empty + 1` → `1` | unchanged (Empty is the default value, not a converted one) |
| `Null` | `Null + 1` → `Null` | unchanged (propagation, not conversion) |

`Option` statements describe the whole script, so they must come before any other statement —
a late `Option Native Off` would otherwise silently change how the code above it compiles.
`Option Explicit` has the same rule in VBScript itself; the error is 1002
(`'Option Native' must appear before any other statement`).

#### Why `Native` and not `Strict`

The keyword was called `Option Strict` until v0.9.26. The rename is worth explaining, because
a keyword has to carry three things at once: what it enforces, what it buys you, and what it
must not mislead anyone about.

- **What it enforces.** "Strict" says *that* something is strict without saying *what*.
  "Native" says it precisely: only native numeric types take part in arithmetic.
- **What it buys you.** This option exists for speed as much as for discipline. VBScript
  semantics force every `+ - * /` through a runtime function that inspects both operand types
  (roughly 60 ns each). Turning this option on is a promise — "both operands are numbers" —
  and the compiler pays you back by emitting **native** operations. The word `Native` states
  the trade; `Strict` only advertised the cost, so people left it off.
- **What it must not mislead about.** In VB6 / VBA / VB.NET, `Option Strict On` is the
  familiar spelling for *compile-time* checks: no implicit narrowing conversions, no late
  binding, `Dim x As T` required. This project already has that switch under a different
  name (`Option Typing`); this one is about operator behaviour at run time. Reusing the
  familiar word with a different meaning is how you mislead a migrating developer — someone
  arriving from VB6 would enable `Option Strict` expecting an error on the `Dim` line and
  instead meet error 13 at `"1" + "2"`.

The name also leaves the option room to grow. If the comparison gap in §4.1 is ever closed,
"native comparisons" is a natural extension of the same word, whereas `Native Arithmetic`
would have pinned the keyword to one operator group.

`Option Strict` is not accepted as an alias: the option was renamed before it was ever
released, so no script in the wild can depend on the old spelling.

### 2.8 Structured exceptions (OpenVBS++ extension)

| Feature | Status | Notes |
|---|---|---|
| Try ... End Try | ✅ | Block scope; an error inside the block is handed to Catch |
| Catch e | ✅ | e is an error snapshot (Number / Description / Source / Line) |
| Catch e When condition | ✅ | Filters by error number; several When clauses match in order and a catch-all may follow |
| Finally | ✅ | Runs on normal exit, `Exit Sub`, error propagation and re-throw |
| Err.Line | ✅ | Extension field: the failing line number (VBS has none) |
| Coexists with On Error | ✅ | Try wins; the outer On Error state is restored when the block is left |

Key point: Try / Catch / Finally / When are **contextual keywords** — recognised only at the
start of a statement (and `Try` must be on its own line), so existing VBS code such as
`Try = 1`, `Catch = 2` stays legal (covered by tests).

### 2.8.1 Multiple files (OpenVBS++ extension)

| Form | Semantics |
|---|---|
| `<script src="lib.vbs">` | **Merge**: the target file's procedures/classes join the current global scope and can be called directly |
| `<namespace src="lib.vbs" alias="u">` | **Namespace**: `u.Double(21)`; only `Public` procedures are exported |
| `<alias>` omitted | The file name (without extension) becomes the alias |
| `nickyname` / `tar` | Synonym attributes for `alias` |

Rules: paths resolve relative to **the declaring file**; the same file included by several
`<script>` tags is merged once; circular includes fail at compile time with the chain
printed; every file is compiled separately (line numbers and file names stay correct); if a
`<namespace>` file itself uses `<namespace>`, the referenced side must already be bound (see
the limitations).

### 2.9 Python interop (OpenVBS++ extension)

| Feature | Status | Notes |
|---|---|---|
| `#python ... #endpython` | ✅ | Preprocessing extracts the block, leaves a placeholder statement in place — **line numbers do not move** |
| `COP` object pool | ✅ | `COP` is a global dictionary; on the Python side `from OpenVBSPP.ovpprt.cop import COP` yields the same object |

Since v0.9.20 Python interop must be imported explicitly: put
`<script src="python-embed" />` at the top of the script. The `python-embed` package ships
with OpenVBS++, and `Option COP` (the previous switch) has been removed — writing it is now
a syntax error whose message points at that import.

### 2.10 Built-in functions

Every VBScript built-in function is available except `LoadPicture`, which exists only to
fail with a clear message (§4); every built-in statement (`Execute`, `ExecuteGlobal`) is
available as well.

- strings: `Len Left Right Mid UCase LCase LTrim RTrim Trim StrReverse Space String Asc Chr InStr InStrRev Replace Split Join Filter StrComp Hex Oct`
- conversion: `CStr CInt CLng CDbl CBool CByte CCur CSng CDate`
- math: `Abs Int Fix Sgn Sqr Round Rnd Randomize Sin Cos Tan Atn Exp Log`
- arrays: `Array UBound LBound`
- tests: `IsNumeric IsEmpty IsNull IsObject IsArray IsDate IsMissing TypeName VarType`
- dates: `Now Date Time Timer Year Month Day Hour Minute Second Weekday DateAdd DateDiff DatePart DateSerial DateValue TimeSerial TimeValue MonthName WeekdayName`
- formatting: `Format FormatNumber FormatCurrency FormatDateTime FormatPercent`
- runtime information: `RGB ScriptEngine ScriptEngineMajorVersion ScriptEngineMinorVersion ScriptEngineBuildVersion GetLocale SetLocale`
- interaction: `MsgBox InputBox` (real dialogs when interactive, console stand-ins otherwise, see §3.10)

Also supported since v0.9.21: `#date#` literals (for example `#2020-05-06#` and
`#1/2/2020 3:04:05 PM#`) and the `&H` / `&O` numeric literals.

| Function | Status | Notes |
|---|---|---|
| `Execute` / `ExecuteGlobal` | ✅ | Statement form; the code runs in the global scope |
| `Eval` | ✅ | Expression form |
| `GetRef` | ✅ | Both `f(21)` and `f()` invoke the reference; a bare `f` reads the object |
| `LoadPicture` | ⚠️ | Raises `OVPPObjectNotSupportedError`: there is no GUI picture library |

## 3. Known semantic differences

These are places where we **deliberately** differ from VBS, or cannot match it exactly.

### 3.1 Logical operators are bitwise

In VBScript `And/Or/Not/Xor` are bit operations on integers and logical operations on
booleans. True is -1 and False is 0, so the two semantics unify naturally:

```vb
echo 5 And 3        ' 1 (bitwise)
echo Not 0          ' -1 (bitwise)
echo (1 = 1) And (2 = 2)   ' True (both sides boolean, stays boolean)
echo True And 5     ' 5 (mixed: bitwise)
```

We implement **exactly this rule**, so `echo -1 And 0` prints `0`, not `False`.

### 3.2 String comparison is case-sensitive

VBS defaults to `Option Compare Binary`, where `"A" = "a"` is false — we match that.
`Option Compare Text` (case-insensitive) is not implemented (see §4).

### 3.3 Error numbers

Error classes reuse the VBS numbering (11 division by zero, 13 type mismatch, 35 procedure
undefined, 94 invalid use of Null, 438 object doesn't support this property or method, 450
wrong number of arguments, 1002 syntax error, …), except that "variable not defined" uses
1001 instead of VBS's 500. That is a long-standing project convention; changing it would
break existing scripts that inspect `Err.Number`.

That error only happens under `Option Explicit`, exactly as in VBS. Without it, reading a
name that was never declared or assigned yields `Empty` (`TypeName(z) = "Empty"`,
`z & "x" = "x"`), and the compiler leaves no runtime check behind — the name is initialised
once at the top of the module, so a tight loop pays nothing for it.

### 3.4 Number display

`Echo` and `&` format numbers the VBS way: `3.0` → `3`, `-0.0` → `0`. Python's native `3.0`
is never printed.

A non-integral Double keeps all of Python's significant digits instead of VBS's fifteen:
`echo 1 / 3` prints `0.3333333333333333` where VBS prints `0.333333333333333`. The value in
memory is the same Double; only the trailing digits of the text differ.

### 3.5 Negative array indices are not checked

`a(-1)` reads the last element in Python, while VBS reports "subscript out of range".

### 3.6 Nesting semantics of `Exit For`

If `Exit For` sits inside a `Do` loop, it only leaves the innermost loop (Python has no
named loops).

### 3.7 ByRef for object methods uses the uniform Ref protocol

In VBS, the ByRef parameter of `obj.Method x` writes back into `x`. Since v0.6 every object
method call compiles to:

```python
_ref_1 = Ref(x)
_ovpp_obj.call_method(obj, "method", [_ref_1])
x = _ref_1.v            # write-back
```

Inside the method body, ByRef parameters use `param.v` directly, while ByVal parameters are
unpacked into locals at the top of the body. Built-in objects (`Dictionary`, …) are not VBS
classes, and `call_method` automatically takes `.v` to call by value, so the same protocol
works for both sides.

### 3.8 The `Option Typing` type system is conservative

When enabled it only checks what can be decided: assignments, arguments, return values.
`Variant`, `Empty`, `Null` and `Object` match anything, so legal programs are never
false-positived. When disabled, annotations are ignored and behaviour is plain VBS.

### 3.9 `Resume` / `Resume Next` need an error handler

`Resume` (retry the failing statement) and `Resume Next` (continue with the next one) are
implemented since v0.9.5 (see section 6): a procedure that contains them compiles into a
statement-index state machine. Both only make sense inside a procedure that has an
`On Error GoTo <label>` handler, so using them without one is a compile-time error.

Do not confuse them with `On Error Resume Next`, which is a mode statement, is supported
everywhere, and does not need a handler.

### 3.10 `MsgBox` and `InputBox` show real dialogs when they can

VBScript's `MsgBox` opens a window and returns the button the user pressed; `InputBox`
opens a text field. OpenVBS++ does both through the bundled `ui-support` package, drawn
by Tkinter, and falls back to a console stand-in whenever a window would be wrong.

A dialog is shown only when all of these hold:

- the script was started by the `openvbs` command line (not by the IDE, and not by
  embedding `vbs_run` in a Python program);
- stdin and stdout are a terminal, and `CI` is not set;
- the `ui-support` package loaded and Tkinter can create a window.

With a window, the arguments mean what they mean in VBScript:

- `MsgBox text [, buttons] [, title]` becomes a message box with the requested button
  set (`vbOKOnly` … `vbRetryCancel`), icon, default button and modal flag, and returns
  the codes `1`–`7` (`vbOK`, `vbCancel`, `vbAbort`, `vbRetry`, `vbIgnore`, `vbYes`,
  `vbNo`).
- `InputBox(prompt [, title] [, default])` becomes a text dialog with `default`
  pre-filled; pressing Cancel or closing the window returns `""`.

Everywhere else the console behaviour applies: `MsgBox` prints the text and returns `1`,
`InputBox` prints the prompt and reads one line from stdin. That is what happens in CI,
in pipes, and whenever the package or Tk is unavailable — a batch job never blocks on a
window.

`OPENVBS_UI` overrides the decision: `console` (or `off`) always uses the stand-ins,
`gui` (or `on`) asks for dialogs, and the default `auto` applies the rules above.
`openvbs --version` style runs and `python -m OpenVBSPP check` never open windows.

### 3.11 `Declare` and native code

`Declare Function/Sub` binds a DLL entry point. OpenVBS++ runs on the Python VM, so the
statement is rejected with a clear syntax error instead of silently doing nothing.

**Short-term decision (v0.9.22): not supported.** A `ctypes`-backed binding was
investigated and is technically feasible, but it lets a script call arbitrary native
code, and a wrong signature kills the process with an access violation that neither
`On Error Resume Next` nor `Try ... Catch` can intercept. That trade is not worth taking
yet, so `Declare` stays out of the language and no flag enables it. The
investigation is recorded as item 7c in
[../dev-docs/ROADMAP.md](../dev-docs/ROADMAP.md).

The supported routes to native code are:

- COM: `CreateObject("ProgID")`, on Windows, through the `com-support` package.
- Native Python: an `over` package, or a `#python` block (see section 2.9).

### 3.12 `End` and `Stop`

- A bare `End` on its own line stops the script immediately, exactly like VBScript's
  `End` statement. It compiles to a `ScriptEnd` exception that derives from
  `BaseException`, so neither `On Error Resume Next` nor `Try ... Catch` can swallow it;
  the runner treats it as a normal exit (process exit code 0). `End If`, `End Sub`,
  `End With`, `End Select`, `End Try`, `End Function`, `End Property` and `End Class`
  keep their block-closing meaning.
- `Stop` is a debugger breakpoint in VBScript. OpenVBS++ has no debugger to break into,
  so the statement is accepted and does nothing. It is not a syntax error, because real
  scripts leave `Stop` calls behind on purpose.

### 3.13 The `WScript` host object

Scripts started by `openvbs` get the same global `WScript` object that Windows Script
Host provides, so legacy code that reaches for it keeps working:

| Member | Behaviour |
|---|---|
| `WScript.Echo a, b` | Prints the values separated by spaces, like the `Echo` statement |
| `WScript.Arguments` | The arguments after the script path; `Count`, `Arguments(0)` (indexed from zero), `Item(n)` and `For Each` work. Out-of-range indexes raise error 9 |
| `WScript.Arguments.Named` | `/name:value` and `/name=value` arguments: `Exists(name)`, `Item(name)`, `Count` |
| `WScript.ScriptName`, `.ScriptFullName` | The file name and full path of the running script (`""` for `-c`/stdin code) |
| `WScript.StdOut`, `.StdErr`, `.StdIn` | `TextStream` objects: `Write`, `WriteLine`, `WriteBlankLines`, and on `StdIn` also `ReadLine`, `ReadAll`, `AtEndOfStream`, `Skip`, `SkipLine` (`ReadLine` at end of file raises error 62) |
| `WScript.Name`, `.Version` | `"Windows Script Host"` and `"5.812"`, the values old scripts test |
| `WScript.FullName`, `.Path` | The running interpreter and its directory (OpenVBS++ is a Python host, not `cscript.exe`) |
| `WScript.Interactive`, `.Timeout` | `Interactive` reports whether stdin and stdout are a terminal; `Timeout` is stored but not enforced |
| `WScript.Sleep ms` | Sleeps; `WScript.Quit code` exits with that code |
| `WScript.CreateObject`, `.GetObject` | Same as the global functions, with the COM error codes of section 2.6 |

`WScript` itself is read-only: assigning to it is a syntax error. Command-line arguments
reach it from the `openvbs script.vbs a b` form (or `vbs_run(..., args=[...])` when the
runtime is embedded).

## 4. Known gaps

`tests/test_vbs_superset.py` tracks the remaining gaps as `expectedFailure` cases, so closing
one turns the suite red on purpose until the case is moved into `VBS_FEATURES`. The earlier
gaps (object model, `Execute`/`Eval`, formatting functions, `Option Compare Text`,
`GetObject`, default members, arithmetic coercion, Null propagation) are closed.

### 4.1 Comparisons do not coerce

VBScript converts when it compares a number with numeric text. We compare by type, which
Python does:

| Expression | VBScript | OpenVBS++ |
|---|---|---|
| `1 = "1"` | `True` (numeric comparison) | `False` |
| `2.5 = "2.5"` | `True` | `False` |
| `"2" > 10` | `False` | error 13 |
| `"abc" = 5` | error 13 | `False` |
| `"abc" < "abd"` | text comparison | text comparison ✅ |
| `1 = 1`, `"A" = "a"` | as VBScript | ✅ |

Why it is still open: every comparison would have to go through a runtime helper, which costs
roughly 60 ns per comparison — a 1M-iteration loop would go from about 92 ms to about 150 ms.
Fixing it is a deliberate trade of speed for VBScript fidelity in a corner that mostly shows up
in scripts that compare an `InputBox` string with a number.

### 4.2 Unsupported or declined

| Item | Explanation |
|---|---|
| COM host objects (`Excel.Application`, `WScript.Shell`, …) | Supported on Windows since v0.9.14 through the `com-support` package; on other platforms they raise `OVPPObjectNotSupportedError` instead of simulating COM |
| The `WScript` host object | Provided since v0.9.23 (§3.13): `Echo`, `Arguments` (indexed from zero, plus `Named`), `ScriptName`, `StdOut`/`StdErr`/`StdIn`, `Sleep`, `Timeout`, `Quit`, `CreateObject` and `GetObject`. It describes the OpenVBS++ console host, so `FullName`/`Path` point at the Python interpreter rather than `cscript.exe`, and `Interactive` follows the terminal |
| The `Collection` built-in class | VBScript itself has none (it belongs to VBA) |
| Other `Attribute` properties (`VB_Name` etc.) | Only `VB_UserMemId` is parsed |
| `MsgBox` / `InputBox` | Real dialogs through the bundled `ui-support` package when the run is interactive (see §3.10); console stand-ins in CI, pipes and embedded runs, where `MsgBox` prints its text and returns 1 and `InputBox` reads one line from stdin |
| `LoadPicture` | The function exists so old scripts fail with a clear message: there is no GUI picture library (`OVPPObjectNotSupportedError`) |
| Procedures defined at run time by `ExecuteGlobal` | They can update variables but cannot be called by name, because procedure names bind at compile time. Use `GetRef` on an existing procedure, or put the code in its own module |
| Backward `GoTo` | Only forward jumps are compiled; a jump to an earlier label raises `Recursive GoTo ... is not supported` |
| `Declare` (DLL declaration) | Deliberately unsupported for now: the runtime is Python, not a DLL host, and a mistyped native call can terminate the process uncatchably (see §3.11). Use `CreateObject` for COM or an over package for native code |

## 5. Change checklist

Run through this before and after touching the compiler:

1. Does the new syntax change the behaviour of some legal VBS snippet? If so, change the
   implementation or document it in §3.
2. Did `tests/test_vbs_superset.py` gain the matching case?
3. Was a gap closed? Remove it from the §4 table, add the case to `VBS_FEATURES`, and update
   this file's status sections.
4. Is the full test suite (`python -m unittest discover -s tests`) still green?
5. Are the support matrices in README and this file in sync?

> The full object-model inventory (including keywords that do not apply) is in
> [OBJECT_MODEL.md](OBJECT_MODEL.md) and [OOP.md](OOP.md).

## 6. Status update (v0.9.5 - v0.9.26)

Resume / Resume Next are implemented: a procedure containing Resume compiles into a
"statement index dispatch" state machine, one `_pc` branch per statement; on error the
failing index is recorded and control jumps to the `On Error GoTo` handler, from where
`Resume` returns to the failing statement and `Resume Next` advances one statement. This
works inside nested If / loops as well. Procedures without Resume keep the original
sequential compilation, with unchanged behaviour and performance.

After the v0.9.21 audit every native VBScript statement form is supported, with one
deliberate exception: `Declare`. As of v0.9.22 `MsgBox` and `InputBox` show real Tk
dialogs in interactive command-line runs (see §3.10), and as of v0.9.23 the `WScript`
host object is provided (§3.13), so the remaining gaps are host details rather than
syntax: `LoadPicture` reports that there is no GUI picture library, procedures defined at
run time by `ExecuteGlobal` cannot be called by name (use `GetRef`), and only forward
`GoTo` is allowed. The section 4 table above is the complete list.

### 6.1 Value semantics closed in v0.9.24

Three value-level differences found by the v0.9.23 audit were fixed against `cscript`
ground truth, so the arithmetic, Null and undeclared-variable behaviour now matches
VBScript rather than Python:

| Rule | Before | Now |
|---|---|---|
| `+ - * / ^` coercion | Mapped straight onto Python operators, so `"1" + "2"` was `"12"` (VBS: also `"12"`), but `"2" * 3` was `"222"` and `1 + "2"` was a type mismatch | `+` concatenates only when both sides are text; `- * / ^` convert text to numbers; `/` and `^` always yield a Double; a non-numeric string raises 13 |
| Null through built-ins | `Len(Null)` was 4, `TypeName(Null)` was `"NullType"`, `UCase(Null)` was `"NULL"` | The functions that VBS propagates Null through (`Len`, `UCase`, `Abs`, `Trim`, date parts, …) return Null; the conversion/math group (`CStr`, `CInt`, `Sgn`, `Sqr`, …) raises 94; the `Format*` group raises 13. `&` treats Null as empty text, except `Null & Null`, which stays Null |
| Undeclared variable read | Raised "variable is undefined (1001)" | Reads as `Empty` unless `Option Explicit` is on (§3.3). The compiler initialises such names once at module top, so nothing is checked at run time |

### 6.2 `Option Native` added in v0.9.26

VBScript's arithmetic coerces text to numbers. That is correct for VBScript, but it also turns
typos into silent results (`"2" * 3` used to be `"222"` before v0.9.24). `Option Native` is the
opt-in answer: it keeps VBScript's default behaviour while letting a project demand numbers, so
`1 + "2"` raises 13 instead of becoming `3`. `Option Typing` switches it on, because type
annotations and silent coercion contradict each other. See §2.7.1 for the full operator table.

The same release records one open fidelity gap: comparisons still compare by type instead of
coercing (§4.1).
