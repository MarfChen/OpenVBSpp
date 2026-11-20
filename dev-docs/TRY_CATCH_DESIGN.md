# Try / Catch design report · structured exceptions with a VBS accent

> English | **[中文](TRY_CATCH_DESIGN.zh-CN.md)**

> **Status: implemented in v0.8.0** (P0–P2). The implementation matches this
> design; deviations are recorded in §7 at the end of this file.
>
> **Abstract** — We *could* just add `Try … Catch … Finally … End Try`. But
> VBScript already owns a distinctive error model: `On Error Resume Next` plus a
> global `Err` object, and the "check `Err.Number` after the call" idiom is the
> language's fingerprint. This report proposes a design where structured
> exceptions and `On Error` are **the same machinery viewed from two sides**, so
> mixing them in one procedure stays predictable and VBS programs keep working
> untouched. The core idea: **do not invent a second exception system**; make
> `Try/Catch` the "local, scoped, nestable" spelling of `On Error`.

---

## 1. Problem

### Today

```vb
On Error Resume Next
DoWork
If Err.Number <> 0 Then
    echo "failed: " & Err.Description
    Err.Clear
End If
On Error GoTo 0
```

The mechanism works, but it has three structural defects (the ones every VBS
programmer has cursed):

| Defect | Consequence |
|---|---|
| **`On Error` is a process-wide switch** with no block scope | one `On Error Resume Next` stays in force until the procedure ends, swallowing every line in between |
| **No way back to "after the failure point"** | `Resume Next` can only skip; no retry / rollback / alternative strategy |
| **No stack when nested** | an inner procedure doing `GoTo 0` leaves the outer one unable to see the error |

### Goals

1. **Block-scoped** exception handling (on entering the block, restored on exit);
2. **`Finally`** support (resource cleanup is mandatory);
3. **Re-raise** and **exception chaining** (an `InnerException`-style cause chain);
4. **Coexist with `On Error`**, with predictable behaviour when both are mixed in
   one procedure;
5. **Pure superset**: no new semantics for existing VBS code; old scripts change
   by not a single byte.

---

## 2. Design: three layers, one runtime

The proposal is not "add a try statement" but to split error handling into
**three layers** sharing one runtime error stack, `_ovpp_err`:

```
layer 3  Try / Catch / Finally            block scope, nestable, re-raisable   <- new syntax
layer 2  On Error Resume Next / GoTo 0    procedure-wide switch (native VBS)   <- unchanged
layer 1  Err object + error stack         shared runtime plumbing              <- extended
```

Key point: **layers 2 and 3 compile into the same `try/except` structure**; only
"who catches it" differs. Mixing them stays predictable, and we do not maintain
two sets of semantics.

### 2.1 Syntax (VBS-flavoured choices)

```vb
Try
    DoWork
Catch e
    echo "error: " & e.Number & " " & e.Description
    echo "source: " & e.Source
Catch e When Err.Number = 53        ' optional filter clause (VBS-style keyword When)
    echo "file not found"
Finally
    CloseAll
End Try
```

Trade-offs (why it does not look like C# / Java):

| Choice | Reason |
|---|---|
| `Catch e` rather than `Catch (e)` | VBS procedure calls do not use parentheses anyway |
| `When` filter clause | `Err.Number` is numeric, so numeric filtering reads more naturally than type matching; `When` is already a VBS word (used only in `Select Case`, so no conflict) |
| The caught object is a **snapshot of `Err`** | introduces no new exception type system; `e.Number` / `.Description` / `.Source` are identical to `Err`, so there is nothing to learn |
| `End Try` rather than `EndTry` | aligns with `End If` / `End Sub` |
| No `Throw` keyword; use `Err.Raise` | **deliberately adds no keyword**: `Err.Raise` already is VBS's way to re-raise, and `Throw x` would only be sugar |

### 2.2 Relationship with `On Error` (the most important section)

**One rule: `On Error` sets "the procedure default", `Try` overrides "inside the
block".**

| Scenario | Behaviour |
|---|---|
| `On Error Resume Next` in the procedure, outside the block | as before: the error is captured and execution continues |
| Same, entering a `Try` block | while `Try` is active the error is **no longer swallowed by Resume Next**; it goes to `Catch` |
| Leaving `Try` (normal or exceptional path) | the `On Error` state from before the block is restored (block scope) |
| No `When` clause matches in `Catch` | keep propagating: outer `Try`, then the procedure's `On Error`, then the caller |
| Re-raising inside `Catch` (`Err.Raise`) | follows the same chain; `Finally` still runs |
| Procedure with `On Error` but no `Try` | **exactly as today** (zero behaviour change) |

Implementation-wise this is an "error handling stack": entering `Try` pushes a
handler, leaving pops it; `_ovpp_err.dispatch()` looks at the top of the stack
and, when empty, falls back to the old `On Error` logic. The old path never
touches the new one, so the superset property holds automatically.

### 2.3 `Finally` semantics

- **Always runs**: normal exit, `Exit Sub/Function`, `Exit Do/For`, error exit,
  re-raise;
- an error raised in `Finally` **replaces** the one propagating (as in .NET); the
  old error is attached to `Err.InnerException`;
- `Exit Sub` inside `Finally` is allowed but skips the remaining statements of
  the `Finally` block.

### 2.4 Error object and chain

Keep the four `Err` fields and add two **read-only** ones:

| Field | Meaning |
|---|---|
| `Err.Number` | error number (negative = COM-style, positive = VBS-style, as today) |
| `Err.Description` | description |
| `Err.Source` | where it was raised (procedure name, filled automatically) |
| `Err.Line` | **new**: the line number that failed (previously only printed to the terminal) |
| `Err.InnerException` | **new**: the previously replaced error (forming a chain) |

`Err.Line` is purely additive: VBS programs never read it, but the debugger, LSP
and `Catch` can.

### 2.5 Generated code (sketch)

```vb
Try
    DoWork
Catch e
    Handle e
End Try
```

compiles to (simplified):

```python
_handler = _ovpp_err.push_handler()
_ovpp_err.begin_try(_handler)
try:
    _ovpp_sub_dowork()
except BaseException as _exc:
    if not _ovpp_err.handle(_handler, _exc):
        raise
    e = _ovpp_err.current_snapshot()
    _ovpp_sub_handle(e)
finally:
    _ovpp_err.pop_handler(_handler)
```

The point: `handle()` owns the `When` filtering and the "nobody wants it, keep
propagating" decision, and `pop_handler` restores the `On Error` state — so
**block scoping is free**, no compiler state analysis needed.

---

## 3. Why not other approaches

| Approach | Rejected because |
|---|---|
| `On Error Resume Next` plus immediate checks (today) | no block scope, no `Finally`, no retry |
| Copy C#: `try/catch/finally` | loses the VBS `Err` mental model; two error object systems confuse newcomers |
| Copy Python: `raise/try/except` | needs new `raise` / `except` keywords and duplicates `Err.Raise` |
| Retry via `Resume` (VBA style) | VBA's `Resume label` needs line numbers or labels; this language has no line-number statements, so cost exceeds benefit |
| `Catch` without `Finally` | resource cleanup has nowhere to live, which invites leaks |

---

## 4. Superset compatibility argument

| Possible overlap | Conclusion |
|---|---|
| Are `Try` / `Catch` / `Finally` / `When` VBS keywords? | **No.** They are contextual keywords recognised only at the start of a statement; `Try = 1` as a variable name remains legal (our lexer keeps no keyword set anyway) |
| Do old scripts contain labels like `Try:`? | No (VBS has no user labels) |
| Does `End Try` conflict with the `End` statement? | No; `End` followed by `Try` is the only form, and bare `End` is unchanged |
| Code that never uses `Try` | zero change (the new syntax never participates in compilation) |

One more hard requirement: **every case in `tests/test_vbs_superset.py` must keep
passing after the feature lands**, and the new `Try` cases live in their own
group; existing assertions may not be modified.

---

## 5. Implementation order (suggested)

| Phase | Content | Risk |
|---|---|---|
| P0 | Runtime: `push_handler/pop_handler/handle/snapshot` plus `Err.Line` | low (purely additive) |
| P1 | Syntax: `Try … Catch e [When expr] … Finally … End Try`, single level | medium (parser block structure) |
| P2 | Nesting, re-raise, `InnerException` chain | medium |
| P3 | `Exit Sub` crossing `Finally`, the `On Error` mixing matrix | high (combinations must be enumerated) |
| P4 | LSP / debugger integration (`Err.Line` plus stack frames) | low |

For P3 the test matrix should be written out: `{normal, error in block, error in
Catch, ExitSub in block}` × `{Finally present/absent}` × `{outer On Error on/off}`
— 16 combinations, all asserted.

---

## 6. Open questions (decided during implementation)

1. **May `Catch` have several clauses without `When`?** Preference: only one
   catch-all `Catch`, and it must be last (same rule as `Select Case Else`).
2. **Is `Exit Sub` allowed inside `Finally`?** Preference: allow it, but document
   that it skips the remaining `Finally` statements.
3. **A one-line shorthand for `Try … Catch`?** Preference: **no**; exception
   handling should be explicit.
4. **Error number space**: should negative numbers passed to `Err.Raise` align
   with `_ERROR_MAP`? Preference: keep as is, document only.

---

## 7. Implementation record (v0.8.0)

| Design item | Implementation | Notes |
|---|---|---|
| Error handling stack | ✅ `push_handler` / `pop_handler` / `handler_active` in `ovpprt/err.py` | nesting supported |
| `Try` overrides `On Error Resume Next` | ✅ `capture()` re-raises when the stack is non-empty | one line settles the precedence |
| `Catch e` snapshot | ✅ `_Snapshot` (Number / Description / Source / **Line**) | maps one-to-one onto `Err` |
| `When` filtering | ✅ compiled into an if/elif chain; a miss re-raises | `e.Number` is available inside `When` |
| `Finally` | ✅ Python `try/finally` plus `pop_handler()` | `Exit Sub` passes through automatically |
| Re-raise | ✅ still `Err.Raise` (no new `Throw`) | matches the design |
| `Err.Line` | ✅ deepest non-package frame from the traceback | generated code carries VBS line numbers, so it works directly |

**Deviations from the design**: none functionally. The four open questions were
implemented with the suggested defaults — only the last `Catch` may omit `When`
(anything else is a parse error), `Exit Sub` is allowed in `Finally`, no
one-line shorthand, error numbers unchanged.

**New tests**: `tests/test_try_catch.py` (14 cases), including two contextual
keyword cases proving `Try` / `Catch` / `Finally` still work as ordinary
variable names with VBS semantics.
