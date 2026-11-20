# Good first issues

> English | **[中文](good-first-issues.zh-CN.md)**

These are scoped, self-contained tasks. Pick one, open an issue saying you are
on it, and send a PR. Ask questions early — there are no silly questions. Each
task names the file to touch and the test to run.

---

## GFI-1 · The `RegExp` object

VBScript ships a `RegExp` object (`New RegExp`, `Pattern`, `Global`,
`IgnoreCase`, `Test`, `Replace`, `Execute`). We only have `Filter`, which uses a
regular expression internally. Implement `RegExp` as a runtime class.

| Item | Detail |
|---|---|
| Where | New `OpenVBSPP/ovpprt/regexp.py`; add the ProgID `vbscript.regexp` in `_transform_create_object` in `transformer.py` |
| Reference | `Dictionary` in `ovpprt/collections.py` (also a built-in object) |
| Tests | New `tests/test_regexp.py`, plus two cases in `test_vbs_superset.py` |
| Difficulty | medium |

---

## GFI-2 · Multi-file `Include`

Add `Include "path.vbs"` that inlines another script at compile time (a
lex-level splice before parsing, so line numbers of the main file stay sane).
Decide and document: path resolution, cycles, and interaction with `Option`
statements.

| Item | Detail |
|---|---|
| Where | New `OpenVBSPP/plugins/include.py` (model it on `plugins/python_embed.py`), wired up in `compiler.py` |
| Pitfall | `plugins/python_embed.py` keeps line numbers with "placeholder + pad newlines"; reuse that idea |
| Tests | `tests/test_include.py` |
| Difficulty | medium |

---

## GFI-3 · `IsMissing` for `Optional` parameters

`Optional` works (defaults are honoured), but a function cannot tell whether an
argument was supplied. Add `IsMissing(x)` and document that it only applies to
`Optional` parameters.

| Item | Detail |
|---|---|
| Where | `OpenVBSPP/ovpprt/builtins.py` (function plus registry entry); in `transformer.py` switch the default of `Optional` parameters from `EMPTY` to a dedicated sentinel |
| Tests | Two cases in `tests/test_optional_paramarray.py` |
| Difficulty | small |

---

## GFI-4 · Better error messages

Runtime errors currently read like `Type mismatch (13): ...`. Add a hint line
for the five most common mistakes (calling a `Sub` as a function, missing `Set`
for objects, array index out of range, wrong argument count, assigning to a
constant).

| Item | Detail |
|---|---|
| Where | `_ERROR_MAP` / `wrap_runtime_error` in `OpenVBSPP/runtime.py` |
| Requirement | hints must be English (error messages are English today) |
| Tests | New `tests/test_errors.py`, asserting that `"hint"` appears in the message |
| Difficulty | small |

---

## GFI-5 · Fill in the deviation table in `doc/VBS_SUPERSET.md`

`doc/VBS_SUPERSET.md` §3 lists where we deviate from VBS. Several rows are one
line long but deserve a minimal repro plus an explanation. Add a repro snippet
and a short "why we chose this" note for each row.

| Item | Detail |
|---|---|
| Where | `doc/VBS_SUPERSET.md` §3 |
| Tests | none (docs only), but every repro snippet must actually run |
| Difficulty | small (a great first contribution) |

---

## GFI-6 · Compile cache

`vbs_compile` re-parses on every run. Add an in-memory LRU cache keyed by source
hash, with an opt-out environment variable `OVPP_NO_CACHE=1`.

| Item | Detail |
|---|---|
| Where | `OpenVBSPP/compiler.py` |
| Careful | error behaviour must not change; `filename` belongs in the key |
| Tests | `tests/test_compiler_cache.py` |
| Difficulty | medium |

---

## How to claim

1. Open an issue titled `[GFI-x] I'll take this one`.
2. Fork, then branch as `feat/gfi-x-short-name`.
3. Send a PR following the checklist in `CONTRIBUTING.md`.
4. Sign off your commit (`Signed-off-by`, DCO).
