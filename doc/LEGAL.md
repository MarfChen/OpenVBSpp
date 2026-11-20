# Legal review: how this project relates to Microsoft VBScript

> This is a **self-check record** kept by the maintainer, not a legal opinion from a lawyer.
> For commercial releases or internal compliance reviews, have a professional double-check it.
> The conclusion is ordered from lowest to highest risk, with the handling for each item.

Last reviewed: v0.9.20 (2026-11-20).

---

## 1. The one-line conclusion

**Low risk.** This repository is an **independent implementation** written from scratch: it
contains no Microsoft source code, no verbatim Microsoft documentation, and distributes no
Microsoft binaries; `VBScript` is only mentioned as the *compatibility target* (nominative
use). What actually deserves care is not "may we build this" but "do not present ourselves
as Microsoft's product" — that is already stated in `README.md` / `README.zh-CN.md` / `NOTICE`.

---

## 2. What the repository actually contains (item by item)

| Item | Current state | Verdict |
|---|---|---|
| Microsoft VBScript engine source code | **None.** The compiler is `lexer.py` → `parser.py` → `transformer.py` → `compiler.py`, all original Python | Safe |
| Microsoft documentation text (HTML/PDF) | **None.** Reference material that used to live in `doc/` was moved out of the repository (moved to a local folder outside this repository, not in version control, not in the wheel/sdist) | Safe |
| `doc/OpenVBS spec.docx` | Our own language specification (v0.1.4, original Chinese prose), not a translation of a Microsoft specification | Safe |
| Microsoft executables / DLLs | **None.** `dist/openvbs.exe` is packaged by PyInstaller from this repository's code plus the Python runtime | Safe |
| Trademarks | `VBScript` is a trademark of Microsoft Corporation; it is used only to name the compatibility target | Nominative use, see §3 |
| Third-party code | Zero runtime dependencies; optional `pywin32` (PSF-2.0), declared in `NOTICE` | Compliant |

Self-check command (re-run after changing code or docs):

```powershell
# Expect hits only in the disclaimer wording, plus the CSS font name Microsoft YaHei
Select-String -Path README.md,README.zh-CN.md,NOTICE -Pattern "Microsoft"
```

---

## 3. Risk analysis item by item

### 3.1 Implementing a language compatible with VBScript (the core question)

- **Copyright**: language syntax, keywords and statement structure are *ideas / function*,
  which copyright does not protect; what is protected is the **specific expression**
  (source code, documentation prose, comments). We copied none of Microsoft's expression,
  and this can be checked file by file (our parser is a hand-written recursive-descent
  parser with no mapping to Microsoft's lexical or grammar tables).
- **Interoperability**: independently implementing an interpreter so that existing `.vbs`
  scripts keep running is a long-standing industry practice (Wine for the Windows API,
  Mono for .NET, and every BASIC-compatible implementation). Even in the jurisdictions
  with the strongest copyright protection, independent reimplementation for
  interoperability has clear room.
- **Patents**: VBScript shipped in 1996, so any related patent is long past its 20-year
  term. Low risk.

### 3.2 Trademark: how to use `VBScript` safely

Acceptable (nominative / descriptive use):

- "a scripting language compatible with VBScript", "built on VBScript syntax style",
  "migrate existing `.vbs` scripts"
- Metadata such as `keywords = ["vbs", "vbscript", "scripting"]` (for search)

Do not do (these confuse users or imply endorsement):

- Name the product `VBScript++`, `Microsoft VBScript Runtime` or `OpenVBScript`
  (too close to the official names);
- Use Microsoft's logo, icons, the `cscript.exe` / `wscript.exe` icons, or UI screenshots;
- Claim "officially supported by Microsoft", "in partnership with Microsoft", or
  "the Microsoft-approved successor to VBScript";
- Put "Microsoft VBScript" in large type at the top of the README in a way that makes the
  project look like a Microsoft project.

What this project does today: the product is `OpenVBS++` / package `openvbspp` (neither
contains the full word `VBScript`), the front page carries a bilingual disclaimer, and
`NOTICE` plus the README licence section restate the trademark ownership.

### 3.3 Documentation: how specific may it be

- Describing language behaviour ("`Exit For` does not add the step") is **fact** and may be
  written freely.
- Do not translate or copy Microsoft's help pages wholesale; write your own examples.
- All of this project's documentation is original prose plus our own examples, checked to
  contain no copied paragraph runs.

### 3.4 Third-party dependencies and packaging

- Zero runtime dependencies (standard library only), as stated in the README and `NOTICE`.
- Optional `pywin32`: PSF-2.0, redistribution allowed; the repository/wheel/sdist do **not**
  contain it — only the single-file exe on GitHub Releases bundles it, and `NOTICE` already
  says the licence travels with the exe.
- If new third-party code is added in future (over packages included), add a paragraph in
  the same format as `NOTICE`.

### 3.5 Where the user's responsibility starts

This project executes the user's `.vbs` scripts and lets `<% %>` embed Python — the same
trust model as `cscript.exe`: **the script itself can execute arbitrary code**. That is a
product design decision, not a legal defect, and it is stated in the README security note
and in the "how to think about the trust model" section of `SECURITY.md` so users are not
misled.

---

## 4. Compliance actions already done (checklist)

- [x] Microsoft copyright documentation moved out of the repository
      (`doc/微软官方vbscript参考手册*` → a local folder outside the repository)
- [x] `LICENSE` (Apache-2.0) + `NOTICE` (third-party components and trademark ownership)
- [x] Bilingual disclaimer on the README front page (independent implementation, no
      affiliation, trademark ownership, no Microsoft code or documentation)
- [x] Repository-wide secret/credential scan (0 hits)
- [x] `git` history contains no Microsoft documentation and no large generated artifacts
      (`.gitignore` / `.gitattributes` configured)
- [x] Release notes (`dev-docs/RELEASE.md`) and the README changelog tell the same story

## 5. Worth re-checking before publishing

1. Confirm that `git log --stat` shows no "Microsoft documentation / scraped pages /
   reference manual" ever entered the commit history (this repository's history starts as a
   clean snapshot at `1baf3f1`; `git log --diff-filter=A --name-only` is enough to verify).
2. Confirm the GitHub Release ships the exe only with the "COM included" explanation, and
   never claims to be a Microsoft runtime.
3. When promoting on Chinese platforms (Zhihu / Juejin / V2EX), avoid phrasing like
   "replaces the official Microsoft VBScript"; say instead "Microsoft has announced the
   deprecation of VBScript, and here is an open-source implementation that keeps existing
   scripts running".
