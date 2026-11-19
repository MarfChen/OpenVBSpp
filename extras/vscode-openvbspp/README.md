# OpenVBS++ for VS Code

Minimal, dependency-free VS Code support for [OpenVBS++](../../README.md):

- TextMate grammar for VBScript **plus** the OpenVBS++ superset (`Class ... Inherits`,
  `Try / Catch / Finally`, `Option Typing`, `<script>` / `<namespace>` tags, `#python`
  blocks, `?.`, `??`, `$"interpolation"`)
- comment / bracket / auto-indent configuration
- snippets for `class`, `sub`, `func`, `try`, `onerror`, `for`, `foreach`, `select`,
  `namespace`, `script`
- **Run** the current file (`Ctrl+F5`) and **check** it on save, by calling the real CLI:

```bash
openvbs run   --json script.vbs
openvbs check --json script.vbs
```

Diagnostics are real compiler diagnostics (line, column, message, error code), not heuristic
guesses. There is **no IntelliSense / autocomplete / go-to-definition yet** — that needs a
language server, tracked in [doc/IDE_LSP_DESIGN.md](../../doc/IDE_LSP_DESIGN.md).

## Install (from this folder)

```bash
# 1) package it (needs Node.js; npx downloads vsce on first use)
cd extras/vscode-openvbspp
npx --yes @vscode/vsce package        # -> openvbspp-0.1.0.vsix

# 2) install the produced .vsix
code --install-extension openvbspp-0.1.0.vsix
```

For a quick trial without packaging, copy this folder into
`%USERPROFILE%\.vscode\extensions\openvbspp-0.1.0\` and restart VS Code.

## Settings

| Setting | Default | Meaning |
|---|---|---|
| `openvbs.executable` | `openvbs` | CLI to invoke. Use a full path if it is not on `PATH` |
| `openvbs.arguments` | `[]` | Extra arguments inserted before the subcommand |
| `openvbs.checkOnSave` | `true` | Run `check --json` on open/save and publish diagnostics |

If `openvbs.exe` is blocked on your machine (see [doc/RELEASE.md](../../doc/RELEASE.md)),
point the executable at Python and add the module arguments:

```jsonc
{
  "openvbs.executable": "C:\\Path\\To\\python.exe",
  "openvbs.arguments": ["-m", "OpenVBSPP"]
}
```

## Notes

- The extension claims `.vbs`. VS Code ships no built-in VBScript language, so this is
  additive; switch the file's language if you want plain-text editing.
- The grammar is hand-written against the feature table in
  [README.md](../../README.md#what-you-get-beyond-plain-vbscript). If a construct is
  mis-coloured, that is a grammar bug, not a language bug — issues welcome.
