# ui-support — real MsgBox / InputBox dialogs for OpenVBS++

中文：[English](#english) | **[中文](#中文)**

This is the over package that gives OpenVBS++ **real GUI dialogs** for
`MsgBox` and `InputBox`, drawn by Python's `tkinter` (Tcl/Tk 8.6). Scripts do
not import anything: the runtime picks the package up by itself when a dialog
makes sense.

## English

- `over.json` declares `"capabilities": ["ui"]` and `"kind": "feature"`, so it
  is not a Python dependency of your project.
- Resolution order: `packages/ui-support/` in your project → the bundled copy
  in `OpenVBSPP/vendor/ui-support/`.
- A dialog is shown only when **all** of these hold: the script was started by
  the `openvbs` command line (not the IDE, not an embedded `vbs_run()` call),
  stdin/stdout are a terminal, `CI` is unset, and Tk can create a window.
  Everything else falls back to the console behaviour (`MsgBox` prints and
  returns 1, `InputBox` reads one line from stdin).
- `OPENVBS_UI` overrides the decision:
  - `auto` (default) — the rules above;
  - `console` / `off` — never open a window;
  - `gui` / `on` — ask for dialogs (useful under `pythonw.exe`).
- `OVPP_UI_PATH` points at an alternative package directory; the file is loaded
  by path, so a project can pin a patched copy.
- Buttons, icons, the default button and the return codes `1`–`7` follow
  VBScript's `vbOKOnly` … `vbNo` constants.

## 中文

这是让 OpenVBS++ 的 `MsgBox` 与 `InputBox` 变成**真对话框**的 over 包，由
Python 的 `tkinter`（Tcl/Tk 8.6）绘制。脚本里**不需要 import**：该弹窗时运行时
自己会找到它。

- `over.json` 声明 `"capabilities": ["ui"]` 与 `"kind": "feature"`，所以它不是
  项目的 Python 依赖。
- 查找顺序：项目里的 `packages/ui-support/` → 内置副本
  `OpenVBSPP/vendor/ui-support/`。
- 只有同时满足以下条件才弹窗：脚本由 `openvbs` 命令行启动（不是 IDE，不是在
  Python 里嵌入 `vbs_run()`）、stdin/stdout 是终端、没有设置 `CI`、且 Tk 能创建
  窗口。其余场合退回控制台行为（`MsgBox` 打印并返回 1，`InputBox` 从标准输入读
  一行）。
- `OPENVBS_UI` 覆盖这个判断：
  - `auto`（默认）—— 按上面的规则；
  - `console` / `off` —— 永不弹窗；
  - `gui` / `on` —— 要求弹窗（在 `pythonw.exe` 下有用）。
- `OVPP_UI_PATH` 指向另一个包目录；该文件按路径加载，因此项目可以钉住一个打了
  补丁的副本。
- 按钮组合、图标、默认按钮与返回码 `1`–`7` 都遵循 VBScript 的
  `vbOKOnly` … `vbNo` 常量。

## Licensing / 许可

`tkinter` 与 Tcl/Tk 不随 wheel 或 sdist 分发；只有单文件 exe 会内嵌它们。
Tcl/Tk 的许可是 BSD 风格、允许任意目的使用，但不是 OSI 认证的许可，原文见仓库
根目录的 `THIRD_PARTY_LICENSES.md`。只接受 OSI 许可清单的组织可以设置
`OPENVBS_UI=console`，走不含 Tk 的路径。

The Tcl/Tk license is BSD-style and permissive but not OSI-approved; the
verbatim text is in `THIRD_PARTY_LICENSES.md`. Set `OPENVBS_UI=console` for a
Tk-free path.
