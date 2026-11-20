'''
Copyright 2026 Marf Chen(MarfChen)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

GUI dispatcher for `MsgBox` and `InputBox`.

Scripts never import anything: `MsgBox` and `InputBox` ask this module whether
the `ui-support` over package should show a real dialog, and keep the answer for
the rest of the process. Lookup order mirrors the COM bridge:

  1. the directory named by `OVPP_UI_PATH` (tests and advanced setups);
  2. `./packages/ui-support/` inside the project (a project install wins);
  3. the copy bundled with the distribution (`OpenVBSPP/vendor/ui-support/`).

The activation rules are the interesting part, because a dialog that nobody can
click would hang an automated run:

  * `OPENVBS_UI=console` (or `off`) never shows dialogs;
  * `OPENVBS_UI=gui` shows them even without a TTY (`pythonw.exe`);
  * the default `auto` shows them only when the process was started from the
    command line (`mark_cli_entry()`), stdin and stdout are a terminal, `CI` is
    not set, `tkinter` imports and a hidden Tk root can be created.

Any failure while showing a dialog disables GUI mode for the rest of the
process, so a broken Tk installation degrades to the console instead of
repeating the error on every call.

中文：`MsgBox` / `InputBox` 的 GUI 分发层。

脚本里不需要 import：这两个内置函数会询问本模块要不要用 `ui-support` 包弹出
真正的对话框，并把决定保持到进程结束。查找顺序与 COM 桥接一致：
`OVPP_UI_PATH` 环境变量 → 项目里的 `./packages/ui-support/`（项目内安装优先）→
随包分发的 `OpenVBSPP/vendor/ui-support/`。

关键在启用规则，因为一个没人能点的对话框会卡死自动化运行：
`OPENVBS_UI=console`（或 `off`）永不弹窗；`OPENVBS_UI=gui` 即使没有 TTY
也弹（`pythonw.exe` 场景）；默认 `auto` 只在「从命令行启动
（`mark_cli_entry()`）、stdin 与 stdout 都是终端、未设置 `CI`、`tkinter` 能导入
且能创建隐藏 Tk 根」时才弹。

弹窗过程中任何失败都会把 GUI 模式关掉直到进程结束，所以坏掉的 Tk 安装会退回
控制台，而不是每次调用都报一遍错。
'''

import importlib.util
import os
import sys
from pathlib import Path

#: Environment variable that overrides the activation decision.
#: 中文：覆盖启用策略的环境变量。
ENV = "OPENVBS_UI"

#: Environment variable that points at an explicit package directory.
#: 中文：指向显式包目录的环境变量。
ENV_PATH = "OVPP_UI_PATH"

#: Name of the bundled over package.
#: 中文：内置 over 包的名字。
PACKAGE_NAME = "ui-support"

#: File inside the package that holds the implementation.
#: 中文：包里存放实现的文件名。
MODULE_FILE = "ui_support.py"

#: Directory that holds project-installed packages.
#: 中文：项目内安装包所在的目录名。
PACKAGES_DIR = "packages"

#: Manifest file that marks a project root.
#: 中文：标记项目根的清单文件名。
MANIFEST_NAME = "over.json"

_CONSOLE_VALUES = {"console", "off", "0", "false", "no", "disable", "disabled"}
_GUI_VALUES = {"gui", "1", "on", "true", "yes"}

#: Loaded `ui_support` modules: {resolved path: (mtime, module)}.
#: 中文：已加载的 `ui_support` 模块，键是解析后的路径。
_MODULES = {}

#: True once the CLI entry point called `mark_cli_entry()`.
#: 中文：CLI 入口调用过 `mark_cli_entry()` 后为 True。
_CLI_ENTRY = False

#: True after a dialog failed, which switches the process to console mode.
#: 中文：弹窗失败过之后为 True，进程随即改走控制台模式。
_FAILED = False


def mark_cli_entry():
    """Remember that a script is running through the command line.

    中文：记住「脚本是从命令行启动的」这件事。
    """
    global _CLI_ENTRY
    _CLI_ENTRY = True


def reset_cache():
    """Forget the CLI marker, the failure flag and every loaded module.

    中文：清空 CLI 标记、失败标记和所有已加载的模块（测试用）。
    """
    global _CLI_ENTRY, _FAILED
    _CLI_ENTRY = False
    _FAILED = False
    _MODULES.clear()
    for name in [item for item in sys.modules
                 if item.startswith("openvbspp_ui_support_")]:
        sys.modules.pop(name, None)


def _bundled_dir():
    """The bundled copy: `OpenVBSPP/vendor/ui-support/`.

    中文：随包分发的副本 `OpenVBSPP/vendor/ui-support/`。
    """
    return Path(__file__).resolve().parent.parent / "vendor" / PACKAGE_NAME


def _project_dir(start=None):
    """Walk up from `start` (default: cwd) to the nearest project root with `over.json`.

    中文：从 `start`（默认当前目录）向上找一个带 `over.json` 的项目根。
    """
    current = Path(start or Path.cwd()).resolve()
    for candidate in (current, *current.parents):
        if (candidate / MANIFEST_NAME).is_file():
            return candidate
    return None


def candidates():
    """Return `[(directory, origin)]` in priority order.

    中文：按优先级返回 [(目录, 来源标签)]。
    """
    found = []
    explicit = os.environ.get(ENV_PATH)
    if explicit:
        found.append((Path(explicit), "env"))

    project = _project_dir()
    if project is not None:
        found.append((project / PACKAGES_DIR / PACKAGE_NAME, "project"))
    found.append((_bundled_dir(), "bundled"))
    return found


def resolve():
    """Return the effective `ui-support` directory: `(directory, origin)` or `(None, reason)`.

    中文：返回实际生效的 `ui-support` 目录 —— (目录, 来源) 或 (None, 原因)。
    """
    tried = []
    for directory, origin in candidates():
        if (Path(directory) / MODULE_FILE).is_file():
            return Path(directory), origin
        tried.append(f"{directory} ({origin})")
    return None, "ui-support package not found; tried: " + ", ".join(tried)


def _load(directory):
    """Load `ui_support.py` from a directory, cached by modification time.

    中文：从目录加载 `ui_support.py`（按 mtime 缓存）。
    """
    script = Path(directory) / MODULE_FILE
    key = str(script.resolve())
    try:
        stamp = script.stat().st_mtime
    except OSError:
        stamp = None

    cached = _MODULES.get(key)
    if cached is not None and cached[0] == stamp:
        return cached[1]

    module_name = "openvbspp_ui_support_" + str(abs(hash(key)))
    spec = importlib.util.spec_from_file_location(module_name, key)
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load ui-support from {key}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise

    _MODULES[key] = (stamp, module)
    return module


def _env_mode():
    """Return `console`, `gui` or `auto` from `OPENVBS_UI`.

    中文：从 `OPENVBS_UI` 读出 `console`、`gui` 或 `auto`。
    """
    value = os.environ.get(ENV)
    if value is None:
        return "auto"
    text = str(value).strip().lower()
    if text in _CONSOLE_VALUES:
        return "console"
    if text in _GUI_VALUES:
        return "gui"
    return "auto"


def _is_a_terminal():
    """True when both stdin and stdout look like a terminal.

    中文：stdin 与 stdout 是否都像终端。
    """
    try:
        return bool(sys.stdin.isatty() and sys.stdout.isatty())
    except (AttributeError, ValueError):
        return False


def enabled():
    """Decide whether dialogs should be used: `(flag, reason)`.

    中文：决定要不要用对话框 —— (是否启用, 原因)。
    """
    if _FAILED:
        return False, "a dialog failed earlier in this process"

    mode = _env_mode()
    if mode == "console":
        return False, f"disabled by {ENV}"

    if mode == "auto":
        if not _CLI_ENTRY:
            return False, "not started from the command line"
        if not _is_a_terminal():
            return False, "stdin/stdout are not a terminal"
        if os.environ.get("CI"):
            return False, "CI is set"

    directory, origin = resolve()
    if directory is None:
        return False, origin

    try:
        module = _load(directory)
    except Exception as exc:                     # broken package, not a crash
        return False, f"ui-support failed to load ({exc})"

    check = getattr(module, "is_available", None)
    if check is not None and not check():
        return False, "tkinter is not available"

    return True, f"{origin}: {directory}"


def provider():
    """Return the loaded `ui_support` module, or None when dialogs are off.

    中文：返回已加载的 `ui_support` 模块；不弹窗时返回 None。
    """
    flag, _reason = enabled()
    if not flag:
        return None
    directory, _origin = resolve()
    try:
        return _load(directory)
    except Exception:
        return None


def _failed():
    """Switch the rest of the process to console mode.

    中文：让本进程后续调用都改走控制台。
    """
    global _FAILED
    _FAILED = True


def msgbox(message, buttons=0, title=""):
    """Show `MsgBox`; returns the VBScript code 1-7, or None for the console path.

    中文：执行 `MsgBox`；返回 VBScript 的 1-7，走控制台路径时返回 None。
    """
    module = provider()
    if module is None:
        return None
    try:
        return int(module.show_message(message, buttons, title or ""))
    except Exception:
        _failed()
        return None


def inputbox(prompt, title="", default=""):
    """Show `InputBox`; returns the text, or None for the console path.

    中文：执行 `InputBox`；返回文本，走控制台路径时返回 None。
    """
    module = provider()
    if module is None:
        return None
    try:
        return str(module.ask_text(prompt, title or "", default or ""))
    except Exception:
        _failed()
        return None


def describe():
    """One-line status for diagnostics and the package manager.

    中文：给诊断与包管理器用的一句话状态。
    """
    flag, reason = enabled()
    state = "enabled" if flag else "console"
    return f"{state} ({reason})"
