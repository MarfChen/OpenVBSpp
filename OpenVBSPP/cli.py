'''
Copyright 2027 Marf Chen(MarfChen)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

OpenVBS++ command line entry point (installed as the `openvbs` command).

中文：OpenVBS++ 命令行入口（安装后即为 `openvbs` 命令）。
'''


import argparse
import sys
from pathlib import Path

from OpenVBSPP import __version__
from OpenVBSPP.exceptions import OVPPError


DEMO = """
rem No arguments: run the built-in demo
rem 中文：没有参数时运行内置示例
a = 123
echo "Hello World"
echo a
"""


def build_arg_parser():
    """Build the argument parser for `openvbs [script] [-v]`.

    中文：构造命令行参数解析器（`openvbs [script] [-v]`）。
    """
    parser = argparse.ArgumentParser(
        prog="openvbs",
        description="Run an OpenVBS++ (VBScript-flavoured) script.",
    )
    parser.add_argument("script", nargs="?",
                        help="Script to run; omit to run the built-in demo / 要运行的 .vbs 文件，省略则运行内置示例")
    parser.add_argument("--save", action="store_true",
                        help="Also write .pyc and .ast.txt / 同时写出 .pyc 与 .ast.txt")
    parser.add_argument("--save-dir", default=None,
                        help="Output directory, default ./__pycache__ / 落盘目录")
    parser.add_argument("-c", "--code", default=None, metavar="CODE",
                        help="Run inline source (`-` reads standard input) / 直接运行一段源码，`-` 表示从标准输入读")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Print tokens, AST and generated Python / 打印 token 流、AST 与生成的 Python 代码")
    parser.add_argument("--version", action="version",
                        version=f"OpenVBS++ {__version__}",
                        help="Show the version and exit / 显示版本号并退出")
    return parser


def _relax_console():
    """Old Windows consoles default to GBK/CP1252; encoding errors degrade to
    replacement characters so a script never dies printing Chinese text.

    中文：Windows 旧控制台默认 GBK/CP1252 —— 把编码错误降级为替换字符，而不是让脚本
    A single CJK character used to crash printing.

    中文：因为打印一个中文就崩溃。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(errors="replace")
        except (ValueError, OSError):
            pass


def main(argv=None):
    """Command line entry: the `ide` subcommand opens the web IDE, anything else runs a
    script.

    中文：命令行入口 —— `ide` 子命令走网页 IDE，否则运行脚本。"""
    _relax_console()
    argv = list(sys.argv[1:] if argv is None else argv)

    # Global --json flag, allowed before or after the subcommand.
    # 中文：全局开关 --json（可放在子命令前后）。
    as_json = "--json" in argv
    argv = [item for item in argv if item != "--json"]

    # Subcommands: openvbs [--json] check|run <script>|-c CODE
    # 中文：子命令 openvbs [--json] check|run <script>[-c CODE]。
    if argv and argv[0] in ("check", "run"):
        rest = argv[1:]
        inline = None
        for flag in ("-c", "--code"):
            if flag in rest:
                index = rest.index(flag)
                if index + 1 >= len(rest):
                    print(f"usage: openvbs {argv[0]} {flag} <code>", file=sys.stderr)
                    return 2
                inline = rest[index + 1]
                rest = rest[:index] + rest[index + 2:]
                break
        target = next((item for item in rest
                       if item == "-" or not item.startswith("-")), None)
        if target is None and inline is None:
            print("usage: openvbs [--json] check|run <script>", file=sys.stderr)
            return 2
        if argv[0] == "check":
            return _command_check(target, as_json, inline)
        return _command_run(target, as_json, inline)
    # Subcommand: openvbs over ... — the package manager (over packages plus
    # Python dependencies).
    # 中文：子命令 openvbs over ... —— 包管理器（over 包 + python 依赖）。
    if argv and argv[0] == "over":
        from OpenVBSPP import over as over_pkg
        rest = argv[1:]
        here = Path.cwd()
        usage = (
            "usage: openvbs over <command> [options]\n"
            "\n"
            "  init [name]                    write an over.json / 写一个 over.json\n"
            "  install [--frozen] [--dry-run] [--no-python] [--with-com]\n"
            "                                 install over dependencies and their python deps / 安装 over 依赖并自动装 python 依赖\n"
            "  list                           list installed packages (over / bundled / python) / 列出已安装的包\n"
            "  info <name>                    show one package in detail / 查看包详情\n"
            "  remove <name> [--python]       uninstall a package; --python hands it to pip / 卸载包\n"
            "  update                         re-resolve and upgrade per over.json / 重新解析并升级\n"
            "  add <name> [source]            add and install; source may be omitted for a bundled package / 添加并安装\n"
            "  search <term>                  search the local registry / 在本地 registry 里搜索\n"
            "  com                            show COM support status / 查看 COM 支持状态\n"
            "  pack <dir> [out.ovppkg]        pack into an .ovppkg / 打包成 .ovppkg\n"
            "  publish|register <pkg>         register into the local registry / 登记到本地 registry\n"
            "  index [dir]                    register every .ovppkg in a directory / 批量登记目录下的 .ovppkg\n")
        try:
            if not rest or rest[0] in ("-h", "--help", "help"):
                print(usage)
                return 0
            command = rest[0]

            if command == "init":
                path = over_pkg.init(here, rest[1] if len(rest) > 1 else None)
                print(f"created {path}")
            elif command == "install":
                frozen = "--frozen" in rest
                dry_run = "--dry-run" in rest or "-n" in rest
                python_deps = "--no-python" not in rest
                with_com = "--with-com" in rest
                items = over_pkg.install(here, frozen=frozen, dry_run=dry_run,
                                         python_deps=python_deps,
                                         with_com=with_com)
                for item in items:
                    if dry_run:
                        print(f"[dry-run] {item['action']:9} {item['name']} "
                              f"{item['version']} ({item['source']})")
                    else:
                        print(f"installed {item['name']} {item['version']}")
                if dry_run and not items:
                    print("[dry-run] nothing to do")
                if not dry_run:
                    for entry in over_pkg.lock_python(here):
                        who = ", ".join(entry.get("required_by") or [])
                        suffix = f" (required by {who})" if who else ""
                        print(f"python {entry['name']} {entry['version']}{suffix}")
            elif command == "search":
                term = rest[1] if len(rest) > 1 else ""
                found = over_pkg.search(term, here)
                for item in found:
                    print(f"{item['name']} {item['version'] or '-'} [{item['source']}]")
                if not found:
                    print("no packages found")
            elif command == "list":
                details = over_pkg.list_installed_details(here)
                if not details:
                    print("no packages installed")
                for item in details:
                    suffix = ""
                    if item["kind"] == "bundled":
                        suffix = " (bundled)"
                    elif item["kind"] == "python":
                        who = ", ".join(item.get("required_by") or [])
                        suffix = f" (python, required by {who})" if who else " (python)"
                    missing = "" if item.get("present", True) else " [missing]"
                    print(f"{item['name']} {item['version']}{suffix}{missing}")
                com = over_pkg.com_support_status(here)
                print(f"COM support: {com['origin']}"
                      + ("" if com["available"] else " (unavailable)"))
                bundled = over_pkg.bundled_packages()
                if bundled:
                    names = ", ".join(f"{name} {version}"
                                      for name, version in bundled)
                    print(f"bundled with OpenVBS++: {names}")
            elif command == "info":
                if len(rest) < 2:
                    print("usage: openvbs over info <name>", file=sys.stderr)
                    return 2
                entry = over_pkg.info(here, rest[1])
                if entry is None:
                    print(f"over: '{rest[1]}' is not installed", file=sys.stderr)
                    return 1
                for key in ("name", "version", "kind", "source", "description",
                            "platforms", "required_by", "dependencies",
                            "python_dependencies", "present"):
                    if key not in entry:
                        continue
                    value = entry[key]
                    if isinstance(value, (list, dict)) and not value:
                        continue
                    print(f"{key}: {value}")
            elif command == "remove":
                if len(rest) < 2:
                    print("usage: openvbs over remove <name> [--python]",
                          file=sys.stderr)
                    return 2
                over_pkg.remove(here, rest[1], python="--python" in rest)
                print(f"removed {rest[1]}")
            elif command == "com":
                from OpenVBSPP.ovpprt import com as bridge
                status = over_pkg.com_support_status(here)
                if "--json" in rest:
                    import json as _json
                    payload = dict(status)
                    payload["detail"] = bridge.describe()
                    print(_json.dumps(payload, ensure_ascii=False))
                else:
                    print(f"directory: {status['directory'] or '-'}")
                    print(f"origin: {status['origin']}")
                    print(f"platform: {'windows' if status['platform_ok'] else sys.platform}")
                    print(f"detail: {bridge.describe()}")
            elif command == "update":
                for item in over_pkg.update(here):
                    print(f"updated {item['name']} {item['version']}")
            elif command == "pack" and len(rest) >= 2:
                output = over_pkg.pack(Path(rest[1]), rest[2] if len(rest) > 2 else None)
                print(f"packed {output}")
            elif command in ("publish", "register") and len(rest) >= 2:
                info = over_pkg.register(rest[1])
                print(f"registered {info['name']} {info['version']} "
                      f"sha256={info['hash'][:12]} -> {info['registry']}")
            elif command == "index":
                folder = rest[1] if len(rest) > 1 else "."
                items = over_pkg.index_directory(folder, verbose=True)
                for info in items:
                    print(f"registered {info['name']} {info['version']}")
                if not items:
                    print("no .ovppkg found")
            elif command == "add" and len(rest) >= 2:
                over_pkg.add(here, rest[1], rest[2] if len(rest) > 2 else None)
                print(f"added {rest[1]}")
            else:
                print(usage, file=sys.stderr)
                return 2
        except over_pkg.OverError as exc:
            print(f"over: {exc}", file=sys.stderr)
            return 1
        return 0

    # Subcommand: openvbs ide [--port 7860] starts the built-in web IDE.
    # 中文：子命令 openvbs ide [--port 7860] —— 启动内置网页 IDE。
    if argv and argv[0] in ("ide", "studio"):
        from OpenVBSPP.ide.__main__ import main as ide_main
        return ide_main(argv[1:])

    args = build_arg_parser().parse_args(argv)

    if args.code is not None and args.script:
        print("openvbs: use either `-c CODE` or a script file, not both / 用 `-c CODE` 和脚本文件只能二选一", file=sys.stderr)
        return 2

    if args.code is not None:
        if args.code == "-":
            code, filename = sys.stdin.read(), "<stdin>"
        else:
            code, filename = args.code, "<-c>"
    elif args.script == "-":
        code, filename = sys.stdin.read(), "<stdin>"
    elif args.script:
        path = Path(args.script)
        if not path.is_file():
            print(f"File not found: {path}", file=sys.stderr)
            return 1
        # utf-8-sig also strips the BOM Notepad adds on Windows.
        # 中文：utf-8-sig 顺手吃掉 Windows 记事本加的 BOM。
        code = path.read_text(encoding="utf-8-sig")
        filename = str(path)
    else:
        code = DEMO
        filename = "<demo>"

    try:
        # Imported lazily: runtime pulls in compiler, transformer and typecheck
        # (about 27 ms), which `openvbs --version` and `openvbs over ...` never need.
        # 中文：延迟 import —— runtime 会把 compiler/transformer/typecheck 一路拉进来
        # which costs about 27 ms and neither `openvbs --version` nor `openvbs over ...`
        # needs it.
        # 中文：（约 27 ms），而 `openvbs --version`、`openvbs over ...` 根本用不到。
        from OpenVBSPP.runtime import vbs_run
        from OpenVBSPP.ovpprt import ui

        # Scripts started from the command line are the ones allowed to open
        # dialogs; the library entry point `vbs_run` deliberately is not.
        # 中文：只有从命令行启动的脚本才允许弹窗；库入口 `vbs_run` 刻意不算。
        ui.mark_cli_entry()

        vbs_run(code, filename=filename, verbose=args.verbose,
                mode="save" if args.save else "exec",
                save_dir=args.save_dir)
    except OVPPError as err:
        print(err, file=sys.stderr)
        return 1

    return 0

def _load_source(target, inline=None):
    """Read source: `inline` supplies the code directly, `target` `-` reads stdin, and
    `utf-8-sig` strips a BOM.

    中文：读源码 —— `inline` 直接给源码，`target` 为 `-` 时读标准输入，`utf-8-sig`
    eat a BOM.

    中文：吃掉 BOM。"""
    if inline is not None:
        if inline == "-":
            return sys.stdin.read(), "<stdin>"
        return inline, "<-c>"
    if target == "-":
        return sys.stdin.read(), "<stdin>"
    path = Path(target)
    if not path.is_file():
        return None, str(target)
    return path.read_text(encoding="utf-8-sig"), str(path)


def _command_check(target, as_json, inline=None):
    """Compile only: return the diagnostics.

    中文：只编译不运行 —— 拿诊断。
    """
    import json as _json
    from OpenVBSPP.ide import service

    source, name = _load_source(target, inline)
    if source is None:
        payload = {"ok": False, "diagnostics": [],
                   "error": {"kind": "FileNotFound", "message": f"File not found: {name}"}}
        print(_json.dumps(payload, ensure_ascii=False) if as_json
              else f"File not found: {name}", file=sys.stderr if not as_json else sys.stdout)
        return 1

    result = service.check_source(source, filename=name)
    if as_json:
        payload = {"ok": result["ok"], "diagnostics": result["diagnostics"]}
        print(_json.dumps(payload, ensure_ascii=False))
    elif result["ok"]:
        print(f"{name}: OK")
    else:
        for item in result["diagnostics"]:
            line = item.get("line")
            where = f" (line {line})" if line else ""
            print(f"{name}: {item.get('kind')}: {item.get('message')}{where}",
                  file=sys.stderr)
    return 0 if result["ok"] else 1


def _command_run(target, as_json, inline=None):
    """Compile and run.

    中文：编译并运行。
    """
    import json as _json
    from OpenVBSPP.ide import service

    source, name = _load_source(target, inline)
    if source is None:
        payload = {"ok": False, "output": "",
                   "error": {"kind": "FileNotFound",
                             "message": f"File not found: {name}"},
                   "duration_ms": 0.0}
        print(_json.dumps(payload, ensure_ascii=False) if as_json
              else f"File not found: {name}", file=sys.stderr if not as_json else sys.stdout)
        return 1

    result = service.run_source(source, filename=name)
    if as_json:
        print(_json.dumps(result, ensure_ascii=False))
    else:
        sys.stdout.write(result.get("output") or "")
        if result.get("error"):
            error = result["error"]
            line = error.get("line")
            where = f" (line {line})" if line else ""
            print(f"{error.get('kind')}: {error.get('message')}{where}",
                  file=sys.stderr)
    return 0 if result["ok"] else 1
