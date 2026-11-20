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

Module compilation and real module isolation.

中文：模块编译与真正的模块隔离。

Model:

中文：模型：
  * a **module** = the entry file plus every file merged in through `<script src>`
    中文：一个模块 = 入口文件 + 它用 `<script src>` 合并进来的所有文件
    (a merged file shares the entry's namespace, so procedures with the same name still
    collide, which is exactly what merging means);
    中文：（合并进来的文件与入口共享同一个命名空间，所以同名过程仍然会冲突——这是「合并」的本意）；
  * `<namespace src alias>` pulls in **another module** with its own globals; only
    中文：`<namespace src alias>` 引入的是另一个模块，有自己独立的 globals，
    only `Public` procedures surface through the alias object;
    中文：只有 `Public` 过程通过别名对象暴露出来；
  * modules run in **topological order** (memoised), so a library may use namespaces
    itself;
    中文：模块按依赖图拓扑执行（记忆化），所以「库文件自己也用 namespace」没问题，
    a cycle is the only error.
    中文：唯一的错误是循环依赖。

Each file compiles to its own code object (with its own filename and line numbers),
and modules execute in declaration order.
中文：每个文件独立编译成 code object（各自的文件名与行号），模块内按声明顺序执行。
'''

import ast
import importlib.util
import marshal
import struct
import time
from pathlib import Path

from OpenVBSPP.ast_nodes import FunctionDef, SubDef
from OpenVBSPP.exceptions import OVPPSyntaxError
from OpenVBSPP.lexer import lex
from OpenVBSPP.parser import Parser
from OpenVBSPP.plugins.modules import extract as extract_modules
from OpenVBSPP.plugins.python_embed import extract as extract_python_blocks
from OpenVBSPP.transformer import Transformer
from OpenVBSPP.typecheck import check as type_check


COMPILED_FILENAME = "<OpenVBS++>"
DEFAULT_SAVE_DIR = "__pycache__"

PYTHON_FEATURE = "python-embed"
PYTHON_IMPORT = '<script src="python-embed" />'


def _capabilities(path):
    """Capabilities declared by the over package owning the entry file (`capabilities`
    in over.json).

    中文：入口文件所属 over 包声明的能力（over.json 的 `capabilities` 字段）。"""
    from OpenVBSPP import over as _over

    directory = Path(path).parent
    if not (directory / _over.MANIFEST_NAME).is_file():
        return frozenset()
    try:
        manifest = _over.Manifest.load(directory)
    except _over.OverError:
        return frozenset()
    return frozenset(cap.lower() for cap in manifest.capabilities)


def _python_block_line(cleaned):
    for number, line in enumerate(cleaned.splitlines(), start=1):
        if line.strip().startswith("__ovpp_python_block_"):
            return number
    return None


def _require_python_feature(features, cleaned):
    """Explain how to import Python interop when `#python` blocks are used without it.

    中文：用了 `#python` 块却没导入 python-embed 时给出明确的迁移提示。
    """
    if PYTHON_FEATURE in features or "__ovpp_python_block_" not in cleaned:
        return
    raise OVPPSyntaxError(
        "Python blocks need an explicit import: add "
        f"{PYTHON_IMPORT} before the first #python block "
        "(since 0.9.20 'Option COP' no longer enables Python interop)",
        line_number=_python_block_line(cleaned))


class ModuleUnit:
    """One module: its code objects, the namespace modules it needs, and its exports.

    中文：一个模块 —— 若干 code object + 它依赖的命名空间模块 + 导出表。
    """

    def __init__(self, name):
        self.name = name
        self.codes = []          # [(filename, code_obj)]
        self.namespaces = []     # [(alias, ModuleUnit)]
        self.exports = []
        self.features = set()
        self.globals = {}
        self._done = False

    # ------------------------------------------------------------------ execute
    # 中文：---------------------------------------------------------------- 执行
    def ensure(self):
        """Execute this module in dependency order, exactly once.

    中文：按依赖顺序执行本模块（只执行一次）。
    """
        if self._done:
            return self.globals
        self._done = True

        from OpenVBSPP.ovpprt.builtins import delegate
        from OpenVBSPP.ovpprt.collections import Dictionary

        for alias, child in self.namespaces:
            child.ensure()
            namespace = Dictionary()
            for vbs_name, py_name, flags in child.exports:
                function = child.globals.get(py_name)
                if function is not None:
                    namespace.add(vbs_name, delegate(function, vbs_name, flags))
            self.globals[alias.lower()] = namespace

        for _filename, code_obj in self.codes:
            exec(code_obj, self.globals)
        return self.globals

    def __repr__(self):
        return f"ModuleUnit({self.name}, {len(self.codes)} files)"


class ProgramBundle:
    """The execution unit of a whole program: the entry module plus its dependencies.

    中文：整个程序的执行单元 —— 入口模块 + 它的依赖。
    """

    def __init__(self, root):
        self.root = root

    def exec(self, globals_dict):
        """Run the program: initialise modules in dependency order, then merge the entry
    module globals into the outer dict.

    中文：执行整个程序 —— 先按依赖初始化各模块，再把入口模块的全局合并进外层字典。"""
        root_globals = self.root.ensure()
        globals_dict.update(root_globals)

    def __repr__(self):
        return f"ProgramBundle({self.root.name})"


# --------------------------------------------------------------------- compile
def _normalize_positions(py_ast):
    """Since Python 3.13 `compile()` validates position ranges, so end must be >= start.

    中文：Python 3.13 起 `compile()` 会校验位置区间，必须保证 end >= start。
    """
    for node in ast.walk(py_ast):
        lineno = getattr(node, "lineno", None)
        if lineno is None:
            continue
        end_lineno = getattr(node, "end_lineno", None)
        if end_lineno is None or end_lineno < lineno:
            node.end_lineno = lineno
        col = getattr(node, "col_offset", 0) or 0
        end_col = getattr(node, "end_col_offset", None)
        if end_col is None or end_col < col:
            node.end_col_offset = col
    return py_ast


def _to_code(py_ast, filename):
    ast.fix_missing_locations(py_ast)
    _normalize_positions(py_ast)
    return compile(py_ast, str(filename) if filename else COMPILED_FILENAME, "exec")


def _parse(source, filename, verbose=False):
    tokens = list(lex(source))
    if verbose:
        print("TOKENS:", tokens)
    parser = Parser(tokens, filename=filename, source=source)
    vbs_ast = parser.parse()
    if verbose:
        print("VBS AST:", vbs_ast)
    type_check(vbs_ast, source=source)
    return vbs_ast


def _resolve(source_path, base_dir):
    # First resolve as a package name: over installs into ./packages/<name>/.
    # 中文：先按「包名」解析（over 安装到 ./packages/<name>/）。
    from OpenVBSPP import over
    packaged = over.resolve_source(str(source_path), base_dir)
    if packaged is not None and packaged.is_file():
        return packaged.resolve()

    candidate = Path(source_path)
    if not candidate.is_absolute():
        candidate = Path(base_dir) / candidate
    candidate = candidate.resolve()
    if not candidate.is_file():
        raise OVPPSyntaxError(
            f"Included file not found: {source_path} (looked in {base_dir})")
    return candidate


def _exports(vbs_asts):
    """Module exports: every `Public` procedure from all merged files.

    中文：模块导出 —— 所有合并文件里的 `Public` 过程。
    """
    entries = []
    for vbs_ast in vbs_asts:
        for stmt in vbs_ast.body:
            if not isinstance(stmt, (SubDef, FunctionDef)):
                continue
            if getattr(stmt, "visibility", "public") != "public":
                continue
            key = stmt.name.lower()
            prefix = "_ovpp_sub_" if stmt.kind == "sub" else "_ovpp_fn_"
            entries.append((stmt.name, prefix + key,
                            [bool(getattr(p, "byref", False))
                             for p in stmt.params]))
    return entries


def _merge_source(unit, source, filename, stack, verbose=False):
    """Collect `source` (and the files it merges with `<script>`) into the unit.

    中文：把 `source`（以及它 `<script>` 合并的文件）收进 unit。
    """
    cleaned, python_blocks = extract_python_blocks(source)
    cleaned, directives = extract_modules(cleaned)

    parts = []            # [(filename, ast, blocks)]
    base_dir = Path(filename).parent if filename else Path.cwd()

    for directive in directives:
        target = _resolve(directive.src, base_dir)
        if PYTHON_FEATURE in _capabilities(target):
            unit.features.add(PYTHON_FEATURE)
            continue
        key = str(target)
        if key in stack:
            chain = " -> ".join(list(stack) + [key])
            raise OVPPSyntaxError(
                f"Circular module dependency: {chain}",
                line_number=directive.line)

        if directive.kind == "script":
            child_source = target.read_text(encoding="utf-8-sig")
            parts.extend(_merge_source(unit, child_source, target,
                                       stack + [key], verbose))
        else:
            alias = (directive.alias or target.stem).strip()
            if not alias.isidentifier():
                raise OVPPSyntaxError(
                    f"Invalid namespace alias: {alias!r}",
                    line_number=directive.line)
            child = _build_module(target, stack + [key], verbose)
            unit.namespaces.append((alias, child))

    _require_python_feature(unit.features, cleaned)
    vbs_ast = _parse(cleaned, filename, verbose=verbose)
    parts.append((filename, vbs_ast, python_blocks))
    return parts


_CACHE = {}


def _build_module(path, stack, verbose=False):
    """Compile one module from a path into a ModuleUnit, with a cache.

    中文：编译一个模块（路径 → ModuleUnit），带缓存。
    """
    key = str(path)
    if key in _CACHE:
        return _CACHE[key]

    source = Path(path).read_text(encoding="utf-8-sig")
    unit = ModuleUnit(Path(path).stem)
    _CACHE[key] = unit            # placeholder, reused on repeated imports
    # 中文：先占位，重复引用同一个模块时复用。

    parts = _merge_source(unit, source, path, stack + [key], verbose)

    transformer = Transformer(features=unit.features)
    for _filename, vbs_ast, _blocks in parts:
        transformer.collect(vbs_ast)

    for filename, vbs_ast, blocks in parts:
        transformer.python_blocks = blocks
        py_ast = transformer._transform_module(vbs_ast)
        unit.codes.append((filename, _to_code(py_ast, filename)))

    unit.exports = _exports([item[1] for item in parts])
    return unit


def vbs_compile(code, mode="exec", filename=None, save_dir=None, verbose=False):
    """Compile OpenVBS++ source.

    Returns a single code object when there are no multi-file declarations (the
    original behaviour), otherwise a ProgramBundle that the runtime executes.

    中文：编译 OpenVBS++ 源码。没有多文件声明时返回单个 code object（保持原有行为）；
    otherwise it returns a ProgramBundle, which runtime executes in order.
    中文：否则返回 ProgramBundle（由 runtime 依次执行）。
    """
    if mode not in ("exec", "save"):
        raise ValueError(f"Invalid mode: {mode!r} (expected 'exec' or 'save')")
    if mode == "save" and not filename:
        raise ValueError("mode='save' requires filename")

    _cleaned, directives = extract_modules(code)

    if not directives:
        python_cleaned, python_blocks = extract_python_blocks(code)
        _require_python_feature(frozenset(), python_cleaned)
        vbs_ast = _parse(python_cleaned, filename, verbose=verbose)
        transformer = Transformer(python_blocks=python_blocks)
        py_ast = transformer.transform(vbs_ast)
        ast.fix_missing_locations(py_ast)
        if verbose:
            print(ast.unparse(py_ast))
        code_obj = _to_code(py_ast, filename)
        if mode == "exec":
            return code_obj
        _save(py_ast, code_obj, filename, save_dir, python_cleaned)
        return code_obj

    # ---------------- Multi-file / 中文：多文件 ----------------
    _CACHE.clear()
    entry_key = str(Path(filename).resolve()) if filename else COMPILED_FILENAME
    unit = ModuleUnit(Path(filename).stem if filename else "main")
    _CACHE[entry_key] = unit

    parts = _merge_source(unit, code, filename, [entry_key], verbose)

    transformer = Transformer(features=unit.features)
    for _name, vbs_ast, _blocks in parts:
        transformer.collect(vbs_ast)
    for name, vbs_ast, blocks in parts:
        transformer.python_blocks = blocks
        py_ast = transformer._transform_module(vbs_ast)
        unit.codes.append((name, _to_code(py_ast, name)))
    unit.exports = _exports([item[1] for item in parts])

    if verbose:
        print("MODULE:", unit)
    return ProgramBundle(unit)


def _save(py_ast, code_obj, filename, save_dir, source):
    stem = Path(filename).stem or "output"
    out_dir = Path(save_dir) if save_dir else Path(DEFAULT_SAVE_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)

    (out_dir / f"{stem}.ast.txt").write_text(ast.unparse(py_ast),
                                             encoding="utf-8")

    with open(out_dir / f"{stem}.pyc", "wb") as handle:
        handle.write(importlib.util.MAGIC_NUMBER)
        handle.write(struct.pack("<I", 0))
        handle.write(struct.pack("<I", int(time.time()) & 0xFFFFFFFF))
        handle.write(struct.pack("<I", len(source.encode("utf-8")) & 0xFFFFFFFF))
        marshal.dump(code_obj, handle)
