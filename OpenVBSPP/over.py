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

`over` —— OpenVBS++ 的包管理器（M2：本地源 + 依赖安装 + 锁文件 + pip 依赖）。

设计取舍（没有中央服务器）：
  * 包源只有两种：**本地目录**（含 over.json）和 **.ovppkg 单文件包**（zip）；
  * 安装目标是**项目本地**的 `./packages/<name>/`，不污染全局环境；
  * 版本约束只做 semver 子集：`=1.2.3` / `^1.2` / `>=1.2` / `<2`；
  * `over.lock` 记录精确版本 + 内容哈希，保证可重现；
  * 所有操作离线可用。

清单 `over.json`：
    {
      "name": "my-app",
      "version": "0.1.0",
      "entry": "src/main.vbs",
      "dependencies": {"text-utils": "file:../text-utils"},
      "python": {"requests": ">=2.31"},
      "platforms": ["win32"]
    }

`dependencies` 是 over 包（.vbs）；`python` 是 PyPI 依赖（pip 装）。
安装时会遍历**整张 over 依赖图**收集所有 `python` 需求，一次性交给 pip，
由 pip 自己解析传递依赖，真实安装结果写进 `over.lock`。

源码里用 `<namespace src="text-utils" alias="tu">` 引用包名时，编译器会先找
`./packages/text-utils/`，再退回相对路径。
'''

import hashlib
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path


MANIFEST_NAME = "over.json"
LOCK_NAME = "over.lock"
PACKAGES_DIR = "packages"
VENDOR_DIR = "vendor"
COM_PACKAGE = "com-support"

#: 锁格式：2 = 多了 "extra"/"extras"/"python" 字段（仍能读 1）
LOCK_VERSION = 2


class OverError(Exception):
    """包管理器错误。"""


# --------------------------------------------------------------------- 版本
_VERSION_RE = re.compile(r"(\d+)(?:\.(\d+))?(?:\.(\d+))?")


def parse_version(text):
    match = _VERSION_RE.search(str(text))
    if not match:
        raise OverError(f"Invalid version: {text!r} (expected 1.2.3)")
    major, minor, patch = match.groups()
    return (int(major), int(minor or 0), int(patch or 0))


def satisfies(version, spec):
    """semver 子集：=1.2.3 / ^1.2 / >=1.2 / <2，空表示任意。"""
    spec = (spec or "").strip()
    if not spec or spec == "*":
        return True
    current = parse_version(version)

    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if part.startswith(">="):
            if current < parse_version(part[2:]):
                return False
        elif part.startswith("<="):
            if current > parse_version(part[2:]):
                return False
        elif part.startswith(">"):
            if current <= parse_version(part[1:]):
                return False
        elif part.startswith("<"):
            if current >= parse_version(part[1:]):
                return False
        elif part.startswith("^"):
            base = parse_version(part[1:])
            if current < base or current[0] != base[0]:
                return False
        elif part.startswith("="):
            if current != parse_version(part[1:]):
                return False
        else:
            if current != parse_version(part):
                return False
    return True


# --------------------------------------------------------------------- 清单
class Manifest:
    """一个包的 over.json。"""

    def __init__(self, data, root=None):
        self.data = data
        self.root = Path(root) if root else None

    @property
    def name(self):
        return self.data.get("name", "")

    @property
    def version(self):
        return str(self.data.get("version", "0.0.0"))

    @property
    def entry(self):
        return self.data.get("entry", "")

    @property
    def dependencies(self):
        return dict(self.data.get("dependencies") or {})

    @property
    def description(self):
        return str(self.data.get("description", ""))

    @property
    def kind(self):
        """包类型：空 = 普通 .vbs 包，"python" = 纯 Python 支撑包。"""
        return str(self.data.get("kind", "") or "")

    @property
    def capabilities(self):
        """这个包向语言暴露的能力（例如 ["python-embed"]）。

        `<script src="包名" />` 指向一个带能力的包时，编译器不合并它的源码，
        而是打开对应能力 —— 目前用于 Python 互操作（`#python` 块与 `COP`）。
        """
        raw = self.data.get("capabilities")
        if not raw:
            return []
        if isinstance(raw, str):
            return [raw]
        return [str(item).strip().lower() for item in raw if str(item).strip()]

    @property
    def platforms(self):
        """适用平台列表（空 = 全平台）。写法：["win32"] / ["linux", "darwin"]。"""
        raw = self.data.get("platforms")
        if not raw:
            return []
        if isinstance(raw, str):
            return [raw]
        return [str(item) for item in raw if str(item).strip()]

    @property
    def python_dependencies(self):
        """Python(pip) 依赖：`{"pywin32": ">=306"}`；也接受 `["pywin32>=306"]`。"""
        raw = self.data.get("python") or {}
        if isinstance(raw, str):
            raw = [raw]
        if isinstance(raw, list):
            result = {}
            for item in raw:
                name, spec = split_pip_requirement(item)
                result[name] = spec
            return result
        return {str(name): str(spec or "").strip()
                for name, spec in dict(raw).items() if str(name).strip()}

    def supports_platform(self, platform=None):
        """当前平台是否适用（无 platforms 声明时为真）。"""
        wanted = self.platforms
        if not wanted:
            return True
        current = str(platform or sys.platform).lower()
        aliases = {
            "win32": ("win32", "windows", "win"),
            "windows": ("win32", "windows", "win"),
            "darwin": ("darwin", "macos", "mac"),
            "macos": ("darwin", "macos", "mac"),
            "linux": ("linux",),
        }
        for item in wanted:
            item = item.strip().lower()
            for key in aliases.get(item, (item,)):
                if current == key or current.startswith(key):
                    return True
        return False

    @classmethod
    def load(cls, directory):
        path = Path(directory) / MANIFEST_NAME
        if not path.is_file():
            raise OverError(f"{MANIFEST_NAME} not found in {directory}")
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise OverError(f"Invalid {MANIFEST_NAME}: {exc}") from exc
        return cls(data, directory)

    def save(self):
        if self.root is None:
            raise OverError("manifest has no root directory")
        (self.root / MANIFEST_NAME).write_text(
            json.dumps(self.data, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8")


# --------------------------------------------------------------------- 安装
def _hash_tree(directory):
    digest = hashlib.sha256()
    for path in sorted(Path(directory).rglob("*")):
        if path.is_file():
            digest.update(str(path.relative_to(directory)).encode("utf-8"))
            digest.update(path.read_bytes())
    return digest.hexdigest()[:16]


def _source_path(source, base):
    """把依赖声明解析成本地目录 / .ovppkg 路径。"""
    spec = str(source).strip()
    if spec.startswith("file:"):
        spec = spec[5:]
    candidate = Path(spec)
    if not candidate.is_absolute():
        candidate = Path(base) / candidate
    return candidate


def read_lock(root):
    """读 over.lock（没有就返回 None）。"""
    path = Path(root) / LOCK_NAME
    if not path.is_file():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise OverError(f"Invalid {LOCK_NAME}: {exc}") from exc


def _install_frozen(root, lock, sources, verbose=False, python_deps=True,
                    pip_runner=None, python_executable=None):
    """按锁文件安装：版本必须与锁一致，保证可重现。"""
    packages_dir = Path(root) / PACKAGES_DIR
    packages_dir.mkdir(exist_ok=True)
    installed = []

    for entry in lock.get("packages", []):
        name = entry.get("name", "")
        locked_version = entry.get("version")
        source = sources.get(name) or _locate(entry.get("source", ""), root)
        target = packages_dir / name

        if source.is_dir():
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(source, target,
                            ignore=shutil.ignore_patterns(".git", PACKAGES_DIR))
        elif source.is_file() and zipfile.is_zipfile(source):
            with zipfile.ZipFile(source) as archive:
                archive.extractall(target)
        else:
            raise OverError(f"Cannot restore '{name}' from {source}")

        child = Manifest.load(target)
        if locked_version and child.version != locked_version:
            raise OverError(
                f"Locked {name}@{locked_version} but the source now provides "
                f"{child.version}; run 'over update' to accept it")

        installed.append({"name": name, "version": child.version,
                          "source": entry.get("source", ""),
                          "hash": _hash_tree(target)})
        if verbose:
            print(f"installed {name} {child.version} (locked)")

    for entry in lock.get("extras", []):
        name = entry.get("name", "")
        if not name:
            continue
        try:
            provision_bundled(root, name, verbose=verbose)
        except OverError as exc:
            raise OverError(
                f"Cannot restore bundled package '{name}': {exc}") from exc

    # python 依赖：钉到锁里记录的精确版本
    if python_deps:
        restored = _restore_python_from_lock(root, lock,
                                             python_executable=python_executable,
                                             runner=pip_runner,
                                             verbose=verbose)
        if restored and verbose:
            print(f"python dependencies restored ({len(restored)})")

    (Path(root) / LOCK_NAME).write_text(
        json.dumps(lock, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return installed


def install(root, source_map=None, verbose=False, frozen=False, dry_run=False,
            python_deps=True, pip_runner=None, python_executable=None,
            with_com=False):
    """安装 over.json 里的依赖（本地目录 / .ovppkg / git+URL）以及 python 依赖。

    frozen=True 且有 over.lock 时，按锁文件精确恢复（可重现安装）。
    python_deps=False 可跳过 pip（纯 over 包、离线环境）。
    """
    root = Path(root)
    manifest = Manifest.load(root)
    sources = dict(source_map or {})

    if dry_run:
        return plan_install(root, sources, verbose=verbose)

    packages_dir = root / PACKAGES_DIR
    packages_dir.mkdir(exist_ok=True)

    lock = read_lock(root)
    if frozen and lock:
        return _install_frozen(root, lock, sources, verbose,
                               python_deps=python_deps,
                               pip_runner=pip_runner,
                               python_executable=python_executable)

    # 先把整张依赖图走一遍：解析来源、检测循环、收集版本约束
    requirements = collect_requirements(root)

    installed = []
    visited = set()
    python_needs = _PythonCollector()
    python_needs.absorb(manifest)

    def _install(name, spec, depth=0):
        if name in visited:
            return
        visited.add(name)
        if depth > 32:
            raise OverError(f"Dependency chain too deep at '{name}'")

        source_spec, _version_spec = _split_spec(spec)
        target = packages_dir / name
        located = sources.get(name) or _locate(source_spec, root)

        if located.is_dir():
            child = Manifest.load(located)
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(located, target,
                            ignore=shutil.ignore_patterns(".git", PACKAGES_DIR))
        elif located.is_file() and zipfile.is_zipfile(located):
            with zipfile.ZipFile(located) as archive:
                archive.extractall(target)
            child = Manifest.load(target)
        else:
            raise OverError(
                f"Cannot resolve package '{name}' from {located} "
                f"(supported: local folder, .ovppkg, git+URL)")

        # 版本冲突：该包的所有需求都必须被满足
        for requirement in requirements.get(name, []):
            if requirement.spec and not satisfies(child.version, requirement.spec):
                raise OverError(
                    f"Conflict for '{name}': source provides {child.version}, "
                    f"but {requirement.spec!r} is required by "
                    f"{requirement.describe()}")

        entry = {"name": name, "version": child.version,
                 "source": str(located), "hash": _hash_tree(target)}
        commit = _git_commit(located)   # 只对 git 检出有意义（复制到 packages 时会去掉 .git）
        if commit:
            entry["commit"] = commit
        installed.append(entry)
        if verbose:
            print(f"installed {name} {child.version}")

        # 这个包的 python 依赖也并进来（pip 稍后一次性解析）
        python_needs.absorb(child)

        for child_name, child_spec in child.dependencies.items():
            _install(child_name, child_spec, depth + 1)

    for name, spec in manifest.dependencies.items():
        _install(name, spec)

    lock_document = {"version": LOCK_VERSION, "packages": installed}

    # 内置包（com-support）：显式要求时才装进项目；Windows 上即使不装，
    # 随包分发的副本也能让 CreateObject 直接可用。
    if with_com:
        extra = provision_bundled(root, COM_PACKAGE, verbose=verbose)
        lock_document["extras"] = [extra]
        python_needs.absorb(Manifest.load(root / PACKAGES_DIR / COM_PACKAGE))

    python_entries = _resolve_python(python_needs, python_deps=python_deps,
                                     pip_runner=pip_runner,
                                     python_executable=python_executable,
                                     verbose=verbose)
    if python_entries:
        lock_document["python"] = python_entries

    (root / LOCK_NAME).write_text(
        json.dumps(lock_document, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8")
    return installed


def _resolve_python(collector, python_deps=True, pip_runner=None,
                    python_executable=None, verbose=False):
    """把收集到的 python 依赖交给 pip，返回要写进锁文件的条目。"""
    if not python_deps:
        return []
    wanted = collector.entries()
    if not wanted:
        return []

    requests = [entry["requirement"] for entry in wanted.values()]
    if pip_runner is None and not pip_available(python_executable):
        raise OverError(
            "pip is not available for "
            f"{python_executable or sys.executable}; python dependencies "
            f"({', '.join(requests)}) cannot be installed. Use --no-python "
            "to skip them.")

    resolved = install_python(requests, python_executable=python_executable,
                              runner=pip_runner, verbose=verbose)
    entries = []
    for item in resolved:
        entry = dict(item)
        entry["required_by"] = list(
            (wanted.get(item["name"]) or {}).get("required_by", []))
        entries.append(entry)
        if verbose:
            print(f"python {item['name']} {item['version']}"
                  + ("" if entry["required_by"] else " (transitive)"))
    return entries


def update(root, source_map=None):
    """忽略缓存重新解析并安装（git 源会重新拉取）。"""
    import shutil as _shutil
    cache = Path.home() / ".openvbspp" / "over" / "cache"
    if cache.is_dir():
        _shutil.rmtree(cache, ignore_errors=True)
    packages = Path(root) / PACKAGES_DIR
    if packages.is_dir():
        _shutil.rmtree(packages, ignore_errors=True)
    return install(root, source_map=source_map)




def init(root, name=None):
    """生成一个最小的 over.json。"""
    root = Path(root)
    path = root / MANIFEST_NAME
    if path.is_file():
        raise OverError(f"{MANIFEST_NAME} already exists in {root}")
    root.mkdir(parents=True, exist_ok=True)
    data = {"name": name or root.name or "app", "version": "0.1.0",
            "entry": "main.vbs", "dependencies": {}}
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                    encoding="utf-8")
    return path


def list_installed(root):
    packages_dir = Path(root) / PACKAGES_DIR
    if not packages_dir.is_dir():
        return []
    result = []
    for path in sorted(packages_dir.iterdir()):
        if not path.is_dir():
            continue
        try:
            manifest = Manifest.load(path)
        except OverError:
            continue
        result.append((manifest.name or path.name, manifest.version))
    return result


def bundled_package_dir(name=COM_PACKAGE):
    """随包分发的内置 over 包目录（OpenVBSPP/vendor/<name>）。"""
    return Path(__file__).resolve().parent / VENDOR_DIR / name


def bundled_packages():
    """随包分发的全部内置 over 包：[(name, version), ...]。"""
    vendor = Path(__file__).resolve().parent / VENDOR_DIR
    found = []
    if not vendor.is_dir():
        return found
    for directory in sorted(vendor.iterdir()):
        if not (directory / MANIFEST_NAME).is_file():
            continue
        try:
            manifest = Manifest.load(directory)
        except OverError:
            continue
        found.append((manifest.name or directory.name, manifest.version))
    return found


def provision_bundled(root, name=COM_PACKAGE, verbose=False):
    """把随包分发的内置包装进 <项目>/packages/<name>/，返回锁文件条目。

    Windows 上 com-support 只要存在于 packages/ 就优先用它；不装也能用
    （运行期会退回 vendor 副本），显式装进来是为了让包管理器可见、可卸载。
    """
    source = bundled_package_dir(name)
    if not source.is_dir():
        raise OverError(
            f"Bundled package '{name}' is missing from this installation "
            f"({source})")
    target = Path(root) / PACKAGES_DIR / name
    if target.exists():
        shutil.rmtree(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, target,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    manifest = Manifest.load(target)
    if verbose:
        print(f"installed {name} {manifest.version} (bundled)")
    return {"name": name, "version": manifest.version,
            "source": str(source), "kind": "bundled",
            "hash": _hash_tree(target)}


def list_installed_details(root):
    """列出项目里一切可管理的东西：over 包、内置包、python 依赖。"""
    root = Path(root)
    lock = read_lock(root) or {}
    details = []

    for entry in lock.get("packages", []):
        name = entry.get("name", "")
        target = root / PACKAGES_DIR / name
        details.append({"name": name, "version": entry.get("version", ""),
                        "kind": "over", "source": entry.get("source", ""),
                        "present": target.is_dir(), "required_by": []})

    for entry in lock.get("extras", []):
        name = entry.get("name", "")
        target = root / PACKAGES_DIR / name
        details.append({"name": name, "version": entry.get("version", ""),
                        "kind": "bundled", "source": entry.get("source", ""),
                        "present": target.is_dir(), "required_by": []})

    for entry in lock.get("python", []):
        details.append({"name": entry.get("name", ""),
                        "version": entry.get("version", ""),
                        "kind": "python", "source": "pypi",
                        "present": True,
                        "required_by": list(entry.get("required_by") or [])})
    return details


def info(root, name):
    """一个包的详情；找不到返回 None。"""
    wanted = _normalize_name(name)
    for entry in list_installed_details(root):
        if _normalize_name(entry["name"]) != wanted:
            continue
        manifest_path = Path(root) / PACKAGES_DIR / entry["name"] / MANIFEST_NAME
        if manifest_path.is_file():
            manifest = Manifest.load(manifest_path.parent)
            entry["description"] = manifest.description
            if manifest.kind:
                entry["package_kind"] = manifest.kind
            entry["platforms"] = manifest.platforms
            entry["dependencies"] = dict(manifest.dependencies)
            entry["python_dependencies"] = dict(manifest.python_dependencies)
        else:
            entry.setdefault("description", "")
        if entry["kind"] in ("bundled", "over") and entry["name"] == COM_PACKAGE:
            entry["runtime"] = com_support_status(root)
        return entry
    return None


def remove(root, name, python=False, python_executable=None, pip_runner=None,
           verbose=False):
    """移除一个包：over 包删目录 + 锁条目；python=True 走 pip 卸载。"""
    root = Path(root)
    lock = read_lock(root) or {}
    wanted = _normalize_name(name)

    if python:
        entries = [entry for entry in lock.get("python", [])
                   if _normalize_name(entry.get("name", "")) == wanted]
        if not entries:
            raise OverError(
                f"python dependency '{name}' is not recorded in {LOCK_NAME}")
        uninstall_python([entries[0].get("name", name)],
                         python_executable=python_executable,
                         runner=pip_runner, verbose=verbose)
        lock["python"] = [entry for entry in lock.get("python", [])
                          if _normalize_name(entry.get("name", "")) != wanted]
        if not lock["python"]:
            lock.pop("python", None)
        target = entries[0]
    else:
        target = None
        for key in ("packages", "extras"):
            kept = []
            for entry in lock.get(key, []):
                if _normalize_name(entry.get("name", "")) == wanted:
                    target = entry
                else:
                    kept.append(entry)
            if kept:
                lock[key] = kept
            else:
                lock.pop(key, None)
        if target is None and not (root / PACKAGES_DIR / name).exists():
            raise OverError(f"'{name}' is not installed in this project")
        directory = root / PACKAGES_DIR / str((target or {}).get("name", name))
        if directory.exists():
            shutil.rmtree(directory)

    # 顺手从 over.json 的依赖声明里摘掉
    manifest_path = root / MANIFEST_NAME
    if manifest_path.is_file():
        manifest = Manifest.load(root)
        dependencies = manifest.data.get("dependencies") or {}
        python_map = manifest.data.get("python") or {}
        changed = False
        for key in list(dependencies):
            if _normalize_name(key) == wanted:
                dependencies.pop(key, None)
                changed = True
        if isinstance(python_map, dict):
            for key in list(python_map):
                if _normalize_name(key) == wanted:
                    python_map.pop(key, None)
                    changed = True
        if changed:
            manifest.save()

    (root / LOCK_NAME).write_text(
        json.dumps(lock, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    if verbose:
        print(f"removed {name}")
    return target


def com_support_status(root=None):
    """COM 支持的当前状态（给 `over list` / `over info com-support` 用）。"""
    from OpenVBSPP.ovpprt import com as bridge

    directory, origin = bridge.resolve()
    platform_ok = sys.platform == "win32"
    return {"platform_ok": platform_ok,
            "directory": str(directory) if directory else "",
            "origin": origin,
            "available": bool(directory) and platform_ok}


def add(root, name, spec=None):
    """往 over.json 里加一条依赖并安装；内置包可以省略 source。

    省略 source 且 `<name>` 是随包分发的内置包时，把它复制进
    `packages/<name>/` 并记进锁文件（之后 `over remove` 能卸掉）。
    """
    root = Path(root)
    manifest = Manifest.load(root)
    spec = (spec or "").strip()
    if not spec and bundled_package_dir(name).is_dir():
        entry = provision_bundled(root, name)
        lock = read_lock(root) or {"version": LOCK_VERSION}
        extras = [item for item in lock.get("extras") or []
                  if item.get("name") != name]
        extras.append(entry)
        lock["version"] = LOCK_VERSION
        lock["extras"] = extras
        (root / LOCK_NAME).write_text(
            json.dumps(lock, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8")
        return [entry]
    manifest.data.setdefault("dependencies", {})[name] = spec
    manifest.save()
    return install(root)


# --------------------------------------------------------------------- python 依赖
def _normalize_name(name):
    """PEP 503 规范化：`PyWin32` / `py_win32` -> `pywin32`。"""
    return re.sub(r"[-_.]+", "-", str(name)).strip().lower()


_PIP_NAME_RE = re.compile(r"^\s*([A-Za-z0-9][A-Za-z0-9._-]*)\s*(.*)$")


def split_pip_requirement(text):
    """`pywin32>=306` -> ("pywin32", ">=306")。"""
    match = _PIP_NAME_RE.match(str(text))
    if not match:
        raise OverError(f"Invalid python requirement: {text!r}")
    return match.group(1), match.group(2).strip()


def requirement_strings(requirements):
    """{名字: 约束} 或 ["名字约束"] -> 送进 pip 的需求字符串列表。"""
    if isinstance(requirements, dict):
        result = []
        for name, spec in requirements.items():
            if not str(name).strip():
                continue
            result.append(f"{str(name).strip()}{str(spec or '').strip()}")
        return result
    result = []
    for item in requirements or []:
        text = str(item).strip()
        if text:
            result.append(text)
    return result


class _PythonCollector:
    """汇总整张 over 依赖图上的 python 依赖（同名约束合并、去掉重复来源）。"""

    def __init__(self):
        self._found = {}

    def absorb(self, manifest):
        """把一个 over 包清单里的 python 依赖并进来（平台不符则跳过）。"""
        if manifest is None or not manifest.supports_platform():
            return
        for name, spec in manifest.python_dependencies.items():
            entry = self._found.setdefault(
                _normalize_name(name),
                {"name": name, "specs": [], "required_by": []})
            if spec and spec not in entry["specs"]:
                entry["specs"].append(spec)
            owner = manifest.name or "(root)"
            if owner not in entry["required_by"]:
                entry["required_by"].append(owner)

    def entries(self):
        """{规范化名: {name, specs, required_by, requirement}}"""
        result = {}
        for key, entry in self._found.items():
            spec = ",".join(entry["specs"])
            result[key] = {"name": entry["name"],
                           "specs": list(entry["specs"]),
                           "required_by": list(entry["required_by"]),
                           "requirement": f"{entry['name']}{spec}"}
        return result

    def __bool__(self):
        return bool(self._found)


def pip_available(python_executable=None):
    """这个解释器能不能跑 pip。"""
    import subprocess

    executable = python_executable or sys.executable
    try:
        result = subprocess.run([executable, "-m", "pip", "--version"],
                                capture_output=True, text=True)
    except (OSError, ValueError):
        return False
    return result.returncode == 0


def default_pip_runner(requirements, python_executable=None, verbose=False):
    """真正调用 pip 安装，返回 [{name, version, requested}]。

    用 `pip install --report <tmp.json>` 拿到**解析后的完整闭包**（含传递
    依赖），因此锁文件描述的是实际装了什么，而不只是声明了什么。
    """
    import subprocess
    import tempfile

    requests = requirement_strings(requirements)
    if not requests:
        return []
    executable = python_executable or sys.executable

    with tempfile.TemporaryDirectory(prefix="openvbs-over-") as tmp:
        report = Path(tmp) / "report.json"
        command = [executable, "-m", "pip", "install",
                   "--disable-pip-version-check", "--no-input",
                   "--report", str(report), *requests]
        try:
            result = subprocess.run(command, capture_output=True, text=True)
        except (OSError, ValueError) as exc:
            raise OverError(
                f"Cannot run pip with {executable}: {exc}") from exc
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "").strip()
            tail = "\n".join(detail.splitlines()[-12:])
            raise OverError("pip install failed for "
                            + ", ".join(requests) + ":\n" + tail)
        if verbose:
            print(f"pip install {' '.join(requests)}")
        payload = {}
        if report.is_file():
            try:
                payload = json.loads(report.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                payload = {}

    installed = []
    for entry in payload.get("install", []):
        metadata = entry.get("metadata") or {}
        name = metadata.get("name")
        if not name:
            continue
        installed.append({"name": name,
                          "version": str(metadata.get("version", "")),
                          "requested": bool(entry.get("requested"))})
    return installed


def default_pip_uninstaller(names, python_executable=None, verbose=False):
    """调用 pip 卸载一组 python 包，返回真正卸载掉的名单。"""
    import subprocess

    package_names = [str(name) for name in names or [] if str(name).strip()]
    if not package_names:
        return []
    executable = python_executable or sys.executable
    command = [executable, "-m", "pip", "uninstall", "-y",
               "--disable-pip-version-check", *package_names]
    try:
        result = subprocess.run(command, capture_output=True, text=True)
    except (OSError, ValueError) as exc:
        raise OverError(f"Cannot run pip with {executable}: {exc}") from exc
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "").strip()
        raise OverError("pip uninstall failed for "
                        + ", ".join(package_names) + ":\n" + detail)
    if verbose:
        print("pip uninstall " + " ".join(package_names))
    return package_names


def install_python(requirements, python_executable=None, runner=None,
                   verbose=False):
    """安装 python 依赖；`runner` 可注入，便于离线单测。"""
    requests = requirement_strings(requirements)
    if not requests:
        return []
    runner = runner or default_pip_runner
    raw = runner(requests, python_executable=python_executable,
                 verbose=verbose)
    installed = []
    for item in raw or []:
        name = item.get("name") or item.get("package")
        if not name:
            continue
        installed.append({"name": _normalize_name(name),
                          "version": str(item.get("version", "")),
                          "requested": bool(item.get("requested", False))})
    return installed


def uninstall_python(names, python_executable=None, runner=None,
                     verbose=False):
    """卸载 python 依赖；`runner` 可注入，便于离线单测。"""
    package_names = [str(name) for name in names or [] if str(name).strip()]
    if not package_names:
        return []
    runner = runner or default_pip_uninstaller
    return list(runner(package_names, python_executable=python_executable,
                       verbose=verbose) or [])


def lock_python(root):
    """锁文件里记录的 python 依赖（没有则空列表）。"""
    lock = read_lock(root) or {}
    return list(lock.get("python") or [])


def _restore_python_from_lock(root, lock, python_executable=None,
                              runner=None, verbose=False):
    """按锁文件把 python 依赖钉到记录的版本（`--frozen`）。"""
    entries = lock.get("python") or []
    if not entries:
        return []
    pinned = [f"{item.get('name')}=={item.get('version')}"
              for item in entries if item.get("name") and item.get("version")]
    if not pinned:
        return []
    if runner is None:
        return install_python(pinned, python_executable=python_executable,
                              verbose=verbose)
    return install_python(pinned, python_executable=python_executable,
                          runner=runner, verbose=verbose)


# --------------------------------------------------------------------- 解析包路径
def resolve_source(name, base_dir):
    """给编译器用：包名 -> 入口文件路径（找不到返回 None）。

    查找顺序：项目里的 `packages/<name>/`（可钉版本、换实现）优先，
    然后是随包分发的 `OpenVBSPP/vendor/<name>/`（内置包，开箱即用）。
    """
    for candidate in (Path(base_dir) / PACKAGES_DIR / name,
                      bundled_package_dir(name)):
        if not candidate.is_dir():
            continue
        try:
            manifest = Manifest.load(candidate)
        except OverError:
            continue
        entry = manifest.entry or "main.vbs"
        if entry:
            return candidate / entry
    return None


# --------------------------------------------------------------------- Git 源
GIT_PREFIXES = ("git+", "git:")


def _rmtree(path):
    """删目录，顺带处理 git 留下的只读对象文件。"""
    import os
    import stat

    def _force(func, target, _exc):
        try:
            os.chmod(target, stat.S_IWRITE)
        except OSError:
            pass
        try:
            func(target)
        except OSError:
            pass

    shutil.rmtree(path, onerror=_force)


def _git_cache_dir():
    root = Path.home() / ".openvbspp" / "over" / "cache"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _fetch_git(spec):
    """`git+https://host/repo.git#ref` -> 本地缓存的检出目录。"""
    import subprocess

    raw = spec[4:] if spec.startswith("git+") else spec[3:]
    url, _, ref = raw.partition("#")
    name = re.sub(r"[^A-Za-z0-9_.-]", "_", url.rstrip("/").split("/")[-1] or "repo")
    if name.endswith(".git"):
        name = name[:-4]
    target = _git_cache_dir() / (name + ("_" + ref.replace("/", "_") if ref else ""))

    # 上一次 clone 被打断（或缓存被手工清理过）时只剩半个 .git，重来一遍
    if target.exists() and not (target / MANIFEST_NAME).is_file():
        _rmtree(target)

    if not (target / ".git").exists():
        try:
            subprocess.run(["git", "clone", "--depth", "1", url, str(target)],
                           check=True, capture_output=True)
        except FileNotFoundError as exc:
            raise OverError(
                "git is required for 'git+' sources but was not found on PATH"
            ) from exc
        except subprocess.CalledProcessError as exc:
            raise OverError(
                f"git clone failed for {url}: {exc.stderr.decode(errors='replace').strip()}"
            ) from exc

    if ref:
        try:
            subprocess.run(["git", "-C", str(target), "fetch", "--depth", "1",
                            "origin", ref], check=True, capture_output=True)
            subprocess.run(["git", "-C", str(target), "checkout", "FETCH_HEAD"],
                           check=True, capture_output=True)
        except (FileNotFoundError, subprocess.CalledProcessError) as exc:
            raise OverError(f"git checkout {ref} failed for {url}") from exc
    return target


def _locate(spec, base):
    """依赖声明 -> 本地目录（支持目录、.ovppkg、git+URL）。"""
    text = str(spec).strip()
    if text.startswith(GIT_PREFIXES):
        return _fetch_git(text)
    return _source_path(text, base)


# --------------------------------------------------------------------- 解析
class Requirement:
    """一条依赖需求（带来源链，用于冲突报告）。"""

    __slots__ = ("name", "spec", "chain", "source")

    def __init__(self, name, spec, chain, source=None):
        self.name = name
        self.spec = spec
        self.chain = chain
        self.source = source

    def describe(self):
        return " -> ".join(self.chain)


def _split_spec(spec):
    if isinstance(spec, dict):
        return str(spec.get("source", "")), str(spec.get("version", ""))
    return str(spec), ""


def collect_requirements(root):
    """遍历依赖图，收集每个包的版本约束与来源链（含循环检测）。"""
    root = Path(root)
    manifest = Manifest.load(root)
    found = {}

    def walk(name, spec, chain, depth):
        if depth > 32:
            raise OverError(f"Dependency chain too deep at '{name}'")
        if name in chain:
            raise OverError(
                "Circular dependency: " + " -> ".join(chain + [name]))
        source, version = _split_spec(spec)
        found.setdefault(name, []).append(Requirement(name, version, chain, source))

        located = _locate(source, root)
        if located.is_dir():
            child = Manifest.load(located)
        elif located.is_file():
            with zipfile.ZipFile(located) as archive:
                if MANIFEST_NAME not in archive.namelist():
                    raise OverError(f"{located} has no {MANIFEST_NAME}")
                import json as _json
                child = Manifest(_json.loads(
                    archive.read(MANIFEST_NAME).decode("utf-8")), None)
        else:
            raise OverError(f"Cannot resolve package '{name}' from {source}")

        for child_name, child_spec in child.dependencies.items():
            walk(child_name, child_spec, chain + [name], depth + 1)

    for name, spec in manifest.dependencies.items():
        walk(name, spec, [manifest.name], 0)
    return found


def check_conflicts(requirements, versions):
    """同一包被多个约束要求时，报告所有来源链。"""
    problems = []
    for name, reqs in requirements.items():
        version = versions.get(name)
        if version is None:
            continue
        for requirement in reqs:
            if requirement.spec and not satisfies(version, requirement.spec):
                problems.append(
                    f"Conflict for '{name}': source provides {version}, "
                    f"but {requirement.spec!r} is required by "
                    f"{requirement.describe()}")
    return problems


# --------------------------------------------------------------------- pack
def pack(directory, output=None):
    """把一个包目录打成 `.ovppkg`（zip），返回产物路径。"""
    directory = Path(directory)
    manifest = Manifest.load(directory)
    target = Path(output) if output else (
        directory.parent / f"{manifest.name}-{manifest.version}.ovppkg")
    target.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(directory.rglob("*")):
            if not path.is_file():
                continue
            relative = path.relative_to(directory)
            if relative.parts and relative.parts[0] in (".git", PACKAGES_DIR):
                continue
            archive.write(path, relative.as_posix())
    return target


# --------------------------------------------------------------------- dry-run
def plan_install(root, source_map=None, verbose=False):
    """只解析依赖图，不写盘（`over install --dry-run`）。

    返回 [{"name", "version", "source", "spec", "chain", "action"}]，
    action 为 "install" 或 "unchanged"。
    """
    root = Path(root)
    sources = dict(source_map or {})
    requirements = collect_requirements(root)
    installed = {}
    for item in list_installed(root):
        if isinstance(item, dict):
            installed[item.get("name")] = item.get("version")
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            installed[item[0]] = item[1]
    plan = []

    for name, reqs in requirements.items():
        spec = reqs[0].spec if reqs else ""
        chain = reqs[0].describe() if reqs else ""
        source = (reqs[0].source if reqs and reqs[0].source else spec)
        located = sources.get(name) or _locate(source, root)

        version = None
        if located.is_dir():
            version = Manifest.load(located).version
        elif located.is_file() and zipfile.is_zipfile(located):
            with zipfile.ZipFile(located) as archive:
                import json as _json
                if MANIFEST_NAME in archive.namelist():
                    version = _json.loads(
                        archive.read(MANIFEST_NAME).decode("utf-8")).get("version")

        current = installed.get(name)
        action = "install" if not current or current != version else "unchanged"
        plan.append({"name": name, "version": version, "current": current,
                     "source": str(located), "spec": spec, "chain": chain,
                     "action": action})
        if verbose:
            print(f"[dry-run] {action:9} {name} {version}")

    return plan


# --------------------------------------------------------------------- git 源
def _git_commit(path):
    """取一个 git 检出目录的 HEAD commit（非 git 目录返回 None）。"""
    import subprocess

    path = Path(path)
    if not (path / ".git").exists():
        return None
    try:
        result = subprocess.run(["git", "-C", str(path), "rev-parse", "HEAD"],
                                check=True, capture_output=True, text=True)
    except (OSError, subprocess.CalledProcessError):
        return None
    return result.stdout.strip() or None


def git_tags(spec):
    """`git ls-remote --tags` 枚举远程版本（去掉 v 前缀与 ^{}）。"""
    import subprocess

    raw = spec[4:] if spec.startswith("git+") else spec[3:]
    url, _, _ref = raw.partition("#")
    try:
        result = subprocess.run(["git", "ls-remote", "--tags", url],
                                check=True, capture_output=True, text=True)
    except FileNotFoundError as exc:
        raise OverError("git is required for 'git+' sources but was not found "
                        "on PATH") from exc
    except subprocess.CalledProcessError as exc:
        raise OverError(f"git ls-remote failed for {url}: "
                        f"{(exc.stderr or '').strip()}") from exc

    found = []
    for line in result.stdout.splitlines():
        if "refs/tags/" not in line:
            continue
        tag = line.split("refs/tags/", 1)[1].strip()
        if tag.endswith("^{}"):
            continue
        version = parse_version(tag.lstrip("vV"))
        if version is not None:
            found.append((version, tag))
    found.sort()
    return [tag for _version, tag in found]


def select_git_ref(spec, version_spec=None):
    """在远程标签里挑一个满足约束的最新版本；没有标签就返回 None。"""
    tags = git_tags(spec)
    for tag in reversed(tags):
        if not version_spec or satisfies(parse_version(tag.lstrip("vV")), version_spec):
            return tag
    return tags[-1] if tags and not version_spec else None


# --------------------------------------------------------------------- search
def registry_index():
    """本地 registry：`~/.openvbspp/over/registry.json`（没有则空）。"""
    path = Path.home() / ".openvbspp" / "over" / "registry.json"
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}
    return data.get("packages", data if isinstance(data, dict) else {})


def search(term, root=None, registry=None):
    """在本地 registry + 已安装包 + 依赖声明里搜索（大小写不敏感）。"""
    term = (term or "").lower()
    results = []

    for name, entry in (registry if registry is not None else registry_index()).items():
        text = f"{name} {json.dumps(entry, ensure_ascii=False)}".lower()
        if term and term not in text:
            continue
        versions = (entry or {}).get("versions") or {}
        latest = None
        if versions:
            keyed = [(parse_version(item) or 0, item) for item in versions]
            latest = max(keyed)[1]
        elif isinstance(entry, dict):
            latest = entry.get("version")      # 兼容旧的扁平结构
        description = (entry or {}).get("description", "")
        if not description and versions and latest:
            description = (versions.get(latest) or {}).get("description", "")
        results.append({"name": name, "version": latest, "source": "registry",
                        "description": description})

    if root is not None:
        for item in list_installed(root):
            if isinstance(item, dict):
                name, version = item.get("name"), item.get("version")
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                name, version = item[0], item[1]
            else:
                continue
            if not term or term in str(name).lower():
                results.append({"name": name, "version": version,
                                "source": "installed", "description": ""})
    return results


# --------------------------------------------------------------------- registry
REGISTRY_NAME = "registry.json"


def registry_path(explicit=None):
    """registry.json 的位置：显式路径 > ~/.openvbspp/over/registry.json。"""
    if explicit:
        return Path(explicit)
    return Path.home() / ".openvbspp" / "over" / REGISTRY_NAME


def _pkg_fingerprint(path):
    """算一个包的 sha256 与大小，作为 registry 的完整性凭据。"""
    import hashlib

    digest = hashlib.sha256()
    size = 0
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def _manifest_from_pkg(path):
    """从 .ovppkg（zip）里读 over.json。"""
    with zipfile.ZipFile(path) as archive:
        if MANIFEST_NAME not in archive.namelist():
            raise OverError(f"{path} has no {MANIFEST_NAME}")
        return json.loads(archive.read(MANIFEST_NAME).decode("utf-8"))


def register(package, registry=None, source=None, description=None):
    """把 `.ovppkg`（或包目录）登记进本地 registry.json（publish-less 分发）。

    返回 {"name", "version", "hash", "size", "registry"}；
    同名同版本的条目会被覆盖，其它版本保留（多版本并存）。
    """
    package = Path(package)
    target_registry = registry_path(registry)

    if package.is_dir():
        loaded = Manifest.load(package)
        manifest = dict(loaded.data)
        payload = Path(pack(package))
    elif package.is_file() and zipfile.is_zipfile(package):
        manifest = _manifest_from_pkg(package)
        payload = package
    else:
        raise OverError(f"Cannot register '{package}': not a .ovppkg or package folder")

    if not manifest.get("name"):
        raise OverError(f"{payload} has no package name in {MANIFEST_NAME}")

    digest, size = _pkg_fingerprint(payload)
    entry = {
        "version": str(manifest.get("version", "0.0.0")),
        "source": str(source or payload.resolve()),
        "sha256": digest,
        "size": size,
        "entry": manifest.get("entry", ""),
        "description": description or manifest.get("description", ""),
        "dependencies": manifest.get("dependencies", {}),
    }

    data = {}
    if target_registry.is_file():
        try:
            data = json.loads(target_registry.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise OverError(f"Invalid {target_registry.name}: {exc}") from exc
    packages = data.setdefault("packages", {})
    versions = packages.setdefault(manifest["name"], {}).setdefault("versions", {})
    versions[str(entry["version"])] = {key: value for key, value in entry.items()
                                       if key != "version"}
    packages[manifest["name"]]["name"] = manifest["name"]

    target_registry.parent.mkdir(parents=True, exist_ok=True)
    target_registry.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                               encoding="utf-8")
    return {"name": manifest["name"], "version": entry["version"],
            "hash": digest, "size": size, "registry": str(target_registry)}


def index_directory(directory, registry=None, verbose=False):
    """把一个目录下所有 `.ovppkg` 批量登记进 registry（`over index`）。"""
    directory = Path(directory)
    if not directory.is_dir():
        raise OverError(f"Not a directory: {directory}")
    results = []
    for path in sorted(directory.rglob("*.ovppkg")):
        try:
            results.append(register(path, registry=registry))
        except OverError as exc:
            if verbose:
                print(f"skipped {path.name}: {exc}")
            continue
    return results


def registry_versions(name, registry=None):
    """registry 里某个包的所有版本（升序）。"""
    path = registry_path(registry)
    if not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    entry = data.get("packages", {}).get(name, {})
    versions = list((entry.get("versions") or {}).keys())
    pairs = [(parse_version(item), item) for item in versions]
    pairs = [pair for pair in pairs if pair[0] is not None]
    return [text for _key, text in sorted(pairs)]
