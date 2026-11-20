# The `over` package manager

> English | **[中文](OVER.zh-CN.md)**

`over` is an offline-first, project-local package manager. Its sources are local folders,
`.ovppkg` archives (zip) and `git+` URLs; it never requires a central server. Installed
packages land in `./packages/`, a lock file records exact versions and content hashes, and
`<namespace src="pkg">` resolves package names through that directory. Packages may also
declare **Python dependencies** (the `python` field): `over install` hands them to `pip` in
one batch and records what really landed in `over.lock`. On Windows a bundled `com-support`
package makes `CreateObject` talk to real COM without any import.

## 1. Positioning

| Dimension | Choice | Rationale |
|---|---|---|
| Registry | **No central server** | The maintainer has no server; local folders / Git / single-file packages are enough |
| Install location | `./packages/` inside the project | Does not pollute the global environment; projects stay independent |
| Version constraints | semver subset `^1.2` `>=1.0` `<2` `=1.2.3` | Enough in practice; no SAT solver |
| Lock file | `over.lock` (version + content hash) | Reproducible; commit it to Git |
| Dependencies | Standard library only (git sources shell out to `git`) | Keeps "zero third-party dependencies" |
| Python dependencies | Handed to `pip`; no home-grown wheel building | pip already solves transitive deps, wheels and platform tags |
| Platform differences | Declared as `platforms` in the manifest; mismatches are skipped | Simpler than optionalDependencies and sufficient |

## 2. Manifest

```json
{
  "name": "app",
  "version": "0.1.0",
  "entry": "main.vbs",
  "kind": "over",
  "description": "example project",
  "platforms": ["win32", "linux"],
  "python": {
    "pywin32": ">=306",
    "requests": "^2.31"
  },
  "dependencies": {
    "text-utils": {"source": "file:../text-utils", "version": "^1.0"},
    "json-helper": {"source": "git+https://example.com/json-helper.git#v1.2"},
    "legacy": "file:../legacy.ovppkg"
  }
}
```

* `source` accepts three forms: `file:<dir>`, `file:<x>.ovppkg`, `git+<url>[#ref]`;
* The string form means "source only, any version"; the object form may add a `version` constraint;
* `python` lists the **PyPI packages** this package needs: `{"name": "constraint"}`. The pip
  requirement-string form `["requests>=2.31", "pywin32"]` is also accepted;
* `platforms` are the supported platforms (`win32` / `windows`, `darwin` / `macos`, `linux`,
  optionally with a version prefix). Empty means all platforms. A package whose platform
  does not match is **not installed and not an error**, so one `over.json` can be shared
  across platforms;
* `kind` and `description` are only used for `over list` / `over info` output.

## 3. Commands

| Command | Effect |
|---|---|
| `openvbs over` | print the usage (does not fail when `over.json` is missing) |
| `openvbs over install` | resolve the dependency graph → install into `./packages/` → install Python deps → write `over.lock` |
| `openvbs over install --frozen` | restore strictly from `over.lock` (byte-stable), pinning the Python versions in the lock |
| `openvbs over install --dry-run` | resolve only; print each dependency's action without writing |
| `openvbs over install --no-python` | skip Python dependency resolution (offline / restricted environments) |
| `openvbs over install --with-com` | also install the bundled `com-support` into `./packages/` |
| `openvbs over update` | clear the cache, delete `packages/`, resolve and install again |
| `openvbs over list` | list over packages / bundled packages / Python dependencies, plus the COM support origin |
| `openvbs over info <name>` | show one package's version, source, dependencies and (for com-support) runtime status |
| `openvbs over remove <name>` | uninstall an over package (from `packages/`, the lock and `over.json`) |
| `openvbs over remove <name> --python` | uninstall a Python dependency (`pip uninstall`, and drop the `python` entry from the lock) |
| `openvbs over com [--json]` | show the directory, origin, platform and availability of the COM bridge |
| `openvbs over add <name> [source]` | write it into `over.json` and install immediately (source may be omitted for bundled packages) |
| `openvbs over search <term>` | search the local registry and installed packages |
| `openvbs over pack <dir> [out]` | pack a package directory into `.ovppkg` (zip, ignoring `.git` and `packages/`) |
| `openvbs over publish` / `register` / `index` | publish to the local registry, register, generate the index |

## 4. Dependency resolution

1. Walk the dependency graph depth-first starting from the root `over.json`;
2. **Circular dependencies** fail immediately during the walk and print the chain:
   `Circular dependency: app -> a -> b -> a`;
3. Every package collects **all** version constraints and source chains;
4. During install the version actually provided by the source is checked against every
   constraint; conflicts report:

```
Conflict for 'shared': source provides 1.5.0, but '^2.0' is required by app -> b
```

5. Each package is installed once (first come, first served); later requirements only run
   version checks;
6. **Python dependencies** are collected across the whole over graph (transitive packages
   included): constraints for the same name are comma-merged (`six>=1.16,<2`) and handed to
   `pip install` **in one batch**, letting pip resolve transitive dependencies;
7. pip runs with `--report`, so `over.lock` records the versions that were **actually
   installed**, not the constraints; that is what lets `--frozen` pin the same closure on
   another machine.

## 5. How it connects to the compiler

An installed package can be referenced by name — no path needed:

```vb
<namespace src="text-utils" alias="tu">
echo tu.Trim(" x ")
```

When the compiler resolves `<namespace src>`, it first looks in `./packages/<name>/` (using
the `entry` from that package's `over.json`) and falls back to a relative path. So **the
package manager never has to change the compiler**.

## 6. `over.lock`

```json
{
  "version": 2,
  "packages": [
    {"name": "text-utils", "version": "1.2.0",
     "source": "file:../text-utils", "hash": "9f2c1a3b7d4e5f60"}
  ],
  "extras": [
    {"name": "com-support", "version": "1.0.0",
     "source": "bundled", "kind": "bundled", "hash": "1a2b3c4d5e6f7a8b"}
  ],
  "python": [
    {"name": "pywin32", "version": "312",
     "requested": true, "required_by": ["com-support"]}
  ]
}
```

* `hash` is the first 16 hex digits of the SHA-256 over the package directory contents
  (paths + bytes), catching "same version, different contents";
* `extras` records **bundled** packages shipped with the release (currently only
  `com-support`), whose source is written as `bundled`;
* `python` is the closure pip really installed: `version` is the actual version,
  `requested` says whether the manifest declared it (`true`) or pip pulled it in as a
  transitive dependency (`false`), and `required_by` names the packages that demanded it;
* `version` is the lock-format version; v1 locks are still readable (a missing `extras` /
  `python` is treated as empty);
* Put `packages/` into `.gitignore` and commit `over.lock`.

## 7. Milestones

| Stage | Content | Status |
|---|---|---|
| M1 | over.json read/write + local directory install + lock file | ✅ (lock readable since v0.9.11, exact `--frozen` restore) |
| M2 | Semantic versioning + constraint solving (DFS + conflict reporting) | ✅ |
| M3 | `.ovppkg` pack/unpack + Git sources | ✅ |
| M4 | `<namespace src="package">` wiring (compiler side) | ✅ |
| M5 | Python dependencies (`python` field → pip → closure in the lock) | ✅ v0.9.14 |
| M6 | Platform conditions (`platforms`) + bundled packages | ✅ v0.9.14 |

## 8. Python dependencies

Declare them in the manifest:

```json
{"python": {"pywin32": ">=306"}}
```

`over install` then:

1. Walks the whole over graph and collects every package's `python` requirements into one
   set (same-name constraints comma-merged);
2. Checks that `pip` is available (`<python> -m pip --version`); it only fails when there
   really are requirements, and the message points at `--no-python`;
3. Runs `pip install <requirements...> --report <temp json>`;
4. Reads the **actual install result** out of the report (including transitive
   dependencies pip pulled in) and writes it into the `python` array of `over.lock`.

`--frozen` does the reverse: it reads the lock's `python` entries, installs them as
`name==version` pins, and then verifies the result matches the lock.
`over remove <name> --python` goes through `pip uninstall`.

`openvbs over list` prints Python dependencies separately (with `required by ...`), and
`over info <name>` shows the `python_dependencies` each package declares.

## 9. COM support (the `com-support` package)

On Windows, COM objects such as `CreateObject("WScript.Shell")` are provided by the
`com-support` package, which wraps `pywin32` and is fully transparent to scripts:

* **Shipped with the release**: a copy lives in `OpenVBSPP/vendor/com-support/`, usable on
  Windows with zero configuration;
* **Project copy wins**: when `./packages/com-support/` exists it is used instead (so you
  can pin a version or swap the implementation); `over install --with-com` installs it;
* **Inspectable**: the last line of `openvbs over list` is `COM support: <origin>`, and
  `openvbs over info com-support` / `openvbs over com --json` show directory, origin,
  platform and availability;
* **Removable**: `openvbs over remove com-support` deletes the project copy and falls back
  to the vendor copy;
* **Switchable off**: `OPENVBS_COM=0` disables it entirely
  (`OVPP_COM_PATH=<dir>` points at a custom implementation directory);
* **No import needed**: `CreateObject` / `GetObject` find the adapter internally, so
  scripts only ever contain VBS syntax.

`extras/com-support/` is the **distributable version** of the same package (`over.json` +
`com_support.py` + README); it can be installed with
`over add com-support file:../extras/com-support` or published to a registry. The vendor and
extras copies are byte-identical, and `tests/test_com.py` enforces that.

## 10. Known limitations

| Limitation | Explanation |
|---|---|
| No version selection | A package has exactly one source; there is no "pick the version that satisfies all constraints" |
| No remote registry | Dependencies must resolve to a local path or Git URL (Python dependencies via pip are the exception) |
| No integrity signatures | `hash` guards against accidental drift, not malicious replacement |
| Platform differences are only "skip" | Mismatched packages are simply not installed; no per-platform dependency table |
| Python dependencies share an environment | They install into the current interpreter (venv isolation); no per-project site-packages |

All of these are "do it when a real need shows up" items, recorded in the deferred list of
`dev-docs/ROADMAP.md`.

## 11. Capabilities added in v0.9.13

| Capability | API | Description |
|---|---|---|
| Safe preview | `over.plan_install(root, source_map)` / `over install --dry-run` | resolve only, no writes; prints each dependency's action (install / unchanged) |
| Search | `over.search(term, root, registry)` / `over search <term>` | case-insensitive search over the local registry (`~/.openvbspp/over/registry.json`) plus installed packages |
| Tag enumeration | `over.git_tags(spec)` | parses `git ls-remote --tags`, strips the `v` prefix and `^{}`, sorts by version |
| Automatic version choice | `over.select_git_ref(spec, version_spec)` | picks the newest tag satisfying the constraint |
| Reproducible commit | lock `commit` field | for git checkouts, `over.lock` records the 40-hex HEAD SHA |

Tests: `tests/test_over.py` gained `TestOverDryRun` / `TestOverSearch` / `TestOverGitTags` /
`TestOverLockCommit`; `tests/test_over_python.py` covers Python dependencies;
`tests/test_com.py` covers the COM bridge.
