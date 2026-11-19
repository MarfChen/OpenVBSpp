# `over` 包管理器 · Package Manager

> **In English** — `over` is an offline-first, project-local package manager.
> Its sources are local folders, `.ovppkg` archives (zip) and `git+` URLs; it
> never requires a central server. Installed packages land in `./packages/`,
> a lock file records exact versions and content hashes, and `<namespace
> src="pkg">` resolves package names through this directory.
> Packages may also declare **Python dependencies** (`python` field): `over
> install` hands them to `pip` in one batch and records what really landed in
> `over.lock`. On Windows a bundled `com-support` package makes `CreateObject`
> talk to real COM without any import.

## 1. 定位

| 维度 | 选择 | 理由 |
|---|---|---|
| 仓库 | **无中央服务器** | 维护者没有服务器；本地目录 / Git / 单文件包足够分发 |
| 安装位置 | 项目内 `./packages/` | 不污染全局环境，多项目互不影响 |
| 版本约束 | semver 子集 `^1.2` `>=1.0` `<2` `=1.2.3` | 够用，不做 SAT |
| 锁文件 | `over.lock`（版本 + 内容哈希） | 可重现；配合 Git 提交 |
| 依赖 | 只用标准库（git 源调用系统 `git`） | 保持"零第三方依赖" |
| Python 依赖 | 交给 `pip` 解析，不自己实现 wheel/编译 | pip 已经解决了传递依赖、wheel、平台标签 |
| 平台差异 | 清单里 `platforms` 声明，不匹配就跳过 | 比 optionalDependencies 简单，够用 |

## 2. 清单

```json
{
  "name": "app",
  "version": "0.1.0",
  "entry": "main.vbs",
  "kind": "over",
  "description": "示例项目",
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

* `source` 支持三种：`file:<目录>`、`file:<x>.ovppkg`、`git+<url>[#ref]`；
* 字符串写法 = 只有来源、不限版本；对象写法可加 `version` 约束；
* `python` 是本包需要的 **PyPI 包**：`{"包名": "版本约束"}`，也接受
  `["requests>=2.31", "pywin32"]` 这种 pip 需求串写法；
* `platforms` 是适用平台（`win32` / `windows`、`darwin` / `macos`、`linux`，可带版本前缀
  如 `win32`、`darwin`）；为空表示全平台。平台不匹配的包**不安装、也不报错**，
  这样同一个 `over.json` 可以在多平台上共用；
* `kind`、`description` 只用于 `over list` / `over info` 展示。

## 3. 命令

| 命令 | 作用 |
|---|---|
| `openvbs over` | 打印用法（没有 `over.json` 也不会报错退出） |
| `openvbs over install` | 解析依赖图 → 安装到 `./packages/` → 装 Python 依赖 → 写 `over.lock` |
| `openvbs over install --frozen` | 严格按 `over.lock` 还原（字节级不变），并钉住锁里的 Python 版本 |
| `openvbs over install --dry-run` | 只解析依赖图，打印每条依赖的 action，不写盘 |
| `openvbs over install --no-python` | 跳过 Python 依赖解析（离线/受限环境） |
| `openvbs over install --with-com` | 额外把内置的 `com-support` 装进 `./packages/` |
| `openvbs over update` | 清缓存、删 `packages/`、重新解析安装 |
| `openvbs over list` | 列出 over 包 / bundled 包 / Python 依赖，末行显示 COM 支持来源 |
| `openvbs over info <name>` | 显示单个包的版本、来源、依赖与（对 com-support）运行时状态 |
| `openvbs over remove <name>` | 卸载 over 包（从 `packages/`、锁、`over.json` 里摘掉） |
| `openvbs over remove <name> --python` | 卸载 Python 依赖（走 `pip uninstall`，并清锁里的 `python` 条目） |
| `openvbs over com [--json]` | 显示 COM 桥当前用的目录、来源、平台与可用性 |
| `openvbs over add <name> <source>` | 写进 `over.json` 并立即安装 |
| `openvbs over search <term>` | 搜本地 registry 与已安装包 |
| `openvbs over pack <dir> [out]` | 把包目录打成 `.ovppkg`（zip，忽略 `.git` 与 `packages/`） |
| `openvbs over publish` / `register` / `index` | 本地 registry 发布、登记、生成索引 |

## 4. 依赖解析

1. 从根 `over.json` 出发深度优先遍历依赖图；
2. **循环依赖**在遍历时立即报错并打印链条：`Circular dependency: app -> a -> b -> a`；
3. 每个包收集**全部**版本约束和来源链；
4. 安装时用来源里的实际版本校验所有约束，冲突时报：

```
Conflict for 'shared': source provides 1.5.0, but '^2.0' is required by app -> b
```

5. 同一个包只安装一次（先到先得），后续需求只做版本校验；
6. **Python 依赖**在整张 over 依赖图上一起收集（含传递依赖）：同名包的约束用逗号合并
   （`six>=1.16,<2`），然后**一次性**交给 `pip install`，由 pip 解析传递依赖；
7. pip 用 `--report` 输出真实安装结果，`over.lock` 记录的是**实际装上的版本**，
   不是约束；这样在别的机器上 `--frozen` 能钉住同一份闭包。

## 5. 与编译器的衔接

安装后的包可以直接按名字引用，不需要写路径：

```vb
<namespace src="text-utils" alias="tu">
echo tu.Trim(" x ")
```

编译器解析 `<namespace src>` 时先查 `./packages/<name>/`（读该包的 `over.json`
里的 `entry`），找不到再退回相对路径。所以**包管理器不需要改编译器**。

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

* `hash` 是包目录内容（路径 + 字节）的 SHA-256 前 16 位，用来发现"版本号没变但内容变了"；
* `extras` 记录随包分发的 **bundled** 包（目前只有 `com-support`），来源写 `bundled`；
* `python` 是 pip 真实装出来的闭包：`version` 是实际版本，`requested` 表示这条是清单里
  显式声明的（`true`）还是 pip 顺带装的传递依赖（`false`），`required_by` 是提出约束的包；
* `version` 是锁格式版本；v1 的锁仍可读（缺 `extras` / `python` 时按空处理）；
* 建议把 `packages/` 加进 `.gitignore`，`over.lock` 提交进仓库。

## 7. 里程碑对照

| 阶段 | 内容 | 状态 |
|---|---|---|
| M1 | over.json 读写 + 本地目录安装 + 锁文件 | ✅（锁文件 v0.9.11 起可读，--frozen 精确恢复） |
| M2 | 语义化版本 + 约束求解（DFS + 冲突报错） | ✅ |
| M3 | .ovppkg 打包/解包 + Git 源 | ✅ |
| M4 | `<namespace src="包名">` 接线（编译器侧） | ✅ |
| M5 | Python 依赖（`python` 字段 → pip → 锁里记闭包） | ✅ v0.9.14 |
| M6 | 平台条件（`platforms`）+ 随包分发 bundled 包 | ✅ v0.9.14 |

## 8. Python 依赖

清单里写：

```json
{"python": {"pywin32": ">=306"}}
```

`over install` 会：

1. 遍历整张 over 依赖图，把每个包的 `python` 需求收进一个集合（同名约束逗号合并）；
2. 检查 `pip` 可用（`<python> -m pip --version`），不可用且确实有需求才报错，
   报错信息会提示 `--no-python`；
3. 执行 `pip install <需求...> --report <临时 json>`；
4. 从 report 里读出**实际安装结果**（含 pip 自己带出来的传递依赖）写进 `over.lock`
   的 `python` 数组。

`--frozen` 时反过来：读锁里的 `python`，拼成 `name==version` 钉住安装，
再校验安装结果与锁一致。`over remove <name> --python` 走 `pip uninstall`。

`openvbs over list` 会把 Python 依赖单独列出来（含 `required by ...`），`over info <name>`
能看到每个包声明的 `python_dependencies`。

## 9. COM 支持（`com-support` 包）

Windows 上 `CreateObject("WScript.Shell")` 这类 COM 对象由 `com-support` 包提供，
它封装 `pywin32`，对脚本完全透明：

* **随包分发**：`OpenVBSPP/vendor/com-support/` 里有一份，Windows 上零配置即可用；
* **项目内优先**：`./packages/com-support/` 存在时优先用它（可以钉版本或换实现），
  `over install --with-com` 会把它装进项目；
* **可查看**：`openvbs over list` 末行 `COM support: <origin>`，`openvbs over info com-support`、
  `openvbs over com --json` 都能看到目录、来源、平台与可用性；
* **可卸载**：`openvbs over remove com-support` 删掉项目副本，自动退回 vendor 副本；
* **可关闭**：`OPENVBS_COM=0` 彻底关闭（`OVPP_COM_PATH=<目录>` 可指定自定义实现目录）；
* **不需要 import**：`CreateObject` / `GetObject` 内部自己找到适配层，脚本里只有 VBS 语法。

`extras/com-support/` 是同一个包的**可分发版本**（`over.json` + `com_support.py` +
README），可以直接 `over add com-support file:../extras/com-support` 或打进 registry。
vendor 与 extras 两份文件逐字节一致，`tests/test_com.py` 会校验。

## 10. 已知限制

| 限制 | 说明 |
|---|---|
| 没有版本选择 | 一个包只能有一个来源，不能"从多个版本里挑一个满足所有约束的" |
| 没有远程仓库 | 依赖必须能落到本地路径或 Git URL（Python 依赖走 pip，是例外） |
| 不做完整性签名 | `hash` 只防意外漂移，不防恶意替换 |
| 平台差异只做"跳过" | `platforms` 不匹配的包直接不装，没有 per-platform 依赖表 |
| Python 依赖共享环境 | 装在当前解释器里（venv 隔离），没有 per-project 的 site-packages |

这些都是"等真实需求出现再做"的部分，记录在 [ROADMAP.md](ROADMAP.md) 的推迟项里。

## 11. v0.9.13 补齐的能力

| 能力 | API | 说明 |
|---|---|---|
| 安全预览 | `over.plan_install(root, source_map)` / `over install --dry-run` | 只解析依赖图、不写盘；输出每条依赖的 action（install / unchanged） |
| 搜索 | `over.search(term, root, registry)` / `over search <term>` | 搜本地 registry（`~/.openvbspp/over/registry.json`）+ 已安装包，大小写不敏感 |
| 版本枚举 | `over.git_tags(spec)` | `git ls-remote --tags` 解析标签，去掉 `v` 前缀与 `^{}`，按版本排序 |
| 自动选版本 | `over.select_git_ref(spec, version_spec)` | 在标签里挑满足约束的最新一个 |
| 可重现的 commit | 锁文件 `commit` 字段 | 来源是 git 检出时，`over.lock` 记录 40 位 HEAD SHA |

测试：`tests/test_over.py` 新增 `TestOverDryRun` / `TestOverSearch` / `TestOverGitTags` /
`TestOverLockCommit`；`tests/test_over_python.py` 覆盖 Python 依赖；`tests/test_com.py`
覆盖 COM 桥。
