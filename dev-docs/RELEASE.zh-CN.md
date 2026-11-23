# OpenVBS++ 构建与发布
> **[English](RELEASE.md)** | 中文

> **摘要** —— 两条发布路径：PyPI 包（pip install openvbspp）与单文件 exe；
> 含发版前检查、版本号五处一致、TestPyPI/正式 PyPI 上传、exe 打包要点、CI 与徽章说明。


本文覆盖两条发布路径：

1. **PyPI 包**（`pip install openvbspp`）—— 给 Python 用户；
2. **单文件 exe**（`openvbs.exe`）—— 给没装 Python 的 Windows 用户。

---

## 0. 一次性准备

```powershell
# 打包工具（构建 wheel / 上传 / 打 exe）
.\venv\Scripts\python.exe -m pip install -U build twine pyinstaller
```

PyPI 凭证建议放在 `%USERPROFILE%\.pypirc`：

```ini
[distutils]
index-servers =
    pypi
    testpypi

[pypi]
username = __token__
password = pypi-<你的 API Token>

[testpypi]
repository = https://test.pypi.org/legacy/
username = __token__
password = pypi-<TestPyPI 的 Token>
```

---

## 1. 发版前检查

```powershell
# 1) 全量测试必须全绿
.\venv\Scripts\python.exe -m unittest discover -s tests

# 2) 版本号五处一致
Select-String -Path OpenVBSPP\__init__.py,OpenVBSPP\ast_nodes\__init__.py,`
  OpenVBSPP\ovpprt\__init__.py,OpenVBSPP\plugins\__init__.py,`
  pyproject.toml -Pattern "0\.\d+\.\d+"

# 3) 干净构建（先删掉旧产物）
Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue
```

版本号规则：`OpenVBSPP/__init__.py`、`ast_nodes/__init__.py`、`ovpprt/__init__.py`、
`plugins/__init__.py`、`pyproject.toml` **五处必须一致**（测试里也有约定）。

---

## 2. 构建 wheel + sdist

```powershell
.\tools\build.ps1 -Target wheel
```

等价于：

```powershell
.\venv\Scripts\python.exe -m build
```

产物：

```
dist/
├── openvbspp-0.9.20-py3-none-any.whl     # 纯 Python，无平台依赖
└── openvbspp-0.9.20.tar.gz               # sdist（含 examples/tests/doc）
```

本地自检：

```powershell
# 元数据是否合法
.\venv\Scripts\python.exe -m twine check dist\*

# 装到临时环境里试一下
python -m venv .tmp-venv
.tmp-venv\Scripts\python.exe -m pip install dist\openvbspp-0.9.20-py3-none-any.whl
.tmp-venv\Scripts\python.exe -m OpenVBSPP examples\control1.vbs
.tmp-venv\Scripts\python.exe -m OpenVBSPP --version
```

---

## 3. 上传 PyPI

先发 TestPyPI 验证，再发正式：

```powershell
# 3.1 TestPyPI
.\venv\Scripts\python.exe -m twine upload --repository testpypi dist\*

# 3.2 从 TestPyPI 安装验证（注意要带 --index-url 和 --extra-index-url）
python -m pip install --index-url https://test.pypi.org/simple/ `
  --extra-index-url https://pypi.org/simple/ openvbspp

# 3.3 正式 PyPI
.\venv\Scripts\python.exe -m twine upload dist\*
```

发布后检查：

- <https://pypi.org/project/openvbspp/> 页面、README 渲染、classifier；
- `pip install openvbspp` 后 `openvbs --version` 正常；
- 同一个版本号**不能重复上传**（只能升级版本号或使用 `--skip-existing`）。

---

## 4. 打包 exe

```powershell
.\tools\build.ps1 -Target exe
```

产物：`dist\openvbs.exe`（单文件，双击或用命令行运行）：

```
openvbs.exe examples\control1.vbs
openvbs.exe --version
openvbs.exe            # 无参数跑内置示例
```

实现要点（`packaging/openvbs.spec`）：

- 入口是 `packaging/entry.py`，只调 `OpenVBSPP.cli:main`；
- 生成代码在运行期才 `import OpenVBSPP.ovpprt.*`，静态分析看不到，
  所以 spec 里用 `hiddenimports` 把这些模块全列出来（新增运行时模块时**记得加**）；
- `com-support` 是**数据文件**（运行期用 `spec_from_file_location` 按路径加载），
  静态分析发现不了，spec 里必须用 `datas` 把 `OpenVBSPP/vendor/com-support`
  打进包，否则 exe 里 `CreateObject("WScript.Shell")` 会退化成「COM 不可用」；
- `com_support.py` 里的 `import win32com` 同样静态不可见：打包时若装了 pywin32，
  spec 会自动补 `pythoncom` / `pywintypes` / `win32com.*` 到 `hiddenimports`，
  并把 `pywin32_system32/*.dll` 作为 `binaries` 打进去——所以 **exe 自带 COM 能力**，
  目标机器不用另装 pywin32；没装 pywin32 也能打包，只是 exe 里 COM 不可用；
- `excludes` 去掉了 `astpretty` / `unittest` / `pydoc`，减小体积。自 v0.9.22 起
  **不再排除 `tkinter`**：exe 自带真正的 `MsgBox` / `InputBox` 对话框（未压缩 +8.9 MB），
  `OpenVBSPP/vendor/ui-support` 与 `OpenVBSPP.ovpprt.ui` 和 COM 一样在 spec 里声明。
  `WScript` 宿主对象（v0.9.23）对应的 `OpenVBSPP.ovpprt.host` 也在同一清单里；
- 第三方许可文本随 exe 一起分发：`THIRD_PARTY_LICENSES.md`、`NOTICE` 与本机 Tk 的
  `license.terms` 都放进 `datas` 的 `licenses/` 目录。Tcl/Tk 许可要求声明
  「appear in all copies」，因此发布 Release 时也要把 `THIRD_PARTY_LICENSES.md`
  放在 exe 旁边；
- 想要更小的体积可以再装 `upx`，spec 里 `upx=True` 已经打开。

---

## 5. 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| `twine` 报 `File already exists` | 该版本已发布，改版本号重新构建 |
| PyPI 页面 README 不渲染 | `README.md` 里有 PyPI 不支持的语法（如部分 HTML 注释） |
| exe 运行时 `ModuleNotFoundError: OpenVBSPP.ovpprt.xxx` | 新增运行时模块后忘记加进 `hiddenimports` |
| exe 被杀软误报 | PyInstaller 单文件的常见问题，可改用 `--onedir`（改 spec 里的 `EXE`/`COLLECT`） |
| wheel 里缺 examples | sdist 有 `MANIFEST.in` 控制，wheel 只装包本身，属预期行为 |

---

## 6. 发版说明（Release Notes）

PyPI / GitHub Release 上贴的正文从这里取，写完记得和 README 的「更新日志」与
`dev-docs/ROADMAP.md` 的版本表对齐。

### v0.9.20 — Python 互操作显式导入、文档英文化

- Python 互操作（`#python ... #endpython` 与 `COP`）不再由 `Option COP` 打开：
  必须 `<script src="python-embed" />`（随包分发的 over 包，可用 `over add python-embed` 装进项目）。
  `Option COP` 现在是带迁移提示的语法错误。
- `plugins/debug_tools.py`（`ast2str`）删除；`--save` 仍写 `.ast.txt`，改用标准库 `ast.unparse`。
- 文档改为英文为主 + 每篇 `.zh-CN.md` 对照；维护者文档移到 `dev-docs/`，不再进文档站。
- `over list` 末尾列出随包分发的内置包；`over add <name>` 可省略来源；`over` 半残 git 缓存会自愈。

### v0.9.15 — 性能

纯性能版本：**只改生成的 Python 代码，不改语言语义**，全部由
`tests/test_perf_semantics.py` 钉住边界。

- **启动懒加载**：`ovpprt/regexp.py`、`jsonobj.py`、`object.py`、`builtins.py` 里的
  `re`、`json`、`datetime`、`math`、`random` 改成首次使用时才导入；`cli.py` 只在真正
  执行脚本时才 `import OpenVBSPP.runtime`。冷启动开销 27 ms → 13 ms。
- **只读 `ByRef` 参数不再装箱**：编译器静态判断形参是否真的被写（赋值、`ReDim`、`For`、
  或继续传给下层），只读的按值传；被调方内联拆包 `if type(n) is Ref: n = n.v`，
  兼容 `namespace` 跨模块调用与 `AddressOf` 委托仍传 `Ref` 的情况。
  200k 次函数调用 118 ms → 23 ms（13.7× → 1.35× 手写 Python）。
- **无 `On Error` 的过程不再建错误帧**：过程体里没有 `On Error ...` / `Resume` / `Err`
  就不生成 `enter()/try/finally/leave()`。递归 fib(24) 67 ms → 38 ms（12.0× → 6.85×）。
- **`For` 安全时编译成 `range`**：步长是常量、循环体不给计数器赋值时，生成
  `for i in for_seq(...)` + `else: i = i + step`，复现「正常结束后计数器 = 末值 + 步长、
  `Exit For` 不加、空区间保留起点」三条 VBScript 语义。1M 循环 50 → 29.7 ms（0.99× 手写 Python）。

其它：用例数 457 → 485；`tools/bench.py` 的说明与 README 性能小节同步更新。

### v0.9.14 — `over` 的 Python 依赖与 COM 支持

- `over.json` 新增 `python` 字段：整个依赖图的 Python 依赖一次性交给 pip 解析，
  `over.lock`（v2）记录真实闭包与 `required_by`；`over list` / `info` / `remove --python`
  可查看与卸载。
- `over.json` 新增 `platforms` 字段；不匹配的平台只跳过，不报错。
- COM 支持做成 `over` 包 `com-support`（封 `pywin32`）：随包分发（Windows 上零配置可用），
  项目里装了就用项目的，`over install --with-com` 可装进项目、`OPENVBS_COM=0` 可关；
  脚本里照旧 `CreateObject("WScript.Shell")`，**不需要 import**。
- 修掉 `object.py` 三处 `Dictionary` 成员分派里的未定义变量。

---

## 7. CI（GitHub Actions）

有两个工作流：`.github/workflows/ci.yml`（测试 + exe）与
`.github/workflows/docs.yml`（文档站，见 §7.1）。

**`test`（每次 push / PR / 手动触发）**

| 维度 | 值 |
|---|---|
| Windows（硬门禁） | Python 3.10 / 3.11 / 3.12 / 3.13 |
| Linux（硬门禁） | Python 3.10 / 3.13，`continue-on-error` 在 v0.9.15 已移除 |
| 步骤 | toolchain → `pip install -e ".[dev]" --no-build-isolation` → 装 pywin32（仅 Windows）→ `coverage run -m unittest discover -s tests` → `coverage report`（打印在 job 日志里）→ 入口点自检（含 `-c` 内联）→ COM 冒烟（仅 Windows）→ `python -m build` → 干净 venv 装 wheel 冒烟（含 `vendor/com-support` 存在性断言与 `-c` 冒烟）→ 上传 `dist` 产物 |

Windows 之所以是硬门禁：语言支持的 COM、WScript 语义、路径行为都在 Windows 上验证过。
Linux 从 v0.9.15 起同样是**硬门禁**（3.10 与 3.13 两个点），矩阵里已没有
`experimental` 这一维；如果某个 Linux 版本临时不稳定，再给它单独加回
`continue-on-error: true` 即可。

**`exe`（只在打 tag 或手动触发时跑）**

PyInstaller 构建 `dist/openvbs.exe`，再把它拷到临时目录做冒烟（`--version` + 真实
`CreateObject("WScript.Shell")`），最后按架构上传为 `openvbs-exe-x64` /
`openvbs-exe-arm64` 产物。矩阵是
`[{runner: windows-latest, arch: x64}, {runner: windows-11-arm, arch: arm64}]`；
注意 `windows-11-arm` 运行器**只对公共仓库开放**，私有仓库会直接失败。

### 完成 CI 还需要你本人做的事

1. **建仓库并 push**（必须有这一步，否则 Actions 无处运行，徽章只会是灰色）：
   ```bash
   git remote add origin https://github.com/<你>/OpenVBSpp.git
   git push -u origin main --tags
   ```
2. 仓库 Settings → Actions → General：确认允许 workflow（新仓库默认允许，无需改）。
3. **不需要任何 secret** 就能跑测试。只有下面这件事需要额外配：
   - 想在 CI 里发 PyPI：加 `PYPI_API_TOKEN` secret，再加一个 `twine upload` 的 job。
   覆盖率每次 CI 都会测并打印在 job 日志里；README 的覆盖率徽章是**静态** shields.io 字符串
   （本地实测得出），不需要接任何第三方服务。
4. 首次 push 后约 3–5 分钟出结果：6 个测试矩阵任务（4 Windows + 2 Linux）+ 2 个手动 exe 任务。
   全绿之后 CI 徽章自动变绿，**不需要改 README**。
5. 想让文档站上线：Settings → Pages → Build and deployment → Source 选
   **GitHub Actions**（默认是 Deploy from a branch，不选这个 `docs.yml` 会失败）。

---

## 8. 徽章（怎么变绿）

README 首屏有 7 个徽章，分两类：

| 徽章 | 类型 | 变成 / 保持绿色要做什么 |
|---|---|---|
| `tests-649 cases` | 静态（shields.io 手工字符串） | 一直是绿的；**改了用例数必须同时改 `README.md` 与 `README.zh-CN.md`**，否则是假绿 |
| `version-0.9.29` | 静态 | 发版时与五处版本号一起改（README 两个文件 + 两处正文） |
| `license-Apache--2.0` | 静态 | 常绿 |
| `python-3.10+` | 静态 | 常绿；将来放宽/收紧 requires-python 时同步 |
| `dependencies-none` | 静态 | 常绿（`pywin32` 是可选 extra，不算运行期依赖） |
| **CI** | **动态** | 仓库 push 后，`main` 分支上有一次成功的 workflow 运行即变绿；显示 `no status` 说明默认分支还没有运行记录 |
| `coverage-80%` | 静态 | 用 `coverage run --source=OpenVBSPP,cli -m unittest discover -s tests` 在本地实测；数字变化时同步改两个 README |

静态徽章 = 项目自称，CI 徽章 = 机器实测，两者都保留；评审者更看后者。

发布到 PyPI 之后可以再补两个动态徽章（现在加会是红色 "package not found"，所以先不加）：

```markdown
![pypi](https://img.shields.io/pypi/v/openvbspp)
![pyversions](https://img.shields.io/pypi/pyversions/openvbspp)
```

覆盖率从 v0.9.15 起就在 CI 里测：`coverage run --source=OpenVBSPP,cli` 后跟一个
`coverage report`，结果打印在 job 日志里。从 v0.9.29 起 README 用的是**静态**
`coverage-80%` 徽章，不再依赖 Codecov，也就不需要 token 或第三方服务；
数字变化时在本地重跑一次命令并同步改两个 README。

---

## 9. Python 版本、产物与 ARM

### 9.1 支持到多老的 Python

| 项 | 现在的状态 |
|---|---|
| 声明 | `requires-python = ">=3.10"`，classifiers 列 3.10 – 3.13 |
| 静态校验 | `vermin --eval-annotations --target=3.9 --violations OpenVBSPP cli.py tests` → **零违规**，最低需求 3.9（无 `match`、无 `X | Y` 注解、无 `tomllib`、无 `except*`） |
| 运行时实测 | Windows embeddable **3.9.13** 与开发 venv **3.13.15** 各跑 485 个用例**全绿**（3.9 属于兼容性实测，已不在声明范围内） |
| CI | Windows 3.10 / 3.11 / 3.12 / 3.13 硬门禁；Linux 3.10 / 3.13 硬门禁 |

**已经做的调整（v0.9.15）**：3.9 已于 2025-10 结束安全支持，所以声明下限提到 **3.10**。
**0.9.14 是最后一个声明支持 3.9 的版本**；运行期代码本身仍然 3.9 兼容
（vermin 的静态结论没变，只是不再承诺），所以愿意的人仍可以在 3.9 上跑。
将来再提下限，只需同步三处：`pyproject.toml` 的 `requires-python` 与 classifiers、
两个 README 的 Python 徽章。

### 9.2 PyPI 上应该传什么

只传两类文件，**都是纯 Python**：

```powershell
python -m build                        # dist\openvbspp-0.9.20-py3-none-any.whl + dist\openvbspp-0.9.20.tar.gz
python -m twine check dist/*           # 元数据 + README 渲染检查
python -m twine upload --repository testpypi dist/*   # 先发 TestPyPI 验证
python -m twine upload dist/*                          # 再发正式 PyPI
```

- **不要**为 3.9 / 3.10 / … 分别打包：本项目没有 C 扩展、没有 ABI 绑定，一个
  universal wheel（`py3-none-any`）就覆盖所有 Python 版本与所有 CPU 架构。
- **exe 不要传 PyPI**：它是按 OS / 架构区分的自包含可执行文件（约 8 MB），属于
  GitHub Releases 的产物，不是 PyPI 的包格式。
- 推荐用 [Trusted Publishing](https://docs.pypi.org/trusted-publishers/)（OIDC）替代 API token；
  要用 token 就把 `PYPI_API_TOKEN` 放进仓库 secret，不要写进代码或 workflow 明文。

### 9.3 ARM（Windows ARM64 / Linux arm64）

拆成两件事看，结论完全不同：

1. **`pip install` 到 ARM 机器：什么都不用做。** universal wheel 在任何架构上都能装；
   Windows ARM64 上真正的风险点 `pywin32` 已经有 `win_arm64` 轮子
   （实测 pywin32 **312** 提供 cp39–cp315 的 `pywin32-312-cpXX-cpXX-win_arm64.whl`），
   所以 `pip install "openvbspp[com]"` 在 ARM64 Windows 上照常可用。
2. **单文件 exe：需要一台 ARM 机器来编译。** PyInstaller **不能交叉编译**，产物架构
   等于构建机架构。两条路：
   - **GitHub 托管的 ARM 运行器**：公共仓库把 `runs-on` 写成 `windows-11-arm` 即可
     （2025-04 进入公共预览，2025-08 对公共仓库正式可用；**私有仓库不支持这个标签**，
     写了会直接失败）。给 `ci.yml` 的 `exe` 任务加一行矩阵：
     ```yaml
     exe:
       strategy:
         matrix:
           os: [windows-latest, windows-11-arm]
       runs-on: ${{ matrix.os }}
     ```
     产物名建议带架构，例如
     `openvbs-${{ matrix.os == 'windows-11-arm' && 'arm64' || 'x64' }}.exe`。
   - **本机 ARM 设备**（Surface / Snapdragon X / Dev Kit）：直接
     `python -m PyInstaller packaging/openvbs.spec --noconfirm`，产物就是 arm64。
   Linux arm64 同理用 `ubuntu-24.04-arm` 运行器；不过 exe 本身只面向 Windows，
   Linux 分发走 `pip install` 即可。

### 7.1 文档站（GitHub Pages）

文档站从 v0.9.15 起可发布：仓库根的 `mkdocs.yml` 指向 `doc/` 作为站点根目录，
`doc/index.md` 是落地页，`.github/workflows/docs.yml` 在 `doc/**` 或 `mkdocs.yml`
有变动时构建并部署到 GitHub Pages（PR 只构建不部署）。

```powershell
pip install "mkdocs-material>=9.5"     # 或 pip install -e ".[docs]"
mkdocs serve                          # 本地预览 http://127.0.0.1:8000
mkdocs build --site-dir site          # 生成静态站（site/ 已进 .gitignore）
```

`doc/` 里的 `.docx` 会被原样拷进站点，`.md` 的目录结构由 `nav` 决定；新增一篇文档时，
记得同时把它加进 `mkdocs.yml` 的 `nav`，否则只能靠搜索找到。

### 9.4 与版本 / 架构相关的发版自检

```powershell
python -m unittest discover -s tests                                    # 649 OK
python -m twine check dist/*
pip download --only-binary=:all: --no-deps openvbspp==0.9.20 -d tmp    # 必须拿到 py3-none-any.whl
```

拿到的若是 `py3-none-any.whl`，就说明这个包对 Python 版本和 CPU 架构都没有额外要求。

### 9.5 打包元数据卫生（v0.9.15 已达标）

`python -m build` 现在**零警告**，改动如下：

- 许可证元数据改成 PEP 639 写法：`license = "Apache-2.0"` +
  `license-files = ["LICENSE", "NOTICE"]`，`build-system.requires` 提到
  `setuptools>=77`；删掉已弃用的 `License :: OSI Approved` classifier。
  wheel 的 `dist-info/licenses/` 现在带 `LICENSE` 与 `NOTICE`，满足 Apache-2.0
  对分发件保留声明的要求。
- `MANIFEST.in` 去掉 Windows 上会被按 cp936 读坏的中文模式行（相关说明改写在
  `.gitignore` 与本文档里），并补齐 `packaging/*`、`tools/*.py`：
  这样从 sdist 也能重建 exe（`packaging/openvbs.spec`）并跑基准（`tools/bench.py`）。
  `doc/` 的规则是"只收 `.md` + 显式列出的规范"：README 引用的
  `doc/OpenVBS spec.docx` 在 sdist 内，内部规划随笔 `doc/plan.docx` 用
  `exclude` 挡掉，只留在仓库里。
- `packages` 列表补 `OpenVBSPP.ide.static` 与 `OpenVBSPP.vendor`，消除 setuptools 的
  “absent from the `packages` configuration” 提示。

自检：

```powershell
python -m build --outdir dist      # 应当没有任何 warning
python -m twine check dist/*
```
