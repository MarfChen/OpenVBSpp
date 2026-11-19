# OpenVBS++ 构建与发布

> **In English** — How to build and publish: `python -m build`, `twine upload`
> to TestPyPI then PyPI, and `pyinstaller` for a single-file `openvbs.exe`.
> Includes a pre-release checklist and the five places the version number
> must agree.


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

# 2) 版本号四处一致
Select-String -Path OpenVBSPP\__init__.py,OpenVBSPP\ast_nodes\__init__.py,`
  OpenVBSPP\ovpprt\__init__.py,pyproject.toml -Pattern "0\.\d+\.\d+"

# 3) 干净构建（先删掉旧产物）
Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue
```

版本号规则：`OpenVBSPP/__init__.py`、`ast_nodes/__init__.py`、`ovpprt/__init__.py`、
`pyproject.toml` **四处必须一致**（测试里也有约定）。

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
├── openvbspp-0.9.14-py3-none-any.whl     # 纯 Python，无平台依赖
└── openvbspp-0.9.14.tar.gz               # sdist（含 examples/tests/doc）
```

本地自检：

```powershell
# 元数据是否合法
.\venv\Scripts\python.exe -m twine check dist\*

# 装到临时环境里试一下
python -m venv .tmp-venv
.tmp-venv\Scripts\python.exe -m pip install dist\openvbspp-0.9.14-py3-none-any.whl
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
- `excludes` 去掉了 `tkinter` / `astpretty` / `unittest`，减小体积；
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
