# 内置网页 IDE · Built-in Web IDE

> **English** — `openvbs ide` starts a local, zero-dependency web IDE: edit, run,
> check, save and load VBS scripts in the browser. It is an **optional extension
> component**: the compiler itself never imports it, and it has no third-party
> dependencies. A package-manager slot is reserved but disabled.

## 1. 启动 / Start

```bash
openvbs ide                       # 默认 http://127.0.0.1:7860
openvbs ide --port 8000
openvbs ide --workspace ./scripts --no-browser
python -m OpenVBSPP.ide           # 等价写法
```

界面：左边编辑器（带行号、Tab 缩进），右边三个页签：

| 页签 | 内容 |
|---|---|
| 输出 | stdout + 运行期错误（带行号与源码行）+ 耗时 |
| 诊断 | 只编译不运行，`Option Typing` 的类型错误也在这里 |
| 扩展 | 能力矩阵：运行/检查/包管理器/语言服务器 |

快捷键：`Ctrl+Enter` 运行、`Ctrl+S` 保存、`Tab` 缩进。

## 2. 为什么不用 Gradio / Streamlit

评估过 Gradio、Streamlit、NiceGUI，最后选择了**标准库 HTTP + 单页前端**：

| 方案 | 体积 | 适合本项目的点 | 否决理由 |
|---|---|---|---|
| **内置（采用）** | 0 MB | 零依赖、离线可用、可测试、UI 完全可控 | 编辑器功能有限（无高亮） |
| Gradio | ~300 MB（含 pandas 等） | `gr.Code` 开箱即用 | 依赖过重，违背"运行时零依赖"的定位 |
| Streamlit | ~150 MB | 布局快 | 每次交互整页重跑，做 IDE 不合适 |
| NiceGUI | ~40 MB | Quasar 界面漂亮 | 仍是可选依赖，且需要启动 uvicorn |

这里的关键判断：**IDE 的核心是"服务层"，不是前端框架。** 所以先做
`service.py`（纯标准库、可单元测试），前端只是一层薄壳。以后想加 Gradio 版本，
只需要调用同一组函数，前端代码不进核心包。

> 如果确实需要 grayscale 的语法高亮，可以后续引入 CodeMirror 的静态文件
> （仍然是零 Python 依赖，只是多一个静态资源）。

## 3. 架构

```
OpenVBSPP/ide/
├── service.py      后端服务层（运行/检查/文件/示例/能力声明）   ← 唯一需要测试的部分
├── packages.py     包管理器接口（预留，纯契约）
├── server.py       标准库 http.server，只做路由与 JSON 编解码
├── static/index.html   单页前端（内联 CSS/JS，无构建步骤）
├── __init__.py     对外导出
└── __main__.py     python -m OpenVBSPP.ide
```

前端 **只通过 JSON API** 与后端交互，所以换任何前端都不用改后端：

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/capabilities` | 版本、能力矩阵、包管理器契约 |
| GET | `/api/examples` | 内置示例 |
| GET | `/api/scripts` | 工作目录内的脚本列表 |
| POST | `/api/run` | `{source, filename}` → `{ok, output, error, duration_ms}` |
| POST | `/api/check` | `{source}` → `{ok, diagnostics}` |
| POST | `/api/save` | `{name, source}` |
| POST | `/api/load` | `{name}` → `{source}` |
| POST | `/api/packages/search` | **501**（预留） |

## 4. 打开 / 保存：系统原生对话框

浏览器出于安全限制**不能**替网页弹出系统的另存为对话框，所以由本地服务端调用
Tk 的 iledialog 完成——IDE 只监听 127.0.0.1，弹窗出现在你自己的桌面上，
体验和本地程序一致。

| 按钮 | 行为 |
|---|---|
| 打开… | iledialog.askopenfilename（Windows/Linux/macOS 原生对话框） |
| 保存 | 有路径就原地写；没有路径走另存为 |
| 另存为… | iledialog.asksaveasfilename，默认扩展名 .vbs |

三级降级，保证任何环境都能用：

1. **原生对话框**（有 tkinter + 图形环境）；
2. **浏览器文件选择**（无 tkinter，比如最小化 Linux 容器）：<input type=file> + Blob 读入；
3. **工作目录**（无对话框且需要保存时）：存到 ~/.openvbspp/scripts。

GET /api/dialogs 会告诉前端当前用的是哪一级，扩展页签里也能看到。

相关接口：

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | /api/dialogs | {platform, native_dialogs, newline} |
| POST | /api/dialog/open | 弹打开框 → {ok, path, source} |
| POST | /api/dialog/save | 弹另存为框 → {ok, path} |
| POST | /api/file/open | 按绝对路径读 |
| POST | /api/file/save | 按绝对路径写 |

## 5. 安全边界

- 只监听 `127.0.0.1`（默认），不对外网开放；
- 文件读写被限制在工作目录内，`../` 逃逸会被拒绝（有测试覆盖）；
- 请求体上限 2 MB；
- **注意**：VBS 脚本本身可以执行任意代码（`#python` 块、`Execute`），
  所以 IDE 只适合在本地跑，不要部署成公网服务。

## 6. 包管理器接口（预留）

`packages.py` 现在只定义契约，不实现安装：

```python
class PackageProvider:          # 包源协议
    def search(self, keyword=""): ...
    def install(self, name, version=None): ...
    def uninstall(self, name): ...
    def installed(self): ...

class PackageInfo:              # 一个包长什么样
    name, version, description, provides, requires, homepage
```

接入方式（**前端代码不用改**）：

```python
from OpenVBSPP.ide.packages import set_provider, LocalPackageStore

class MyStore(LocalPackageStore):
    def search(self, keyword=""): ...      # 以后接 PyPI / 本地目录 / over 索引
    def install(self, name, version=None): ...

set_provider(MyStore())
```

`/api/capabilities` 会实时反映 `enabled` 状态，IDE 的"扩展"页签据此显示"预留中"
或"已启用"。CLI 侧真正的包管理入口是 `openvbs over ...`（见 [OVER.md](OVER.md)），
IDE 的 `/api/packages/*` 还没有接到 `over` 上。

## 7. 测试

```bash
python -m unittest tests.test_ide -v
```

覆盖：服务层（运行/错误结构/检查/示例全部可跑）、工作目录（保存/读取/路径逃逸）、
HTTP 冒烟（首页、能力、运行、预留接口返回 501）。
