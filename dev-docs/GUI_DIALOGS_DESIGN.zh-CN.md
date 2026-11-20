# MsgBox / InputBox 的 GUI 对话框（Tkinter）——设计与许可分析

> **[English](GUI_DIALOGS_DESIGN.md)** | 中文

> **摘要** —— 目前 `MsgBox` 把文本打印到 stdout 并返回 `1`，`InputBox` 从 stdin
> 读一行。能跑，但恰好破坏了 VBScript 代码真正依赖的东西：**返回值**。
> 写成 `If MsgBox("保存？", vbYesNo) = vbYes Then` 的脚本永远进不去「是」分支。
> 本报告核实 Python 的 `tkinter` 能否补上这个缺口，提出一个无需 import 的内置
> over 包方案，并给出两种分发形态（纯 Python wheel / 打包 exe）各自的许可义务。

---

## 1. VBScript 的「UI 控件」到底指什么

VBScript 本身只有**两个** GUI 入口：

| 语句 | 作用 |
| --- | --- |
| `MsgBox(prompt, buttons, title)` | 模态消息框，返回 `1`–`7` |
| `InputBox(prompt, title, default)` | 模态单行输入框，返回字符串 |

老脚本里其它看起来像「UI 控件」的东西都不是 VBScript，而是 COM 自动化：
`Shell.Application.BrowseForFolder`、`UserAccounts.CommonDialog`、
用 `InternetExplorer.Application` 做的 HTA 表单、带超时的
`WScript.Shell.Popup`。这些继续走 COM（`com-support`，v0.9.14 起），
Tkinter 不替代它们。

所以兼容目标很窄也很明确：**让这两个内置函数和 Windows 版行为一致，包括返回码。**

---

## 2. 本机可行性（实测）

| 事实 | 值 |
| --- | --- |
| Tcl / Tk 版本 | 8.6 / 8.6，patchlevel `8.6.13` |
| 窗口系统 | `win32` |
| 能否创建 Tk 根 | 能（隐藏 root，枚举到 443 个字体族） |
| 屏幕 | 1536×864，会话 `Console` |

`tkinter.messagebox` 提供 `showinfo`、`showwarning`、`showerror`、
`askquestion`、`askokcancel`、`askyesno`、`askyesnocancel`、`askretrycancel`。
其私有函数 `_show(title, message, _icon, _type, **options)` 把 `type`、
`icon`、`default`、`parent` 直接透传给 Tcl 命令 `tk_messageBox`，而后者支持
**全部六种** VBScript 按钮组合：

| VBScript 按钮 | 值 | `tk_messageBox -type` |
| --- | --- | --- |
| `vbOKOnly` | 0 | `ok` |
| `vbOKCancel` | 1 | `okcancel` |
| `vbAbortRetryIgnore` | 2 | `abortretryignore` |
| `vbYesNoCancel` | 3 | `yesnocancel` |
| `vbYesNo` | 4 | `yesno` |
| `vbRetryCancel` | 5 | `retrycancel` |

图标一一对应（`vbCritical`→`error`、`vbQuestion`→`question`、
`vbExclamation`→`warning`、`vbInformation`→`info`），默认按钮映射到
`-default`，返回字符串再映射回 VBScript 常量：

```
ok→1  cancel→2  abort→3  retry→4  ignore→5  yes→6  no→7
```

`InputBox` 映射到 `tkinter.simpledialog.askstring(title, prompt,
initialvalue=default)`，它在「取消」时返回 `None` —— VBScript 返回 `""`，
包装层做转换即可。所需能力全部具备，不引入任何第三方依赖。

**结论：可以，Tkinter 能忠实实现这两个内置函数。** 剩下的工作不是能力问题，
而是启用策略问题。

---

## 3. 包形态：`ui-support`，随包分发、无需 import

沿用 `com-support` 的先例，实现随运行时一起分发，脚本作者不需要 import：

```
OpenVBSPP/vendor/ui-support/over.json
OpenVBSPP/vendor/ui-support/ui_support.py     # 真正的实现
OpenVBSPP/ovpprt/ui.py                        # 内置函数用的薄分发层
extras/ui-support/…                           # 逐字节一致的镜像
```

`over.json`：

```json
{
  "name": "ui-support",
  "version": "1.0.0",
  "kind": "feature",
  "entry": "",
  "description": "MsgBox 与 InputBox 的 GUI 对话框（Tkinter）。",
  "capabilities": ["ui"],
  "dependencies": {}
}
```

解析顺序与 COM 相同：项目内 `packages/ui-support`（可钉版本）→ 随包分发的
`OpenVBSPP/vendor/ui-support/`，按文件 mtime 缓存。移除该包或设置
`OPENVBS_UI=console` 即恢复当前的纯控制台行为。

### 启用规则

只有同时满足下列条件，`MsgBox` / `InputBox` 才走 GUI：

1. `OPENVBS_UI` 未设置（默认 `auto`）或显式为 `gui`；
2. 进程是从 **CLI 入口** 启动的（`python -m OpenVBSPP file.vbs`、
   `openvbs file.vbs`、`openvbs.exe`），而不是通过嵌入 API `vbs_run()` ——
   这样库调用方和测试套件天然不弹窗；
3. `sys.stdin.isatty()`，且环境里没有 `CI`；
4. `tkinter` 能导入，且能创建隐藏的 `Tk()` 根。

任一条件不满足，或弹窗时 Tk 抛 `TclError`（服务会话、无显示），就回落到今天
的控制台行为，并把决定缓存到进程结束。显式 `gui` 失败时向 stderr 打一行警告
后安静退化。

`OPENVBS_UI=console`（别名 `off`）强制控制台；`OPENVBS_UI=gui` 即使没有 TTY
也强制 GUI（`pythonw.exe` 下有用）。

### 必须保留的语义

- `MsgBox` 精确返回 `1`–`7`；未知/组合按钮值退化为 `vbOKOnly`；
  `vbSystemModal` 让对话框置顶。
- 省略 title 时默认 `OpenVBS++`；prompt 就是消息正文。
- `InputBox` 返回输入的文本，点「确定」时返回预填的 `default`，点「取消」返回
  `""` —— 永远不会是 `None`。
- 对话框像真的一样阻塞，没有超时（VBScript 本来也没有）。

---

## 4. 打包影响

| 产物 | 代价 | 说明 |
| --- | --- | --- |
| sdist / wheel | **0 字节** | `ui-support` 约 5 KB Python；Tcl/Tk 来自用户自己的解释器 |
| Windows exe | **未压缩 +8.9 MB** | 本机实测：`tcl86t.dll` 1.82 MB、`tk86t.dll` 1.59 MB、`lib\tcl8.6` 3.14 MB、`lib\tk8.6` 1.56 MB、`Lib\tkinter` 0.72 MB、`_tkinter.pyd` 0.06 MB；UPX/PyInstaller 大约能压掉一半 |
| Linux 安装 | — | 文档注明 `apt install python3-tk`；没有它就走控制台回落 |

`packaging/openvbs.spec` 目前排除 `tkinter`；要让 exe 支持对话框，需要去掉该
exclude，并把 Tcl/Tk 许可文件加进 `datas`（见 §5）。也可以保留「控制台 exe」
另出一个「GUI exe」，但那会让发布矩阵翻倍，收益不大。

---

## 5. 许可证

### 到底分发了什么

| 场景 | 是否再分发 Tcl/Tk | 新增义务 |
| --- | --- | --- |
| `pip install openvbspp`（wheel/sdist） | 否 —— 由解释器提供 | 无 |
| Windows 单文件 exe | **是**（`tcl86t.dll`、`tk86t.dll`、脚本库） | 必须随附 Tcl/Tk 许可文本 |
| 我们自己的源码 | 只有自己的代码 | Apache-2.0（不变） |

### 组件与许可

| 组件 | 许可 | SPDX | 是否 OSI 认证 |
| --- | --- | --- | --- |
| CPython、`tkinter`、`_tkinter` | Python Software Foundation License 2.0 | `PSF-2.0` | 是 |
| Tcl/Tk 8.6 运行时 | Tcl/Tk License（BSD 风格） | `TCL` | **否** |
| libtommath（在 `tcl86t.dll` 内） | 公有领域 | — | — |
| zlib（在 Tcl/Tk 内） | zlib License | `Zlib` | 是 |
| pywin32（COM exe） | PSF-2.0 | `PSF-2.0` | 是 |
| X11/fontconfig（仅 Linux） | MIT / BSD 风格 | `X11` | 是 |

Tcl/Tk 许可文本位于 `tcl8.6/tk8.6/license.terms`（本机 2,267 字节、40 行），
是标准措辞：版权归 Regents of the University of California、Sun
Microsystems、Scriptics、ActiveState、Apple 等；允许 "use, copy, modify,
distribute, and license this software and its documentation for any purpose"，
**条件是** 版权声明与许可声明 "appear in all copies"，并且 "the name of the
copyright holders may not be used to endorse or promote products derived from
this software without specific prior written permission"。

关键性质（对照本机文本与 SPDX）：

- **宽松许可，非 copyleft** —— 允许闭源再分发；
- **无版税、无 copyleft、无需逐文件登记** —— 唯一义务是保留声明；
- **未经 OSI 认证**（SPDX `isOsiApproved: false`）。这是**采购政策**问题而不是
  法律风险问题；Fedora 将其归为 good license，且它是 FSF-free、与 GPL 兼容；
- 末尾一段是常见的美国政府 "Restricted Rights" 条款；
- 解释器自带的 `LICENSE_PYTHON.txt` **不含** Tcl/Tk 字样，所以两份声明互相
  独立，exe 里两份都必须存在。

### 具体义务

1. 在 exe 的数据目录里逐字附带 Tcl/Tk 许可文本，并在 `NOTICE` 中引用。
2. 新增 `THIRD_PARTY_LICENSES.md`（或 `LICENSES/`），收录两份文本（嵌入的
   CPython 用 PSF-2.0、Tcl/Tk 用 TCL），并从 `README` 的 Licensing 一节链接。
3. 保留 `NOTICE` 现有的 Apache-2.0 与 pywin32 条目。
4. 不要用 Tk 给产品命名（"OpenVBS Tk Edition"），也不要暗示版权方背书；写成
   "可选的 GUI 对话框使用 Python 的 `tkinter`（Tcl/Tk 8.6）"这类事实性描述属于
   合理使用，没问题。
5. 在安装文档里写明 Linux 需要 `python3-tk`，避免用户报「MsgBox 没反应」。

### 风险评估

| 风险 | 等级 | 理由 |
| --- | --- | --- |
| wheel/sdist 许可违规 | **无** | 没有再分发任何 Tcl/Tk 字节 |
| exe 许可违规 | **低** | 只需保留声明，§5.1–5.2 即可满足 |
| 企业「只允许 OSI 认证」策略拦截 | **低–中** | `TCL` 未经 OSI 认证；文档写明控制台回落、Tkinter 可选即可 |
| 商标 | **低** | 仅事实性提及，不挂 Tk 品牌 |
| 专利 | **低** | 文本无专利授权条款，但未见覆盖消息框的专利 |
| 出口管制 | **无** | 不含密码学；libtommath 是算术库 |
| GPL 不兼容 | **无** | TCL 与 GPL 兼容 |

结论：**wheel 形态在法律上零成本；exe 形态的代价是一份许可文件和一行
NOTICE。** 唯一摩擦是采购式的「OSI 白名单」政策，而控制台回落已经兜住了它。

---

## 6. 备选方案对比

| 方案 | 结论 |
| --- | --- |
| `ctypes` + `user32.MessageBoxW` | 仅 Windows，但零第三方字节、原生外观。`user32` 里没有 `InputBox`，仍需自绘对话框。适合作为 exe 的**第二后端**，不是替代品。 |
| COM `WScript.Shell.Popup` | Windows 上已由 `com-support` 提供；没有按钮组合、没有 `InputBox` 的默认值、不是独立对话框。保持现状。 |
| HTA / `InternetExplorer.Application` | 已被微软弃用，现代 Windows 上不可用；不作为目标。 |
| Tkinter（本方案） | Windows/macOS/Linux 都能用，pip 用户无新依赖，忠实覆盖两个内置函数。**选定。** |

---

## 7. 测试计划

- 单元测试保持控制台路径：它们直接调用 `vbs_run()`，命中 §3 第 2 条，另外在
  `tests/test_vbs_coverage.py` 里钉住 `OPENVBS_UI=console`，防止将来重构让 CI
  卡在弹窗上。
- 受 `OPENVBS_UI_TEST=1` 保护的 GUI 冒烟测试：创建隐藏 root、断言按钮/图标/
  默认按钮映射表，用 `after(0, …)` 之类的自动化关掉对话框；CI 永不执行。
- 手工矩阵：交互式控制台（弹窗）、`pythonw`（弹窗）、服务/Session 0（回落）、
  未装 `python3-tk` 的 Linux（回落）、`CI=1`（回落）。

---

## 8. 路线图

| 版本 | 范围 |
| --- | --- |
| 0.9.22 | `ui-support` 包、`MsgBox`/`InputBox` 对话框、启用规则、测试、文档 |
| 0.9.23 | 可选的 `ui` API 提供文件/文件夹选择器（明确标注为超集，不属于 VBScript） |
| 以后 | 若 exe 体积成为问题，再加 `ctypes` Windows 后端 |
