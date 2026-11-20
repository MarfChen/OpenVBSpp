# 总计划 · Roadmap

> **[English](ROADMAP.md)** | 中文

> **In English** — Where the project stands and what comes next. Everything the
> maintainer has explicitly **deferred** (remote package sources, the IDE/LSP
> work) is collected at the very end, so it is never mistaken for a near-term
> commitment. COM support used to be deferred; since v0.9.14 it ships as an
> optional `over` package (`com-support`), bundled on Windows.

---

## 1. 已完成

| 版本 | 内容 |
|---|---|
| 0.1–0.3 | 词法/语法/AST 映射、控制流、过程与 ByRef、`On Error`、`With`、`Select Case`、VBS 超集测试 |
| 0.4 | 对象模型（`Class`/`Property`/`Me`/`Dictionary`）、`Option Typing` 静态类型检查 |
| 0.5 | `Execute`/`Eval`、`Format*`、`Option Compare`、`GetObject`、默认成员、私有可见性 |
| 0.6 | 方法/过程重载、类继承 `Inherits`、`#python` 嵌入、`Option COP`、方法 ByRef |
| 0.7–0.8 | 数组元素类型、`Try/Catch/Finally`、标签与 `On Error GoTo`、标签式错误处理 |
| 0.9.0–0.9.5 | `Format()` 格式串、日期函数、`FileSystemObject`、插值与 `?.`/`??`/复合赋值、`AddressOf`、`RegExp`、`JSON`、多文件与模块隔离、`Resume` 状态机 |
| 0.9.6–0.9.7 | `Shared`、`Implements`、`Property Set`/`Let` 分离、构造参数 |
| 0.9.8–0.9.10 | `over` M1（本地源安装 + 锁文件 + 包名解析）、`CreateObject` 回归 VBS 语义 |
| 0.9.11–0.9.13 | 锁文件可重现（`--frozen`、`commit` 字段）、`over pack` / Git 源 / 本地 registry、`--json` 诊断、网页 IDE |
| 0.9.14 | `over` M5/M6：`python` 依赖交给 pip（锁里记真实闭包）、`platforms` 平台条件、bundled `com-support`（Windows 真 COM，脚本里无需 `import`） |
| 0.9.20 | Python 互操作改为 `python-embed` over 包 + 显式导入（`Option COP` 移除）、删除 `ast2str`、文档英文化并拆分 `dev-docs/` |
| 0.9.21 | VBScript 完整性补齐：`&H` / `&O` / `#date#` 字面量、剩下的 19 个内置函数（`Sin` … `GetRef`）、`ExecuteGlobal`、`TimeValue` / `Time()` 日期 bug，以及按实测校正的 `doc/VBS_SUPERSET.md` |
| 0.9.15 | 性能：启动懒加载、只读 `ByRef` 不再装箱（调用 13.7× → 1.35×）、无 `On Error` 的过程不建错误帧（递归 12× → 6.9×）、`For` 安全时生成 `range`（循环 1.65× → 0.99×） |

**当前状态**：VBScript 原生语句形式全齐（只有 `Declare` 是主动不做）；549 个测试全绿；进程内的循环/数组/函数调用已与手写 Python 同量级（0.99×–1.35×）。COM 在 Windows 上由
`com-support` 包真连（非 Windows、或 `OPENVBS_COM=0`、或没装 `pywin32` 时给明确报错），
见 §3 末尾的说明。

---

## 2. 下一步（建议顺序）

### P0 · 开源前必备（各约半天）

| # | 任务 | 为什么现在做 |
|---|---|---|
| 1 | `.github/workflows/ci.yml`（3.10–3.13 × win/linux，含 coverage） | ✅ v0.9.15 已完成 |
| 2 | issue / PR 模板（bug、feature、超集违规） | 降低外部贡献的沟通成本 |
| 3 | `over pack`（把包目录打成 `.ovppkg`） | 没有它，`over` 只能从目录装，没法分发 |
| 4 | ~~`openvbs --json check/run`~~ | ✅ 已完成（v0.9.12） |

### P1 · 语义收口（1–2 天）

| # | 任务 | 说明 |
|---|---|---|
| 5 | `Implements` 与 `Option Typing` 打通 | `Dim s -> IShape` 应接受任意实现类（现在报类型不匹配） |
| 6 | `Option Strict`（可选） | 禁止隐式转换，给类型系统一个"更严"的档位 |
| 7 | `over` M2：Git 源 + 冲突路径报告 + `over update` | 从"本地目录"升级到"团队共用" |

### P2 · 体验（按需）

| # | 任务 | 说明 |
|---|---|---|
| 8 | 编译缓存（按内容哈希） | 重复运行时省解析时间，也为 LSP 增量铺路 |
| 9 | REPL（`openvbs repl`） | 教学场景很吃这个 |
| 10 | 格式化器（基于 AST 输出器） | 社区贡献时能统一风格 |

---

## 3. 明确推迟（按维护者决定）

> 以下内容**不在近期计划内**，列在这里是为了避免被误认为承诺。

| 项目 | 推迟原因 | 前置条件 |
|---|---|---|
| **IDE / LSP** | 语言本身还有工程化任务没收尾 | P0 的 `--json` 先落地 |
| **远程包仓库 / 服务端** | 维护者没有服务器，`over` 坚持离线优先 | 社区出现真实分发需求 |
| **宣传 / 正式开源发布** | 维护者需要时间准备 | 由维护者决定时机 |

**已移出推迟清单**：

* **COM 支持** —— v0.9.14 以 `over` 包形态落地（`extras/com-support`，封 `pywin32`）。
  它没有破坏"零第三方依赖 + 跨平台"：`pywin32` 是**可选**的 `python` 依赖，只有真正
  用 COM 时才需要；非 Windows 平台不装、`CreateObject` 给明确报错。
* **`WScript.Shell` 纯 Python 替身** —— 不必再写：直接走真 COM 更准确。

---

## 4. 判断标准（每个新特性都要过）

1. **超集**：合法的 VBS 程序不能变；扩展只能新增语法，且要在
   `tests/test_vbs_superset.py` 里留用例；
2. **有测试**：新特性至少 3 条测试（正例、边界、错误）；
3. **有文档**：`doc/` 里对应章节 + README 能力表一行；
4. **可替代性说明**：如果一个特性有"不实现也能过"的理由，写进本文 §3。

---

## 5. 远期目标

| 目标 | 说明 | 启动条件 |
|---|---|---|
| **自举（self-hosting）** | 用一个"用 OpenVBS++ 自己写的 OpenVBS++ 编译器"来编译它自己，像 C 语言编译器那样递归式实现（T-diagram 的固定点） | 语言特性基本冻结 + 社区反响良好 + 维护者有余力 |

**现在不做，但先把路径写清楚**，免得变成一句空头承诺。真要启动时按三步走：

1. **确定 Bootstrap 子集**：从现有语言里划出"够写一个编译器"的最小集合，写进
   `doc/BOOTSTRAP.md`，并约定这个子集之后**只增不改**——这是自举能不能收敛的关键，
   子集越小越好。
2. **原型编译器（`.vbs` 写）**：只做词法 + 语法 + 生成 Python AST 三段，
   能编译 `examples/` 里一半的脚本就算成功；用现有的 Python 版 OpenVBS++ 去跑它，
   这是 T-diagram 里的 `T1`（用宿主语言实现的自举编译器）。
3. **固定点验证**：让 `.vbs` 版编译器编译它自己，产出 `stage1`；再用 `stage1` 编译一次
   得到 `stage2`；比对 `stage1` / `stage2` 的产物与行为一致（同时跑
   `tests/` 的子集，误差为零），才算自举成立。

**为什么放在远期**：自举会把"加一个语法特性"的成本从"改 transformer"变成"改三个实现"，
在语言还在演进期做自举是负收益。等 §2 的工程化清单收敛、`doc/VBS_SUPERSET.md` 的
覆盖矩阵稳定之后，再评估。
