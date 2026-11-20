# Contributing to OpenVBS++ / 贡献指南

> **English** first, 中文在后。 / English first, Chinese follows.

---

## English

### What this project needs

OpenVBS++ is a one-person project that wants to become a community project.
The most valuable contributions, in order:

1. **Edge cases that break the VBS superset rule** — any legal VBScript that
   fails to compile or behaves differently is a bug worth a priority issue.
2. **Documentation in your language** — translation is as valuable as code.
3. **New built-in functions** and runtime-library improvements.
4. **Performance** (compile time, generated-code speed).
5. **Packaging** (Windows/macOS/Linux, other Python versions).

### The one rule you must not break

**OpenVBS++ must stay a superset of VBScript.** Extensions may add syntax, but
must never change the meaning of existing VBS code. If you add a feature, run:

```bash
python -m unittest tests.test_vbs_superset
```

and add a case for it in `tests/test_vbs_superset.py`.

### Setup

```bash
git clone https://github.com/MarfChen/OpenVBSpp
cd OpenVBSpp
python -m venv venv
# Windows: venv\Scripts\activate
# Linux/macOS: source venv/bin/activate
pip install -e ".[all]"
```

### Before you open a PR

```bash
python -m unittest discover -s tests     # must be green, no expectedFailure changes
python -m compileall -q OpenVBSPP tests  # must be clean
```

Checklist:

- [ ] Tests pass locally.
- [ ] New syntax → `README.md` + `doc/VBS_SUPERSET.md` + a superset test.
- [ ] New example → update `EXPECTED` in `tests/test_openvbspp.py`.
- [ ] Version strings (4 places) unchanged unless you are releasing.
- [ ] Copyright header `Copyright 2027 Marf Chen(MarfChen)` on new files.
- [ ] Commit signed off (see DCO below).

### Developer Certificate of Origin (DCO)

We use **DCO**, not a CLA. It means: by adding `Signed-off-by` you certify you
have the right to submit the code under Apache-2.0. No paperwork, no copyright
assignment — you keep your copyright.

```bash
git commit -s -m "feat: support Optional parameters"
```

which produces:

```
feat: support Optional parameters

Signed-off-by: Your Name <you@example.com>
```

### Code style

- No third-party runtime dependencies (standard library only).
- Errors are `OVPP*` classes with a line number and the offending source line.
- Comments in the code may be Chinese; keep public docs bilingual.
- Prefer small, focused PRs — one feature or one bug per PR.

### Reporting a superset violation

Please include:

1. The smallest `.vbs` that reproduces it.
2. What real VBScript does (link to docs if possible).
3. What OpenVBS++ does instead.

---

## 中文

### 这个项目需要什么

OpenVBS++ 现在是一个人维护的项目，目标是变成社区项目。最有价值的贡献依次是：

1. **违反"VBS 超集"规则的边界情况** —— 任何合法 VBS 跑不通或行为不同，都是高优先级 bug；
2. **你母语的文档/翻译** —— 翻译和写代码一样有价值；
3. **新内置函数**与运行时库改进；
4. **性能**（编译时间、生成代码速度）；
5. **打包**（更多平台、更多 Python 版本）。

### 唯一不能破的规则

**OpenVBS++ 必须始终是 VBScript 的超集。** 扩展只能新增语法，不能改变已有 VBS 语义。
新增特性后必须跑：

```bash
python -m unittest tests.test_vbs_superset
```

并在 `tests/test_vbs_superset.py` 里补一条用例。

### 本地环境

```bash
git clone https://github.com/MarfChen/OpenVBSpp
cd OpenVBSpp
python -m venv venv
# Windows: venv\Scripts\activate
# Linux/macOS: source venv/bin/activate
pip install -e ".[all]"
```

### 提 PR 之前

```bash
python -m unittest discover -s tests     # 必须全绿，expectedFailure 数量不变
python -m compileall -q OpenVBSPP tests  # 必须无错
```

自查清单：

- [ ] 本地测试通过。
- [ ] 新语法 → 同步 `README.md` + `doc/VBS_SUPERSET.md` + 超集用例。
- [ ] 新示例 → 同步 `tests/test_openvbspp.py` 的 `EXPECTED`。
- [ ] 非发版 PR 不改版本号（四处保持一致）。
- [ ] 新文件带 `Copyright 2027 Marf Chen(MarfChen)` 版权头。
- [ ] 提交带 `Signed-off-by`（见下面的 DCO）。

### DCO（开发者原创声明）

我们采用 **DCO**，不用 CLA。含义：你加一行 `Signed-off-by` 即表示你有权以
Apache-2.0 提交这段代码。**不需要签合同，也不需要转让版权**——版权始终是你的。

```bash
git commit -s -m "feat: 支持 Optional 参数"
```

### 代码风格

- 运行时不引入第三方依赖（只用标准库）；
- 错误统一用 `OVPP*` 类，必须带行号和出错的源码行；
- 代码注释可以写中文，公开文档保持双语；
- 一个 PR 只做一件事（一个特性或一个 bug）。

### 报告"超集违规"

请附上：

1. 能复现的最小 `.vbs`；
2. 真实 VBScript 的行为（最好附官方文档链接）；
3. OpenVBS++ 现在的行为。

---

## Where to start / 从哪里开始

见 [`dev-docs/good-first-issues.md`](dev-docs/good-first-issues.md)。
