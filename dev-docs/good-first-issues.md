# Good First Issues / 新手任务

> **English** first, 中文在后。每个任务都写清了"改哪个文件、跑哪条测试"。

These are scoped, self-contained tasks. Pick one, open an issue saying you are
on it, and send a PR. Ask questions early — there are no silly questions.

下面这些任务都有明确范围。挑一个，在 issue 里说一声"我来做"，然后提 PR。
有疑问随时问，没有"蠢问题"。

---

## GFI-1 · `RegExp` object / 正则对象

**EN** — VBScript ships a `RegExp` object (`New RegExp`, `Pattern`, `Global`,
`IgnoreCase`, `Test`, `Replace`, `Execute`). We only have `Filter`, which uses a
regex internally. Implement `RegExp` as a runtime class.

**中文** —— VBScript 自带 `RegExp` 对象（`Pattern` / `Global` / `IgnoreCase` /
`Test` / `Replace` / `Execute`），我们目前只有内部用正则的 `Filter`。

| 项目 | 说明 |
|---|---|
| 改哪里 | 新建 `OpenVBSPP/ovpprt/regexp.py`，在 `transformer.py` 的 `_transform_create_object` 里加 ProgID `vbscript.regexp` |
| 参考 | `ovpprt/collections.py` 的 `Dictionary`（同样是内置对象） |
| 测试 | `tests/test_regexp.py`（新建），并在 `test_vbs_superset.py` 加两条用例 |
| 难度 | 中 |

---

## GFI-2 · 多文件 `Include` / 多文件包含

**EN** — Add `Include "path.vbs"` that inlines another script at compile time
(lex-level splice, before parsing, so line numbers of the main file stay sane).
Decide and document: path resolution, cycles, and interaction with `Option`
statements.

**中文** —— 增加 `Include "path.vbs"`，在编译期把另一个脚本内联进来
（放在词法之前做拼接，主文件行号不能乱）。需要定好：路径解析、循环包含检测、
与 `Option` 语句的交互。

| 项目 | 说明 |
|---|---|
| 改哪里 | 新建 `OpenVBSPP/plugins/include.py`（可参照 `plugins/python_embed.py`），在 `compiler.py` 里接入 |
| 坑 | `plugins/python_embed.py` 用"占位 + 补换行"保持行号，这个思路可以复用 |
| 测试 | `tests/test_include.py` |
| 难度 | 中 |

---

## GFI-3 · `Optional` 参数的 `IsMissing` / 判断是否省略

**EN** — `Optional` is implemented (defaults work), but a function cannot tell
whether an argument was supplied. Add `IsMissing(x)` and document that it only
works for `Optional` parameters.

**中文** —— `Optional` 已支持默认值，但函数内部无法判断"调用时到底传没传"。
补一个 `IsMissing(x)`，并说明只对 `Optional` 参数有效。

| 项目 | 说明 |
|---|---|
| 改哪里 | `OpenVBSPP/ovpprt/builtins.py`（加函数 + 注册名），`transformer.py` 里 `Optional` 参数的默认值从 `EMPTY` 改成专用哨兵 |
| 测试 | `tests/test_optional_paramarray.py` 加 2 条 |
| 难度 | 小 |

---

## GFI-4 · 错误信息国际化 / Better error messages

**EN** — Runtime errors currently read like `Type mismatch (13): ...`. Add a
hint line for the five most common mistakes (calling a `Sub` as a function,
missing `Set` for objects, array index out of range, wrong argument count,
assigning to a constant).

**中文** —— 现在的错误信息偏"机器味"。给最常见的五类错误加一行"人话提示"：
把 `Sub` 当函数用、对象忘了 `Set`、数组下标越界、参数个数不对、给常量赋值。

| 项目 | 说明 |
|---|---|
| 改哪里 | `OpenVBSPP/runtime.py` 的 `_ERROR_MAP` / `wrap_runtime_error` |
| 要求 | 提示必须双语或纯英文（错误信息目前是英文） |
| 测试 | `tests/test_errors.py`（新建），断言 `"hint"` 出现在消息里 |
| 难度 | 小 |

---

## GFI-5 · 完善 `doc/VBS_SUPERSET.md` 的语义偏差表

**EN** — `doc/VBS_SUPERSET.md` §3 lists our deviations from VBS. Several are
described in one line but deserve a minimal repro + explanation. Add a repro
snippet for each row and a short "why we chose this" note.

**中文** —— §3 列了我们和 VBS 的语义偏差，但每条只有一句话。给每条补一个最小
复现片段 + 一句"为什么这样选"。

| 项目 | 说明 |
|---|---|
| 改哪里 | `doc/VBS_SUPERSET.md` §3 |
| 测试 | 无（纯文档），但每条复现片段要能真跑通 |
| 难度 | 小（非常适合第一次贡献） |

---

## GFI-6 · 性能：编译期缓存 / Compile cache

**EN** — `vbs_compile` re-parses on every run. Add an in-memory LRU cache keyed
by source hash, with an opt-out env var `OVPP_NO_CACHE=1`.

**中文** —— 现在每次运行都重新解析。加一个按源码哈希做 key 的内存 LRU 缓存，
并支持 `OVPP_NO_CACHE=1` 关闭。

| 项目 | 说明 |
|---|---|
| 改哪里 | `OpenVBSPP/compiler.py` |
| 注意 | 不能改变报错行为；`filename` 要参与 key |
| 测试 | `tests/test_compiler_cache.py` |
| 难度 | 中 |

---

## 怎么认领 / How to claim

1. 在本仓库开一个 issue，标题写 `[GFI-x] 我要做这个`；
2. fork → 新建分支 `feat/gfi-x-short-name`；
3. 按 `CONTRIBUTING.md` 的自查清单提 PR；
4. 提交带 `Signed-off-by`（DCO）。
