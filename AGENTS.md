# AGENTS.md

给后续维护者 / AI 代理的项目约定。**动手前先读完这一页。**

## 项目是什么

`OpenVBS++`（Python 包 `OpenVBSPP`）是一个 VBScript 语法风格的**编译器**：把 `.vbs` 编译成 Python AST 后交给 CPython 执行。

```
lexer.py → parser.py → transformer.py → compiler.py → (runtime.py 执行)
```

## 铁律

1. **必须是 VBScript 的超集。**
   任何合法 VBS 程序都要能编译运行，除非它列在 `doc/VBS_SUPERSET.md` §4 的已知缺口里。
   扩展只能新增语法，不能改变已有 VBS 语义。改动后必须跑
   `python -m unittest tests.test_vbs_superset`，并在该文件里维护用例。
2. **错误类统一用 `OVPP*` 前缀**（`OVPPError` / `OVPPSyntaxError` / …），编号沿用 VBS 约定。
3. **版权头统一为 `Copyright 2026 Marf Chen(MarfChen)`**，新文件也要带。
4. **运行时不引入第三方依赖**（只允许标准库；`astpretty` 是可选的调试依赖）。
5. **错误必须带行号与源码行**：语法错误抛 `OVPPSyntaxError(..., line_number=, source_line=)`，
   运行期异常由 `runtime.wrap_runtime_error()` 翻译。

## 常用命令

```powershell
# 全量测试
.\venv\Scripts\python.exe -m unittest discover -s tests

# 只跑合法 VBS 一致性
.\venv\Scripts\python.exe -m unittest tests.test_vbs_superset -v

# 运行脚本
.\venv\Scripts\python.exe cli.py examples\control1.vbs

# 看生成出来的 Python（调试编译器时最有用）
.\venv\Scripts\python.exe cli.py examples\procedures.vbs -v

# 语法编译检查
.\venv\Scripts\python.exe -m compileall -q OpenVBSPP tests
```

## 目录

```
OpenVBSPP/        编译器与运行时（23 个模块）
  lexer.py        正则词法器：(kind, value, line, col)；注释、CRLF、行续接、`:` 都在这里
  parser.py       递归下降；块解析遇到 EOF 一律报错（历史上会死循环）
  transformer.py  两遍：先扫描登记过程签名/Option Explicit，再生成 Python AST
  compiler.py     vbs_compile()：exec / save 两种模式
  runtime.py      vbs_run()：编译 + 执行 + 异常翻译
  exceptions.py   OVPP* 错误类
  ast_nodes/      扁平 AST 节点
  ovpprt/         运行时库（boolean/string/array/empty/null/nothing/ref/err/builtins/format）
examples/         示例脚本（每个都会被 tests 跑）
tests/            回归测试（examples + 过程 + 语法糖 + VBS 超集 + 性能语义）
doc/              使用者文档，每篇英文 `.md` + 中文 `.zh-CN.md` 成对（语言参考、
                  对象模型、OOP、现代特性、over、网页 IDE、法务、首页）
dev-docs/         维护者文档（不发布到文档站，每篇也中英成对）：RELEASE、ROADMAP、IDE_LSP_DESIGN、
                  TRY_CATCH_DESIGN、OBJECT_MODEL_DESIGN、good-first-issues、plan.docx
extras/           可分发的示例包：com-support（COM 支持）、python-embed（Python 互操作）、
                  web-ide、vscode-openvbspp（VS Code 语法/片段/运行诊断扩展）
.github/          CI（ci.yml：Windows 3.10–3.13 + Linux 3.10/3.13，全是硬门禁，带 coverage）、
                  docs.yml（mkdocs → GitHub Pages）、issue/PR 模板
```

## 关键实现约定（改代码前必读）

- **过程被改名**：`Sub Foo` → `_ovpp_sub_foo`，`Function Bar` → `_ovpp_fn_bar`。
  函数返回值用**函数原名的小写**做局部变量，末尾 `return`。
- **标识符大小写不敏感**：`canon(name) = name.lower()`，生成代码统一用小写。
- **ByRef**：实参包进 `ovpprt/ref.Ref`，过程内读写成 `参数.v`，调用后写回。
  表达式里的 ByRef 调用靠 `_pending_pre/_pending_post` + `_hoist()` 把语句提升到外层。
  v0.9.15 起只给「真的会被写」的 ByRef 参数装箱（`_ref_param_names()`：过程体里被赋值 /
  `ReDim` / `For` / 作为实参继续传递才算）；只读参数按值传，被调方开头内联拆包
  `if type(n) is _ovpp_ref.Ref: n = n.v` —— 因为跨模块 `namespace` 与 `AddressOf` 委托
  仍可能把 `Ref` 传进来，这个拆包不能省。
- **作用域**：过程里对模块级变量赋值会自动生成 `global` 声明。
- **`On Error Resume Next`**：编译期开关，之后的语句被包进 `try/except` 调 `_ovpp_err.capture`。
  v0.9.15 起只有过程体里出现 `On Error ...` / `Resume` / `Err` 才生成
  `_ovpp_err.enter()`/`try/finally` + `leave()`（判定见 `_needs_error_frame()`），
  其余过程直接执行——`ovpprt/err.py` 模块加载时就有基帧，`capture()` 不会越界。
- **Python 互操作要显式导入**：`#python ... #endpython` 与 `COP` 公共对象池自 v0.9.20 起
  必须先 `<script src="python-embed" />`（随包分发的 over 包，声明 `capabilities: ["python-embed"]`）。
  `compiler._merge_source` 遇到带该 capability 的包只登记 feature、不合并源码；
  `_require_python_feature()` 在没有 feature 时报错。`Option COP` 已从语言中移除，
  写了就是语法错误（parser 的提示直接给出该导入语句）。
- **`With`**：对象表达式存进临时变量，`_with_stack` 保存节点，`.member` 编译成对它取属性。
- **类**：`Class Foo` → `class _ovpp_cls_foo`，带 `_vbs_name`；字段用 Python 属性，
  `Property Get` → `@property`，`Property Let/Set` → setter，`Me` → `self`，
  `Class_Initialize` → `__init__` 里调用，`Class_Terminate` → `__del__`。
  类内对字段的读写会被改写成 `self.<field>`（参数/局部变量优先，不改写）。
  无参方法读取（`d.Keys`）走 `_ovpp_obj.member_get`，**带参调用一律走
  `_ovpp_obj.call_method(obj, "name", [Ref...])`**：VBS 类的方法参数按 Ref 协议传
  （ByRef 用 `.v` 写回，ByVal 在方法体开头解包），内置对象自动取 `.v` 按值调用。
  私有成员改名 `_priv_*`，配合生成的 `__setattr__` 白名单拒绝外部写入。
- **继承**：`Class Child Inherits Parent` 直接映射成 Python 基类；`__init__` 先调
  `super().__init__()`，`_ovpp_attrs` 用基类列表 + 自己的字段拼出来。
- **`#python` 块**：`plugins/python_embed.py` 在词法前抽出代码块、原地留占位语句
  （行号保持不变），Transformer 把占位编译成 `exec(源码, globals(), globals())`。
- **`Option Typing`**：`OpenVBSPP/typecheck.py` 在编译前跑一遍，只在出现
  `Option Typing` 时生效；注解语法是 `->`（`:` 已用于语句分隔）。
- **`Option Strict`**：`+ - * / \ Mod ^` 与一元 `+/-` 要求数值操作数，文本操作数报
  错误 13 而不是转换；`&`、比较运算、内置函数不受影响，`Empty` 仍当 0、`Null` 仍传播。
  `Option Typing` 隐含开启它。编译期由 `self.option_strict` 决定走 `ovpp_add` 还是
  `ovpp_add_strict`（`ARITH_NAMES` 里 18 个名字成对登记，`_runtime_imports` 一次性按名导入）。
- **`Option` 必须写在所有其它语句之前**：`_collect` 先扫全模块 Option，迟到的
  `Option Strict Off` 会追溯改变上方代码的编译方式。`parser.statement()` 用
  `self.seen_code` 门禁，违反时报 `Syntax error (1002): 'Option X' must appear before
  any other statement`。
- **算术不再是 Python 运算符**：`op_map` 已清空，`arith_call_map` 把 `PLUS/MINUS/MUL/DIV/POW`
  映射到 `ovpp_add`/`ovpp_sub`/`ovpp_mul`/`ovpp_div`/`ovpp_pow`（`_transform_binop`），
  一元走 `ovpp_neg`/`ovpp_pos`（`_transform_unaryop`）。常量情形先经 `_fold_arith()`
  折叠成 `ast.Constant`，运行时函数开头有 int/float 内联快路径——不要改成
  `_plain_number(a) and _plain_number(b)`，那会每次运算多两次函数调用（实测更慢）。
- **内置函数的 Null 规则**：`_NULL_PROPAGATES`（返回 Null）、`_NULL_REJECTED`
  （抛 `OVPPInvalidUseOfNullError` = 94）、`_NULL_MISMATCH`（抛 13）三张名单在
  `builtins.py` 里用 `_propagate_null`/`_reject_null` 就地包装，**必须写在
  `REGISTRY` 构建之前**。名单依据 cscript 实测，改动要重新对拍。
- **未声明变量的读取**：无 `Option Explicit` 时读未声明的名字得到 `Empty`（VBS 行为）。
  编译期在模块顶部与过程开头补 `if "<name>" not in globals(): <name> = _ovpp_empty`
  （`_empty_prelude()`/`_empty_assign()`）；**必须用 `globals()` 判重**，因为编译器会在
  exec 前把 `<namespace src alias>` 的别名注入模块 globals。`Option Explicit` 下整块跳过。
- **`For` 快路径**：步长是常量、循环体不给计数器赋值时生成
  `for i in _ovpp_arr.for_seq(start, end, step)` + `else: if seq: i = i + step`
  （`_transform_for_seq()`）；其余情况仍生成 `while`。三条必须保住的语义：正常结束后
  计数器 = 末值 + 步长、`Exit For` 不加步长、空区间保留起点值。
- **启动性能**：`ovpprt` 里 `re`/`json`/`datetime`/`math`/`random` 一律延迟导入
  （`regexp.py`、`jsonobj.py`、`object.py`、`builtins.py` 的 `_LazyModule`）；
  `cli.py` 只在真正执行脚本时才 `import OpenVBSPP.runtime`。加新的顶层 import 前先想清楚
  它会不会被每个脚本付一次启动成本。
- **运行时别名**：`_ovpp_bool` `_ovpp_str` `_ovpp_arr` `_ovpp_builtin` `_ovpp_ref` `_ovpp_err`
  `_ovpp_obj` `_ovpp_collections` `_ovpp_empty` `_ovpp_null` `_ovpp_nothing`、
  类注册表 `_ovpp_classes`，用户变量不会撞名。

## 改完自检

1. `unittest discover -s tests` 全绿（含 expectedFailure 数量不变）。
2. 新增语法：更新 `README.md`、`doc/VBS_SUPERSET.md`、`tests/test_vbs_superset.py`。
3. 新增示例：同步 `tests/test_openvbspp.py` 的 `EXPECTED` 表（有覆盖性检查）。
4. 版本号：`OpenVBSPP/__init__.py`、`ast_nodes/__init__.py`、`ovpprt/__init__.py`、
   `plugins/__init__.py`、`pyproject.toml` 五处一致。
5. 改到 `vendor/com-support/` 时，同步 `extras/com-support/` 的 `com_support.py` 与
   `over.json`（两份必须逐字节一致，`tests/test_com.py` 会校验）。
6. 动编译器生成的代码形态（Ref 装箱、错误帧、`For` 快路径、懒加载）必须同时跑
   `tests/test_perf_semantics.py`；它专门钉这些优化的语义边界，改性能前后都要绿。
7. 性能数字变了就同步 `README.md` 的「## 性能」小节和 `tools/bench.py` 的说明文字——
   那张表是项目的对外承诺，不能和实测漂移。
8. **README 是双语的且英文为主**：`README.md`（英文主文档）与 `README.zh-CN.md`（中文全本）
   是两份文档，改功能、用例数、徽章、性能表时**两边都要改**；数字不一致就是假绿。
   原来的 `README.en.md` 已删除，不要再引用它。
9. 改 `extras/vscode-openvbspp/` 后至少要自检：4 个 JSON 都能 `python -m json.tool` 解析、
   `node --check extension.js` 通过、语法里的 `include` 都能在 `repository` 找到；
   新增的语法特性要同步 `syntaxes/openvbspp.tmLanguage.json` 与 README 的特性表。
10. 改 `.github/workflows/ci.yml`（或 `docs.yml`、`mkdocs.yml`）后用 PyYAML `safe_load`
    校验一遍。Windows 与 Linux 现在都是硬门禁（Linux 的 `experimental` 维已在 v0.9.15
    移除）；某个版本临时不稳定时再给它单独加回 `continue-on-error: true`。
11. 动 `cli.py` 的取参逻辑（`-c` / `--code` / stdin / 与脚本文件互斥）必须跑
    `tests/test_cli_inline.py`；新增子命令也照这个模式补用例。
12. 改 `doc/*.md` 或 `mkdocs.yml` 后跑一次 `mkdocs build --site-dir site`（需要
    `pip install "mkdocs-material>=9.5"` 或 `pip install -e ".[docs]"`），新文档要进 `nav`，
    构建出的 `site/` 不要提交（已在 `.gitignore`）。
13. 用例数统计口径：README 徽章、README 的 `## Tests` 行、`dev-docs/RELEASE.md` §9.4
    三处必须一致（当前 **644**），改测试后一起改。
14. **`doc/` 的每一篇都是中英成对**：改 `doc/X.md` 就同步 `doc/X.zh-CN.md`（反之亦然），
    两版首行都要保留 `> English | **[中文](X.zh-CN.md)**` 这类切换行。维护者内容
    （发版、路线、设计、新手任务）只放 `dev-docs/`，不要塞进 `doc/`（它会进文档站）。
    `dev-docs/` 同样逐篇中英成对（`X.md` 英文主文 + `X.zh-CN.md` 中文），新增维护者
    文档必须一次交两版；唯一例外是 `plan.docx`（二进制，无法双语）。
15. **新增 vendor over 包要三处同步**：`OpenVBSPP/vendor/<pkg>/`（随包分发）、
    `extras/<pkg>/`（项目可 `over add <pkg>`，与 vendor 逐字节一致，`tests/test_python_embed.py`、
    `tests/test_com.py` 与 `tests/test_ui_dialogs.py` 会校验），以及打包清单 `MANIFEST.in`、
    `pyproject.toml` 的 `package-data`、`packaging/openvbs.spec` 的 `datas`。
16. **第三方代码/许可**：新增会随产物分发的第三方代码（如 Tcl/Tk）时，必须更新
    `THIRD_PARTY_LICENSES.md` 与 `NOTICE`，并把许可文件加进 `packaging/openvbs.spec`
    的 `datas`。许可原文逐字保留英文，不得改写或翻译。

## 环境坑

- **不要用 PowerShell 的 `Get-Content`/`Set-Content` 改源码**：它会把 UTF-8 重新编码成 GBK，
  中文注释会破坏（本项目踩过）。要么用编辑器/写文件工具，要么用 Python 读写 `encoding="utf-8"`。
- 仓库自带 `venv/`（Python 3.13）。`-v` 与 `--save` 打印/写出的 Python 用标准库
  `ast.unparse`，不再依赖 `astpretty`（`plugins/debug_tools.py` 已在 v0.9.20 删除）。
- `tests/` 没有 `__init__.py`，`unittest discover -s tests -t .` 会报
  `Start directory is not importable`；用 `python -m unittest discover -s tests`
  （不带 `-t .`），或直接 `python -m unittest tests.test_over` 跑单个模块。
- **COM 相关**：`pywin32` 只能在 `com_support.py` 里延迟 import，`ovpprt/com.py` 与
  `object.py` 的顶层永远不要 import win32com——否则非 Windows 平台或没装 pywin32 的环境
  连 `import OpenVBSPP` 都会失败。测试里要把 `OPENVBS_COM` 设为 `0`，
  否则开发机上真装了 pywin32 时会去连本机的 Excel/Office，测试变慢还依赖环境。
- **本机（H: 盘的"私人空间"卷）不能执行 setuptools 生成的 console-script 存根**：
  `venv\Scripts\pip.exe`、`openvbs.exe` 这类小 exe 直接跑会报
  “Access is denied”，而 `python.exe` 与 `.cmd` 正常。所以验证命令行要走
  `venv\Scripts\python.exe -m OpenVBSPP ...`，或用仓库里备好的
  `venv\Scripts\openvbs.cmd`（内容就是 `"%~dp0python.exe" -m OpenVBSPP %*`）。
  这是机器策略问题，不是包本身的问题。


## 文档约定

- **版权头**：所有 `.py` 文件（含 `tests/`、`cli.py`、`packaging/`）开头统一是
  `Copyright 2026 Marf Chen(MarfChen)` + Apache 2.0 全文。已经有模块 docstring 的，
  把版权插在同一个 docstring 的开头，不要出现两个 docstring。
- **注释与 docstring 一律中英双语、面向使用者**（v0.9.20 起）：先英文、再
  `中文：…`。模块 docstring 在版权块之后先写英文段落再写 `中文：…` 段落；类/函数
  docstring 写成 `"""English` + 空行 + `中文：…` + 收尾 `"""`（英文首行绝不能
  提前闭合 `"""`，否则中文变成裸代码，报 `SyntaxError: invalid character '：'`）；
  行内注释写 `# English` 一行 + `# 中文：…` 一行，短的可以并成 `# English / 中文：…`。
  内容标准是"使用者视角"：这个模块/函数给脚本作者什么能力、怎么用、有什么限制；
  面向开发者的实现备忘（构建细节、内部权衡、待办计划）写在 `dev-docs/`，不要留在源码里。
- **模块 docstring**：一律三单引号，版权块之后先英文段、再 `中文：` 段。
- **函数/类 docstring**：只给"读完名字还需要解释"的写；`peek` / `eat` 这类自解释的不写；
  异常类由类名 + 错误号说明，不重复写 docstring。
- **版本号**：`OpenVBSPP/__init__.py`、`ast_nodes/__init__.py`、`ovpprt/__init__.py`、
  `plugins/__init__.py`、`pyproject.toml` 五处必须一致。
- **破坏性改动的写法**：移除语法（如 `Option COP`）要在 parser/transformer 里给出
  「已移除 + 怎么迁移」的明确报错，不能静默忽略；同时写进 `CHANGELOG.md` 的 Breaking
  小节，并在两版 README 里留迁移说明。
- **对象模型**：使用者的模型说明在 `doc/OBJECT_MODEL.md`（中文 `OBJECT_MODEL.zh-CN.md`），
  语言层面的 OOP 用法在 `doc/OOP.md`，设计/风险备忘在 `dev-docs/OBJECT_MODEL_DESIGN.md`。
  改 OOP 前先读这三份。
- **法务口径**：与微软 VBScript 的关系（独立实现、商标指名性使用、不复制其文档）
  写在 `doc/LEGAL.md`；新增文档/示例别整段照抄微软帮助文本，演示代码自己写。
- **双语与入口**：仓库对外的第一份文档是英文 `README.md`；`README.zh-CN.md` 是完整中文版，
  任何功能/数字变更都要同时落地。徽章含义与"怎么变绿"写在 `dev-docs/RELEASE.md` §8，
  CI 矩阵与提级方式写在 §7。
- **编辑器支持**：语法高亮/片段/运行诊断在 `extras/vscode-openvbspp`；LSP 尚未实现，
  设计与里程碑在 `dev-docs/IDE_LSP_DESIGN.md`（§0 是已交付现状，别把"计划"写成"已支持"）。

## 环境坑（续）

- **Python 3.13 起 compile() 会校验位置区间**：AST 必须满足
  end_lineno >= lineno 且 end_col_offset >= col_offset，否则报
  ValueError: line 1, column 1-0 is not a valid range。
  compiler._normalize_positions() 会在 fix_missing_locations() 之后统一夹紧，
  新增 AST 生成代码时不要绕过 _to_code()。
- **`MANIFEST.in` 的路径不能带空格**：distutils 会按空白切分，`include doc/OpenVBS spec.docx`
  被读成 `doc/OpenVBS` 和 `spec.docx` 两个不存在的 pattern（build 时只给 warning，产物静默少文件）。
  所以 `doc/` 一律写成 `recursive-include doc *.md *.docx` + 需要挡掉的用无空格路径
  `exclude dev-docs/plan.docx`。改完必须验证 sdist 内容，不能只看 "Successfully built"：
  ```powershell
  .\venv\Scripts\python.exe -c "import tarfile;print([n for n in tarfile.open('dist/openvbspp-<ver>.tar.gz').getnames() if '/doc/' in n])"
  ```
- **`python -m build` 的隔离环境可能装不上 setuptools**（本机离线/代理原因）：
  用 `python -m build --no-isolation --outdir dist`，之后照常 `twine check dist/*`。
