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
3. **版权头统一为 `Copyright 2027 Marf Chen(MarfChen)`**，新文件也要带。
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
tests/            回归测试（examples + 过程 + 语法糖 + VBS 超集）
doc/              语言规范（docx）、VBS 超集矩阵、对象模型设计
```

## 关键实现约定（改代码前必读）

- **过程被改名**：`Sub Foo` → `_ovpp_sub_foo`，`Function Bar` → `_ovpp_fn_bar`。
  函数返回值用**函数原名的小写**做局部变量，末尾 `return`。
- **标识符大小写不敏感**：`canon(name) = name.lower()`，生成代码统一用小写。
- **ByRef**：实参包进 `ovpprt/ref.Ref`，过程内读写成 `参数.v`，调用后写回。
  表达式里的 ByRef 调用靠 `_pending_pre/_pending_post` + `_hoist()` 把语句提升到外层。
- **作用域**：过程里对模块级变量赋值会自动生成 `global` 声明。
- **`On Error Resume Next`**：编译期开关，之后的语句被包进 `try/except` 调 `_ovpp_err.capture`；
  每个过程 `_ovpp_err.enter()`/`leave()` 一帧。
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

## 环境坑

- **不要用 PowerShell 的 `Get-Content`/`Set-Content` 改源码**：它会把 UTF-8 重新编码成 GBK，
  中文注释会破坏（本项目踩过）。要么用编辑器/写文件工具，要么用 Python 读写 `encoding="utf-8"`。
- 仓库自带 `venv/`（Python 3.13）；`astpretty` 只用于 `-v` 打印 AST 结构。
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
  `Copyright 2027 Marf Chen(MarfChen)` + Apache 2.0 全文。已经有模块 docstring 的，
  把版权插在同一个 docstring 的开头，不要出现两个 docstring。
- **模块 docstring**：一律三单引号，第一行写模块职责，后面写实现要点。
- **函数/类 docstring**：只给"读完名字还需要解释"的写；`peek` / `eat` 这类自解释的不写；
  异常类由类名 + 错误号说明，不重复写 docstring。
- **版本号**：`OpenVBSPP/__init__.py`、`ast_nodes/__init__.py`、`ovpprt/__init__.py`、
  `plugins/__init__.py`、`pyproject.toml` 五处必须一致。
- **对象模型**：完整清单与"不适用关键字"的说明在 `doc/OOP.md`，改 OOP 前先读它。

## 环境坑（续）

- **Python 3.13 起 compile() 会校验位置区间**：AST 必须满足
  end_lineno >= lineno 且 end_col_offset >= col_offset，否则报
  ValueError: line 1, column 1-0 is not a valid range。
  compiler._normalize_positions() 会在 fix_missing_locations() 之后统一夹紧，
  新增 AST 生成代码时不要绕过 _to_code()。
