# 现代脚本语言特性对照

> **In English** — A checklist of modern scripting-language features
> (first-class functions, static members, modules, pattern matching,
> async, generics, tooling) mapped to ✅ implemented / 🚧 planned / ❌
> not applicable, so contributors can see where help is wanted.


OpenVBS++ 的定位是：**VBScript 的超集**（语法 100% 兼容），同时具备现代脚本语言该有的能力。
本文用一张表说明"现代特性 → 当前状态 → 落地方式"。

图例：✅ 已实现 · 🚧 规划中 · ❌ 不做（与 VBS 冲突或无意义）

## 1. 语言核心

| 现代特性 | 状态 | 说明 |
|---|---|---|
| 包/模块级别的函数与过程 | ✅ | `Sub` / `Function`，可先用后定义 |
| 一等函数（把过程当值） | 🚧 | 目前过程名只用于调用；v0.7 计划引入 `Function` 值类型 |
| 按值/按引用参数 | ✅ | `ByVal` / `ByRef`，含对象方法（统一 Ref 协议） |
| 默认参数 | 🚧 | `Optional` 尚未实现（VBS 有，属超集缺口） |
| 可变参数 | 🚧 | `ParamArray` 规划中 |
| 重载 | ✅ | **按参数个数**分派，编译期解析（v0.6.1） |
| 递归 / 闭包 | ✅ / ❌ | 递归支持；闭包不是 VBS 语义，不做 |
| 类与对象 | ✅ | `Class` / `Property` / `Me` / 构造析构 |
| 继承与覆盖 | ✅ | `Inherits`（超集扩展），方法覆盖走 MRO |
| 接口 / 多态 | 🚧 | 继承已提供多态；接口规划中 |
| 异常处理 | ✅ | `On Error Resume Next` / `GoTo 0` / `Err.Raise` |
| 结构化异常（try/catch） | 🚧 | `Try ... Catch ... End Try` 规划为超集扩展 |
| 编译期类型检查 | ✅ | `Option Typing` + `->` 注解，**编译期**报错 |
| 类型推断 | ✅ | 字面量、运算、内置函数返回、类成员、数组元素 |
| 字符串插值 $"...{x}..." | ✅ | 编译期展开成 & 拼接（v0.9.1） |
| 可选链 ?.b / 空合并  ?? b | ✅ | Nothing/Empty/Null 安全（v0.9.1） |
| 复合赋值 += -= *= /= &= | ✅ | 纯新增语法（v0.9.1） |
| 泛型 | 🚧 | 依赖类型系统扩展，后续规划 |
| 模式匹配 | ❌ | `Select Case` 已覆盖 VBS 语义 |

## 2. 工程能力

| 现代特性 | 状态 | 说明 |
|---|---|---|
| 模块化（多文件） | 🚧 | 目前单文件；`Include` 在规划中 |
| 包管理器 | 🚧 | 设计里叫 `over`，见 README 路线图 |
| 静态分析/类型检查 | ✅ | `Option Typing`，编译期完成 |
| 调试信息（行号、源码行） | ✅ | 语法/运行期错误都带 `Line N` + 源码 |
| AST 内省/反编译 | ✅ | `-v` 打印 token/AST/生成的 Python；`ast2str` |
| 代码格式化 | 🚧 | 规划中（基于 AST 的输出器已有雏形） |
| 语言服务器（LSP） | 🚧 | v1.0 目标 |

## 3. 互操作与扩展

| 现代特性 | 状态 | 说明 |
|---|---|---|
| 嵌入宿主语言 | ✅ | `#python ... #endpython`，行号保持不变 |
| 跨语言共享对象 | ✅ | `Option COP` 公共对象池，Python 侧同一个对象 |
| 动态执行 | ✅ | `Execute` / `Eval`（复用编译器，全局作用域） |
| 外部库调用 | 🚧 | 通过 `#python` 间接可用；原生 FFI 不做 |
| 标准库 | ✅ | `ovpprt`：字符串/数组/日期/格式化/集合/JSON/正则/文件系统 |
| 正则 | ✅ | `RegExp` 对象：`Test` / `Replace` / `Execute`（v0.9.2） |

## 4. 运行时质量

| 现代特性 | 状态 | 说明 |
|---|---|---|
| 不可变值/值语义 | ✅ | 数字/字符串/布尔；对象是引用 |
| 空值安全 | ✅ | `Empty` / `Null` / `Nothing` 三态，语义与 VBS 一致 |
| 垃圾回收 | ✅ | 交给 CPython；`Class_Terminate` 用 `__del__` |
| 迭代器协议 | ✅ | `For Each` 支持数组、`Dictionary`、类 |
| Unicode | ✅ | 全链路 UTF-8，支持中文标识符 |
| 可移植 | ✅ | 纯标准库，无平台绑定 |

## 5. 下一步（v0.7 建议）

1. `Optional` / `ParamArray`（VBS 有，属超集缺口，优先级最高）；
2. `Try ... Catch ... End Try`（超集扩展，和 `On Error` 并存）；
3. 过程作为值 / 简单闭包；
4. 多文件 `Include` + `over` 包管理器雏形；
5. `RegExp` 对象。
