# 面向对象特性清单 · Object Model Status

> **In English** — Status of every object-oriented feature: classes, fields,
> visibility, `Property Get/Let/Set`, `Me`, constructor/destructor, the
> `Inherits` extension, method overriding (virtual by default), arity-based
> overloading, `Shared` static members and `Implements` interface contracts.
> Keywords VBScript does not have (`Abstract`, `Static`, `Friend`, `Protected`,
> `Event`) are listed with the reason we deliberately skip them.


> 这张表回答一个问题：**VBScript 的对象特性有哪些，我们实现了多少，哪些是不适用、哪些是故意不做。**
>
> 图例：✅ 已实现并测试 · ⚠️ 已实现但有语义差异 · ❌ 不适用（VBS 没有） · 🚧 规划中（超集扩展）

## 1. 类与成员

| 特性 | 状态 | 实现方式 / 说明 |
|---|---|---|
| `Class … End Class` | ✅ | 编译成 `class _ovpp_cls_<name>`，带 `_vbs_name` |
| 字段 `Public x` | ✅ | Python 实例属性 |
| 字段 `Private n` / `Dim n` | ✅ | 改名为 `_priv_n`；外部读写都被拒绝（`__setattr__` 白名单） |
| `Property Get` | ✅ | `@property` |
| `Property Let` | ✅ | `@name.setter` |
| `Property Set` | ⚠️ | **与 `Let` 合并为同一个 setter**（VBS 区分"给对象赋值"与"给值赋值"，我们统一走 setter） |
| `Me` | ✅ | 编译成 `self`；只在类方法内有效 |
| `Class_Initialize` | ✅ | 在 `__init__` 里调用（无参数，与 VBS 一致） |
| `Class_Terminate` | ⚠️ | 映射到 `__del__`，**时机由 CPython GC 决定**，不保证立即执行 |
| 默认成员 `Attribute X.VB_UserMemId = 0` | ✅ | `x = obj`（Let）自动取默认属性；`Set x = obj` 保持引用 |

## 2. 创建与判断

| 特性 | 状态 | 说明 |
|---|---|---|
| `New Foo` | ✅ | 编译期解析成类构造 |
| `CreateObject("脚本类")` | ✅ | 运行时注册表 |
| `CreateObject("Scripting.Dictionary")` | ✅ | 内置 `Dictionary` |
| `CreateObject("Scripting.FileSystemObject")` | ✅ | v0.9 新增（文本文件 / 目录 / 路径） |
| `CreateObject("VBScript.RegExp")` / `New RegExp` | ✅ | v0.9.2 新增 |
| `CreateObject("WScript.Shell")` 等 COM 对象 | ✅ | v0.9.14 起 Windows 上由 `com-support` 包真连（需 pywin32；非 Windows 明确报错） |
| `GetObject(...)` | ⚠️ | 只对脚本类/内置类有效（不做对象持久化）；已运行的 COM 对象可 `GetObject("Excel.Application")` 连上 |
| `Is Nothing` / `IsObject` / `TypeName` | ✅ | 与 VBS 语义一致 |

## 3. 继承与多态（超集扩展）

| 特性 | 状态 | 说明 |
|---|---|---|
| `Class Child Inherits Parent` | ✅ | 直接映射成 Python 基类 |
| 方法覆盖 | ✅ | **默认就是虚方法**：子类同名方法覆盖基类，按 MRO 解析 |
| 字段/属性继承 | ✅ | 子类方法可直接读写基类 `Public` 字段；`_ovpp_attrs` 合并基类白名单 |
| 构造链 | ✅ | 子类 `__init__` 先调 `super().__init__()`，再初始化自己的字段，最后跑自己的 `Class_Initialize` |
| 多继承 | ❌ | 只允许单继承（VBS 没有继承，这是我们加的；单继承够用且语义清晰） |

## 4. 重载

| 特性 | 状态 | 说明 |
|---|---|---|
| 过程重载 | ✅ | **按参数个数**分派，编译期解析（零运行期开销） |
| 方法重载 | ✅ | 生成 `*args` 分发器，转发到 `_ovpp_m_<name>_<arity>` |
| 按类型重载 | ❌ | 需要运行期类型分派，会破坏"编译期检查"的目标；如需区分类型，用不同名字或 `When` 分支 |
| 重载 + `Optional`/`ParamArray` | ✅ | 参数个数区间参与匹配 |

## 5. 不适用（VBScript 本身没有）

这些关键字在 VBS 里不存在，因此我们**不加**——加了就等于发明新语法，需要单独设计：

| 关键字 | 出处 | 我们的态度 |
|---|---|---|
| `Abstract` / `MustOverride` | VB.NET | ❌ 不适用：VBS 无抽象类 |
| `Overridable` / `Overrides` | VB.NET / VBA | ❌ 不适用：我们的方法**默认可覆盖**，不需要修饰符 |
| `Static`（过程级持久变量） | VBA | ❌ 不适用：VBS 不支持；需要持久状态请用模块级变量 |
| `Shared`（类静态成员） | VB.NET | 🚧 规划：可用模块级变量 + `Public` 过程模拟 |
| `Friend` / `Protected` | VB.NET | ❌ 不适用：只有 `Public` / `Private` 两级 |
| `Implements`（接口） | VBA / VB.NET | 🚧 规划：继承已能提供多态，接口等真实需求出现再做 |
| `Event` / `RaiseEvent` | VBA / VB.NET | ❌ 不适用：VBS 无事件模型；需要回调时用过程值（`AddressOf`） |

## 6. 与 VBS 的语义差异汇总

1. **`Property Set` 与 `Let` 合并**：VBS 里 `Set obj.Prop = x` 走 `Property Set`，`obj.Prop = x` 走 `Property Let`。我们两者都走同一个 setter（Python 的 `@x.setter`）。
   影响：一个属性不能同时有 `Let` 和 `Set` 两套逻辑。实践里极少见。
2. **`Class_Terminate` 的时机**：依赖 Python 引用计数/GC，不保证在作用域结束时立刻调用。
3. **方法参数默认 `ByRef`**：与 VBS 一致，通过统一 `Ref` 协议实现（包括对象方法）。
4. **`Private` 成员在子类中不可见**：与 VBS 一致（VBS 的 `Private` 也是类内私有）。

## 7. 测试覆盖

| 测试文件 | 覆盖点 |
|---|---|
| `tests/test_objects.py` | 类、字段、方法、属性、`Me`、构造析构、可见性、默认成员、`GetObject`、方法 `ByRef` |
| `tests/test_try_catch.py` | 类方法里的结构化异常 |
| `tests/test_overloading.py` | 过程/方法重载、按个数分派、错误信息 |
| `tests/test_option_typing.py` | 类字段/属性/继承链上的类型检查 |
| `tests/test_vbs_superset.py` | 对象相关语法在"超集"矩阵里逐条验证 |

## 8. 未来扩展（按需）

1. `Shared` 静态成员（模块级变量 + 类名字典模拟）；
2. `Implements` 接口（多态已由继承满足，等真实需求）；
3. `Property Set` 与 `Let` 分离（如果遇到必须区分的兼容性问题）；
4. 构造参数（`New Foo(1, 2)` → 传给 `Class_Initialize` 的参数，VBS 不支持，属扩展）。
