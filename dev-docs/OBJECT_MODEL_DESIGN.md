# OpenVBS++ 对象模型设计报告

> 版本：v0.4（P1 + P2 已落地）
> 目标读者：项目作者 / 后续维护者
> 状态：**P1（集合 / `Set` / `CreateObject`）与 P2（`Class` / `Property` / `Me`）已实现**，
> 实现与本文的差异见 §8「实现现状对照」。P3（继承等超集扩展）仍为规划。

---

## 1. 目标与硬约束

### 1.1 必须满足的前提

| 约束 | 说明 |
|---|---|
| **VBS 超集** | 任何合法 VBScript 程序（含 `Class` / `Property` / `Set` / `CreateObject`）都必须能编译。对象相关语法是目前**唯一还缺的**一大块（见 `doc/VBS_SUPERSET.md`）。 |
| **不破坏现有语义** | 现有实现：过程改名为 `_ovpp_fn_*` / `_ovpp_sub_*`；ByRef 用 `Ref` 盒子 + 调用点写回；`Empty` / `Null` / `Nothing` 是运行时哨兵；`Option Explicit` 在编译期查名字；`On Error Resume Next` 靠"每条语句包 try/except + `_ovpp_err.capture`"实现。对象模型必须与这些机制**正交**，不能推倒重来。 |
| **零第三方依赖** | 运行时只能用标准库（这是教育向项目的底线）。 |
| **可移植** | 不启动 Windows COM。VBS 的对象在真实环境里是 IDispatch，我们不追这条路线，而是提供**等价的 Python 对象 + 同名内置类**。 |
| **错误信息友好** | 行号 + 源码行 + VBS 错误号，与现有 `OVPP*` 错误体系一致。 |

### 1.2 设计取向

- **语义优先**：优先保证"同一个 VBS 程序在我们的实现里行为一致"，其次才考虑性能与实现优雅。
- **编译期能判定的就别留到运行期**：`Set`/`Let` 的区分、默认成员的解析、`Option Explicit` 与字段可见性，都尽量在编译期确定。
- **不做 COM 模拟**：`CreateObject("Scripting.Dictionary")` 走内置注册表；`CreateObject("Excel.Application")` 之类的 COM ProgID 直接给出清晰错误。

---

## 2. VBScript 对象语义速览

| 概念 | VBS 语义 | 关键陷阱 |
|---|---|---|
| 值 vs 对象 | 变量可以装标量（Integer/String/…）或对象引用 | `obj = other` 在 VBS 里是**默认属性赋值**，不是引用赋值 |
| `Let` | 显式值赋值（通常省略） | `Let obj = x` 写的是默认属性 |
| `Set` | 引用赋值，只用于对象 | `Set x = 5` 是运行时错误 |
| 默认成员 | 类可声明 `Attribute Value` 默认成员，`obj = v` / `v = obj` 等价于访问它 | 没有默认成员时 `obj = v` 会报"对象不支持此属性" |
| `Property Get/Let/Set` | 属性过程，可带参数 | `Set` 属性过程名叫 `Property Set`，与 `Set` 语句同名不同物 |
| `Me` | 当前对象引用 | 在没有类上下文时不可用 |
| `Class_Initialize` / `Class_Terminate` | 构造 / 析构 | 析构时机不确定，且不能抛错 |
| `Nothing` | 空对象引用 | 用 `Is` 比较，不能用 `=` |
| `Is` | 对象身份比较 | 与 `Case Is` 语法同名但无关 |
| `For Each` | 遍历集合/数组，迭代器语义 | 遍历中修改集合是未定义行为 |
| `With` | 表达式求值一次，块内 `.member` 指向它 | 块内不能重新绑定 |
| `CreateObject` | 按 ProgID 创建 COM 对象 | 我们的实现只认内置类名 |
| 生命周期 | 引用计数（COM），最后引用释放时触发 `Class_Terminate` | 循环引用会泄漏 |

---

## 3. 落到 Python 上的设计

### 3.1 总体表示

**每个 VBS `Class` 编译成一个 Python class**，命名规则与过程一致：

```
VBS:  Class Person
Python: class _ovpp_cls_person: ...
```

类体里的字段用**实例字典**保存（`self._fields`），而不是直接映射成 Python 属性。理由：

1. VBS 字段是 `Variant`，可以是任意类型，用 dict 便于统一处理 `Empty` 初值；
2. `Property` 与字段可能同名，dict 存储能明确区分"字段槽"和"属性过程"；
3. `With` / 默认成员 / `For Each` 都需要统一的成员查找入口。

```python
class _ovpp_cls_person:
    _field_names = ("name", "age")          # 编译期已知，用于初始化与 Option Explicit 检查
    _default_member = None                  # 默认成员名（Attribute Value 指定）

    def __init__(self):
        self._fields = {"name": _ovpp_empty, "age": _ovpp_empty}
        self._init()                        # 调用 Class_Initialize

    def _init(self):
        # Class_Initialize 的编译结果（如有）
        ...

    # 属性过程
    def _get_name(self):
        return self._fields["name"]
    def _let_name(self, value):
        self._fields["name"] = value

    # 方法（VBS 的 Sub/Function）
    def _ovpp_sub_say(self, msg):
        ...
```

### 3.2 成员查找顺序（编译期生成的统一入口）

所有成员访问（`obj.x`、`With .x`、`Me.x`）都编译成对**同一个**访问器的调用，由它在运行时决定语义：

```python
_ovpp_obj.get_member(obj, "x")   # 读
_ovpp_obj.set_member(obj, "x", v)  # 写（Let）
_ovpp_obj.set_object(obj, "x", v)  # 写（Set）
```

查找顺序（与 VBS 一致）：

| 顺序 | 目标 | 命中后的行为 |
|---|---|---|
| 1 | `Property Get x` | 调 `_get_x()` |
| 2 | 字段 `x` | 返回 `_fields["x"]` |
| 3 | 方法 `x` | 返回绑定方法（供 `obj.x(...)` 调用） |
| 4 | 内置成员（`Dictionary` 等） | 走内置类的 Python 实现 |
| 5 | 都没有 | `OVPPObjectNotSupportedError`（438） |

写入时：先找 `Property Let/Set`，再找字段，都没有就报错（VBS 同样不允许动态加字段）。

### 3.3 `Let` 与 `Set` 的区分

- 语法层：`Set x = expr` 保留 `Set` 标记，`Let x = expr` / `x = expr` 是普通赋值。
- 编译期：
  - 目标变量已知是对象类型（静态推断：被 `Set` 赋值过、或 `Dim` 后首次 `Set`）→ 按引用赋值；
  - 目标是类字段 → 生成 `set_member`；
  - **无法静态判定时**（VBS 大量存在，例如 `x = obj` 想写默认属性）→ 生成运行时分派：

```python
_ovpp_obj.let_assign(current_value, new_value)
# 若 current_value 是对象且有默认成员 → set_member(default, new_value)
# 否则 → 普通赋值
```

这条"运行时分派"是 `Let`/`Set` 语义正确性的关键，代价是一次类型判断（可接受）。

### 3.4 `Property Get/Let/Set`

| VBS | Python |
|---|---|
| `Property Get X()` | `def _get_x(self):` |
| `Property Let X(v)` | `def _let_x(self, v):` |
| `Property Set X(v)` | `def _setobj_x(self, v):` |
| 带参数的 `Property Get X(i)` | `def _get_x(self, i):`（参数照常传递） |

- `Let` 与 `Set` 属性过程同名（`X`）时，两份实现共存，由调用方是 `let` 还是 `set` 决定用哪个——这正是 VBS 的行为。
- 属性过程体里的 `X = value`（给属性名赋值）在 VBS 中表示"设置返回值"（Get 时）或递归——按 VBS 规则：**Get 过程内 `X = v` 是设置返回值**，编译成 `ret_x = v` 最后 `return ret_x`（复用函数返回值的做法）。

### 3.5 `Me`

`Me` 编译为方法内的第一个隐式参数（Python 的 `self`）。在 `With` 块内出现 `Me` 时正常解析；在没有类的上下文中使用 `Me` → 编译期报错（与 VBS 一致）。

### 3.6 默认成员

VBS 的 `Attribute Value` 声明默认成员。编译期记录 `_default_member`，运行时的 `let_assign` / 取值分流都查它。

没有默认成员的类被当作值使用时 → 运行期 438 错误（`Object doesn't support this property or method`）。

### 3.7 `CreateObject` / `GetObject`

```vb
Set d = CreateObject("Scripting.Dictionary")
```

编译策略：`CreateObject` 是内置函数，参数是**字符串字面量**时在编译期解析成类，参数是变量时走运行期注册表。

| ProgID | 我们的实现 |
|---|---|
| `Scripting.Dictionary` | `ovpprt.collections.Dictionary` |
| `Scripting.FileSystemObject` | 暂不实现（涉及文件系统） |
| `WScript.Shell` 等宿主对象 | 不实现，给出明确错误 |
| 其他（`Excel.Application` 等） | `OVPPObjectNotSupportedError`，消息里提示"COM 对象不受支持" |

`GetObject` 同理：只支持我们能提供的对象，否则报错。

### 3.8 内置集合

| VBS | Python 实现 |
|---|---|
| `Dictionary`：`Add` / `Exists` / `Item` / `Remove` / `RemoveAll` / `Count` / `Keys` / `Items` | dict 包装，保持插入顺序（Python 3.7+ 天然有序） |
| `Collection`：`Add` / `Item` / `Remove` / `Count` | list 包装，索引 1-based（VBS 的 Collection 是 1 起的！） |
| 数组 | 已有的 list |

`For Each` 对 `Dictionary` 迭代 **Keys**（VBS 行为），对 `Collection` 迭代元素，对数组迭代元素 —— 正好都是 Python 的迭代协议，`ast.For` 直接可用。

### 3.9 与 `With` / `For Each` / `Err` 的交互

- `With obj`：`obj` 求值一次存进临时变量（现有实现已如此），块内 `.x` 走同一套 `get_member` / `set_member`。
- `For Each x In obj`：`iter(obj)`；对象类可以实现 `_iterate()` 返回键或元素列表。
- `Err`：`Err` 是内置单例对象，不是脚本类；`Err.Number` 等走 Python `@property`（当前实现）。
- 对象方法里抛出的错误：由 `_ovpp_err` 的 try/except 捕获，`Err.Number` 填对应编号；方法内部若有 `On Error Resume Next`，只影响方法自己那一帧（现有 `enter`/`leave` 机制）。

### 3.10 生命周期与 `Class_Terminate`

| 方案 | 优点 | 缺点 |
|---|---|---|
| `__del__` | 简单 | 时机不确定、异常被吞、循环引用不触发 |
| `weakref.finalize` | 比 `__del__` 可控 | 仍需 GC 回收才触发；闭包捕获要小心 |
| 显式 `Destroy` / 作用域结束 | 确定 | 与 VBS 语义不符（VBS 是引用计数） |

**建议**：默认用 `weakref.finalize(self, _ovpp_term, self._fields)` 近似；同时在文档里明确"析构时机不保证，不要依赖它释放关键资源"。`Class_Terminate` 里抛出的异常一律吞掉并记入 `Err`（VBS 也不允许从这里往外抛）。

### 3.11 `Option Explicit` 与对象

- 脚本级/过程级变量：现有检查不变。
- 对象字段：`Option Explicit` 不影响字段（VBS 里字段在类内声明），但**类内未声明的变量**要检查——编译 `Class` 时把字段名和方法参数加入作用域声明集即可。

---

## 4. 生成代码示意

```vb
' 输入
Class Counter
    Private n
    Public Property Get Value()
        Value = n
    End Property
    Public Sub Bump(ByVal step)
        n = n + step
    End Sub
    Public Constructor(start)
        n = start
    End Sub
End Class

Set c = CreateObject("Counter", 10)
c.Bump 5
echo c.Value
```

```python
# 输出（示意，实际会带行号与错误处理包装）
class _ovpp_cls_counter:
    _field_names = ("n",)
    _default_member = None

    def __init__(self, start=_ovpp_empty):
        self._fields = {"n": _ovpp_empty}
        self._ctor(start)

    def _ctor(self, start):
        set_member(self, "n", start)

    def _get_value(self):
        value = _ovpp_empty
        value = get_member(self, "n")
        return value

    def _ovpp_sub_bump(self, step):
        set_member(self, "n", get_member(self, "n") + step)

c = _ovpp_builtin.CreateObject("counter", 10)
_ovpp_obj.call(c, "bump", 5)
_ovpp_builtin.Echo(get_member(c, "value"))
```

---

## 5. 分阶段落地计划

| 阶段 | 内容 | 验收标准 | 粗估 |
|---|---|---|---|
| **P1 语法与内置集合** | `Set`/`Let` 语义分流、`CreateObject`/`GetObject`（内置类）、`Dictionary`/`Collection`、`For Each` 支持集合、`Is Nothing` | VBS 超集测试里 `dictionary` / `create_object` 两个 expectedFailure 变成通过；新增 ≥20 条用例 | 1~1.5 天 |
| **P2 类与属性** | `Class ... End Class`、字段、`Property Get/Let/Set`、`Me`、`Class_Initialize`/`Class_Terminate`、默认成员 | `class_definition` / `property_procedures` 通过；对象模型专项用例 ≥30 条 | 2~3 天 |
| **P3 扩展（非 VBS）** | 继承（`Inherits`）、接口、运算符重载 | 必须保证**不改变**任何已有 VBS 语义；扩展语法只允许出现在新关键字后 | 后续 |

每阶段完成后：更新 `doc/VBS_SUPERSET.md`、把对应 expectedFailure 移进 `VBS_FEATURES`、跑全量测试。

---

## 6. 风险与取舍

| # | 风险 | 说明 | 建议 |
|---|---|---|---|
| 1 | **`Let`/`Set` 静态不可判定** | `x = obj` 到底写默认属性还是普通赋值，取决于运行时类型 | 走运行时分派（3.3），不要试图纯静态决定 |
| 2 | 默认成员歧义 | 同一对象可能同时有默认成员和方法同名的属性 | 查找顺序固定为 属性过程 → 字段 → 方法（3.2 表），文档写死 |
| 3 | `Property Set` 与 `Set` 语句 | 两者语法无关但都叫 Set，容易混淆 | 词法上已区分（关键字 + 上下文），测试覆盖两种场景 |
| 4 | COM 缺失（已缓解） | 大量真实 VBS 依赖 `Excel.Application` 等 | v0.9.14 起由 `com-support` 包在 Windows 上真连 COM（封 pywin32，脚本里无需 import）；非 Windows / 关闭开关时明确报错，绝不假装能跑 |
| 5 | `Class_Terminate` 时机 | Python GC 与 VBS 引用计数不同 | 用 `weakref.finalize`，文档声明"不保证"；关键资源用显式 Close |
| 6 | `Private` 字段被 `With` 绕过 | `.n` 在类内部可见，外部不应可见 | 访问器按"当前编译单元"判断可见性，类外访问私有字段报 438 |
| 7 | 属性与方法同名 | VBS 允许 `Property Get X` 和 `Sub X` 同时存在？实际不允许 | 编译期报重名错误，比运行期诡异行为好 |
| 8 | `For Each` 中修改集合 | 行为未定义 | 迭代快照（`list(obj)`）后再遍历，避免 RuntimeError |
| 9 | 类内 `Dim` 与字段 | VBS 类里 `Dim x` 就是字段 | 统一当字段处理，加入 `_field_names` |
| 10 | 循环引用泄漏 | Python GC 能回收，但 `Class_Terminate` 不触发 | 文档说明；提供 `Destroy` 显式清理 |
| 11 | `TypeName` / `IsObject` 对类实例 | 需要返回 `"Counter"` 而不是 Python 类名 | 注册类名映射表 |
| 12 | 与 `Option Explicit` 的交互 | 类内变量未声明要不要报错 | 类内字段视作声明；方法局部变量照常检查 |

---

## 7. 与"VBS 超集"原则的关系

- 本设计完成后，`tests/test_vbs_superset.py` 里的 8 个 expectedFailure 中，
  `class_definition`、`property_procedures`、`create_object`、`dictionary` 4 个应当转为通过；
  `execute`、`eval`、`format_functions`、`option_compare_text` 属于另外的独立特性，另行排期。
- 超集原则要求：**新增的对象语法不得让已有 VBS 程序的行为发生变化**。
  例如引入 `Class` 后，`Set x = Nothing`、`x Is Nothing`、`With`、`Err` 等现有语义必须保持原样——

---

## 8. 实现现状对照（v0.4）

实际落地时做了若干简化，都是为了控制复杂度、且不违反 VBS 语义：

| 设计里的方案 | 实际实现 | 原因 / 影响 |
|---|---|---|
| 字段存 `self._fields` 字典 + 统一访问器 | **直接用 Python 属性**（`self.name`） | 生成代码更短、调试更直观；字段名在编译期已小写化 |
| `get_member` / `set_member` 运行时分派 | 读走 `_ovpp_obj.member_get`，带参调用走 `member_call`，写走普通属性赋值 | 只需在"读一个可能的方法/字段"处做一次运行期判定 |
| `Property Get/Let/Set` → `_get_x` / `_let_x` | → Python `@property` / `@x.setter` | 属性名与字段名不同即可共存；`Let` 与 `Set` 属性过程合并为同一个 setter |
| 默认成员自动取值 | 语法已解析（`Attribute X.VB_UserMemId = 0` 记录到 `_default_member`），**不做自动取值/赋值** | 价值低、二义性高，列入 v0.5 |
| `Class_Terminate` 用 `weakref.finalize` | `__del__` + 异常吞掉 | CPython 下引用计数立即生效，测试可观察 |
| 方法参数支持 ByRef | ✅ v0.6 起统一 Ref 协议：`call_method(obj, name, [Ref...])`，ByVal 在方法体入口解包 | 调用点不需要知道接收者的类 |
| `CreateObject` 走 ProgID 注册表 | 字面量在编译期解析，动态字符串走运行时 `_ovpp_classes` 注册表 | 常见写法零开销 |
| `Collection` 内置类 | 未实现 | VBScript 本身没有内建 `Collection`（属于 VBA） |

已实现并通过测试：`Class`、字段、方法、属性、`Me`、`New`、`CreateObject`、
`Dictionary`、`With`、`Is Nothing`、`TypeName`、`Class_Initialize`、`Class_Terminate`。
对应测试见 `tests/test_objects.py`，超集用例已从 `TestKnownGaps` 移入 `VBS_FEATURES`。
  这也是 P2 验收项之一。
