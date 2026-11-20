# 对象模型

> 中文 | **[English](OBJECT_MODEL.md)**

这里讲 OpenVBS++ 里对象、属性、方法调用、默认成员与集合的行为，全部是**脚本作者**
视角；内部设计备忘在 `dev-docs/OBJECT_MODEL_DESIGN.md`。

## 1. 值与对象

VBScript 的变量可以装标量（`Integer`、`String`…）或对象引用，OpenVBS++ 保持这个区分：

- `x = 5` 是值赋值；
- `Set x = SomeObject` 是引用赋值；
- `x = SomeObject` 会把对象的**默认成员**值赋给 `x`；对象没有默认成员时报错 438；
- `x Is Nothing` 判断空引用，不能用 `=`。

`TypeName(x)` 返回 VBS 类型名；对对象返回类名（`"Counter"`，不是内部的 Python 名）。
`IsObject(x)` 对任何对象引用为真。

## 2. 类

```vb
Class Counter
    Private n
    Public Property Get Value()
        Value = n
    End Property
    Public Sub Bump(ByVal step)
        n = n + step
    End Sub
    Public Sub Class_Initialize()
        n = 0
    End Sub
End Class

Set c = New Counter
c.Bump 5
echo c.Value          ' 5
```

- 字段用类体里的 `Public` / `Private` / `Dim` 声明；
- 私有字段与私有过程在编译期改名，类外访问会报 438（和 VBS 对未声明成员的行为一致）；
- `Me` 是当前对象，在类外使用它是编译错误；
- 创建对象时调用 `Class_Initialize`，回收时调用 `Class_Terminate`。**不要依赖析构发生的
  确切时机**——需要确定性释放的资源请写显式的 `Close` / `Dispose` 方法；
- `Class Child Inherits Parent`（超集扩展）继承字段、属性与方法，覆盖按正常方法解析规则生效。

## 3. 属性

| VBS | 含义 |
|---|---|
| `Property Get X()` | 读 `obj.X` |
| `Property Let X(v)` | `obj.X = v` |
| `Property Set X(v)` | `Set obj.X = v` |

同名的 `Property Let` 与 `Property Set` 可以共存，由语句形式决定走哪一个。在 `Get` 过程里
给属性名赋值（`X = value`）就是设置返回值。

属性可以带参数（`Property Get Item(i)`），此时 `obj.Item(3)` 会按参数读取。

## 4. 默认成员

类可以声明默认成员：

```vb
Class Bag
    Public Attribute Value
End Class
```

之后 `v = bag` 读默认值，`bag = v` 通过 `Property Let Value` 写入。没有默认成员的类被当作
值使用时报错 438。

## 5. 方法调用

`obj.Method a, b` 与 `Call obj.Method(a, b)` 等价。参数遵循 VBS 规则：`ByRef`（默认）在调用点
写回，`ByVal` 不写回；`Dictionary` 这类内置对象按值接收。

`New` 与 `CreateObject` 都能创建脚本类：

```vb
Set a = New Counter
Set b = CreateObject("Counter")
```

参数会传给 `Class_Initialize`。

## 6. `With`

```vb
With c
    .Bump 2
    echo .Value
End With
```

表达式只求值一次，块内的 `.name` 针对该对象解析；块内不能重新绑定对象。

## 7. 集合与 `For Each`

| 对象 | 成员 | 迭代 |
|---|---|---|
| `Dictionary` | `Add` `Exists` `Item` `Remove` `RemoveAll` `Count` `Keys` `Items` | 迭代**键** |
| 数组 | `UBound` `LBound`、下标 | 迭代元素 |
| 脚本类 | 自己定义的成员 | 按自身实现迭代 |

```vb
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
For Each k In d
    echo k & " = " & d.Item(k)
Next
```

`For Each` 会先给集合做快照再遍历，所以循环里删除元素不会炸——但本次迭代看到的是快照。

## 8. `CreateObject` 的目标

| ProgID | 行为 |
|---|---|
| `Scripting.Dictionary` | 内置 `Dictionary` |
| `Scripting.FileSystemObject` | 内置文件系统对象 |
| `VBScript.RegExp` | 内置正则对象 |
| 脚本类名 | 该脚本类（`CreateObject("Counter")`） |
| Windows COM ProgID（`WScript.Shell`、`Excel.Application`…） | Windows 上由 `com-support` 包真连 COM；其它平台明确报错 |

`GetObject` 接受同一批目标（Windows 上可以连接已在运行的 COM 对象）。

COM 是可选的：非 Windows 平台、设了 `OPENVBS_COM=0`、或缺少 `pywin32` 时，报错信息会指明
到底是哪一种情况——语言不会假装 COM 对象存在。

## 9. 错误

与对象相关的失败沿用 VBS 错误号，并附带行号与源码行：

| 错误号 | 出现时机 |
|---|---|
| 438 | 成员不存在、类外访问私有成员、对象当值用却没有默认成员 |
| 450 | 方法实参个数不对 |
| 424 | 需要对象（对非对象用 `Set`，或在 `Empty`/`Null` 上取成员） |

这些错误可以通过 `Err.Number` / `Err.Description` 读、可以被 `Try ... Catch` 捕获，也会在
CLI 的 `--json` 输出里带同样的编号。
