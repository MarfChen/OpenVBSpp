# Object model

> English | **[中文](OBJECT_MODEL.zh-CN.md)**

How objects, properties, method calls, default members and collections behave in
OpenVBS++. Everything here is what a **script author** sees; the internal design notes live
in `dev-docs/OBJECT_MODEL_DESIGN.md`.

## 1. Values and objects

A VBScript variable can hold a scalar (`Integer`, `String`, …) or an object reference.
OpenVBS++ keeps that distinction:

- `x = 5` assigns a value;
- `Set x = SomeObject` assigns a reference;
- `x = SomeObject` assigns the object's **default member** value (and fails with error 438
  when the object has no default member);
- `x Is Nothing` tests for an empty reference; `=` cannot be used for that.

`TypeName(x)` returns the VBS type name, and for objects the class name (`"Counter"`, not
the internal Python name). `IsObject(x)` is true for any object reference.

## 2. Classes

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

- Fields are declared with `Public`, `Private` or `Dim` inside the class body.
- Private fields and procedures are renamed at compile time; touching them from outside
  raises error 438, exactly as VBS does for undeclared members.
- `Me` refers to the current object; using it outside a class is a compile error.
- `Class_Initialize` runs when the object is created, `Class_Terminate` when it is
  collected. **Do not rely on the exact moment of termination** — call an explicit `Close`
  or `Dispose` method for resources that must be released deterministically.
- `Class Child Inherits Parent` (a superset extension) inherits fields, properties and
  methods; overriding works by normal method resolution.

## 3. Properties

| VBS | Meaning |
|---|---|
| `Property Get X()` | reading `obj.X` |
| `Property Let X(v)` | `obj.X = v` |
| `Property Set X(v)` | `Set obj.X = v` |

`Property Let` and `Property Set` with the same name coexist: the statement form decides
which one runs. Inside a `Get` procedure, assigning to the property name (`X = value`) sets
the return value.

Properties may take arguments (`Property Get Item(i)`), in which case `obj.Item(3)` reads
through them.

## 4. Default members

A class may declare a default member:

```vb
Class Bag
    Public Attribute Value
End Class
```

Then `v = bag` reads the default value and `bag = v` writes it (through `Property Let
Value`). Without a default member, using an object as a plain value raises error 438.

## 5. Method calls

`obj.Method a, b` and `Call obj.Method(a, b)` are equivalent. Parameters follow VBS rules:
`ByRef` (the default) is written back at the call site, `ByVal` is not. Built-in objects
such as `Dictionary` take their values by value.

`CreateObject` and `New` both work for script classes:

```vb
Set a = New Counter
Set b = CreateObject("Counter")
```

Argument lists are passed on to `Class_Initialize`.

## 6. `With`

```vb
With c
    .Bump 2
    echo .Value
End With
```

The expression is evaluated once; inside the block `.name` resolves against that object.
The block cannot rebind the object.

## 7. Collections and `For Each`

| Object | Members | Iteration |
|---|---|---|
| `Dictionary` | `Add` `Exists` `Item` `Remove` `RemoveAll` `Count` `Keys` `Items` | iterates **keys** |
| Array | `UBound` `LBound`, indexing | iterates elements |
| Script class | whatever it defines | iterates according to its own implementation |

```vb
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", 1
d.Add "b", 2
For Each k In d
    echo k & " = " & d.Item(k)
Next
```

`For Each` snapshots the collection before iterating, so removing items inside the loop does
not blow up — but the iteration still sees the snapshot.

## 8. `CreateObject` targets

| ProgID | Behaviour |
|---|---|
| `Scripting.Dictionary` | built-in `Dictionary` |
| `Scripting.FileSystemObject` | built-in file-system object |
| `VBScript.RegExp` | built-in regular-expression object |
| a script class name | that class (`CreateObject("Counter")`) |
| a Windows COM ProgID (`WScript.Shell`, `Excel.Application`, …) | real COM through the `com-support` package on Windows; a clear error elsewhere |

`GetObject` accepts the same set (a running COM object can be attached on Windows).

COM is optional: on non-Windows platforms, with `OPENVBS_COM=0`, or when `pywin32` is
missing, the error message says exactly which of those is the case — the language never
pretends a COM object exists.

## 9. Errors

Object-related failures use the VBS error numbers, with the line number and source line
attached:

| Error | When |
|---|---|
| 438 | member does not exist, private member accessed from outside, object used as a value without a default member |
| 450 | wrong number of arguments to a method |
| 424 | object required (`Set` on a non-object, or a member access on `Empty`/`Null`) |

These surface through `Err.Number` / `Err.Description`, through `Try ... Catch`, or through
the CLI (`--json` reports the same numbers).
