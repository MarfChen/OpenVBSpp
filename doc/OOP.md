> English | **[中文](OOP.zh-CN.md)**

# Object-Oriented Features Checklist · Object Model Status

> **Summary** — Status of every object-oriented feature: classes, fields,
> visibility, `Property Get/Let/Set`, `Me`, constructor/destructor, the
> `Inherits` extension, method overriding (virtual by default), arity-based
> overloading, `Shared` static members and `Implements` interface contracts.
> Keywords VBScript does not have (`Abstract`, `Static`, `Friend`, `Protected`,
> `Event`) are listed with the reason we deliberately skip them.

> This table answers one question: **which of VBScript's object features exist, how much of it have we implemented, what does not apply, and what have we deliberately left out.**
>
> Legend: ✅ implemented and tested · ⚠️ implemented with a semantic difference · ❌ not applicable (VBS has no such thing) · 🚧 planned (superset extension)

## 1. Classes and Members

| Feature | Status | How it is implemented / Notes |
|---|---|---|
| `Class … End Class` | ✅ | Compiles to `class _ovpp_cls_<name>`, carrying `_vbs_name` |
| Field `Public x` | ✅ | Python instance attribute |
| Field `Private n` / `Dim n` | ✅ | Renamed to `_priv_n`; reads and writes from outside are both rejected (the `__setattr__` whitelist) |
| `Property Get` | ✅ | `@property` |
| `Property Let` | ✅ | `@name.setter` |
| `Property Set` | ⚠️ | **Merged with `Let` into the same setter** (VBS distinguishes "assigning an object" from "assigning a value"; we route both through the setter) |
| `Me` | ✅ | Compiles to `self`; valid only inside class methods |
| `Class_Initialize` | ✅ | Called from `__init__` (no parameters, same as VBS) |
| `Class_Terminate` | ⚠️ | Mapped to `__del__`; **the timing is decided by the CPython GC**, immediate execution is not guaranteed |
| Default member `Attribute X.VB_UserMemId = 0` | ✅ | `x = obj` (Let) picks up the default property automatically; `Set x = obj` keeps the reference |

## 2. Creation and Inspection

| Feature | Status | Notes |
|---|---|---|
| `New Foo` | ✅ | Resolved into a class construction at compile time |
| `CreateObject("脚本类")` | ✅ | Runtime registry |
| `CreateObject("Scripting.Dictionary")` | ✅ | Built-in `Dictionary` |
| `CreateObject("Scripting.FileSystemObject")` | ✅ | Added in v0.9 (text files / directories / paths) |
| `CreateObject("VBScript.RegExp")` / `New RegExp` | ✅ | Added in v0.9.2 |
| `CreateObject("WScript.Shell")` and other COM objects | ✅ | Since v0.9.14, a real connection on Windows through the `com-support` package (requires pywin32; non-Windows platforms raise an explicit error) |
| `GetObject(...)` | ⚠️ | Valid only for script classes / built-in classes (no object persistence); a COM object that is already running can be attached with `GetObject("Excel.Application")` |
| `Is Nothing` / `IsObject` / `TypeName` | ✅ | Semantics match VBS |

## 3. Inheritance and Polymorphism (superset extension)

| Feature | Status | Notes |
|---|---|---|
| `Class Child Inherits Parent` | ✅ | Maps directly onto a Python base class |
| Method overriding | ✅ | **Methods are virtual by default**: a same-named subclass method overrides the base class one and is resolved through the MRO |
| Field/property inheritance | ✅ | Subclass methods can read and write base-class `Public` fields directly; `_ovpp_attrs` merges the base class whitelist |
| Construction chain | ✅ | The subclass `__init__` calls `super().__init__()` first, then initializes its own fields, then runs its own `Class_Initialize` |
| Multiple inheritance | ❌ | Single inheritance only (VBS has no inheritance, this is our addition; single inheritance is sufficient and keeps the semantics clear) |

## 4. Overloading

| Feature | Status | Notes |
|---|---|---|
| Procedure overloading | ✅ | **Dispatched by argument count**, resolved at compile time (zero runtime overhead) |
| Method overloading | ✅ | Generates an `*args` dispatcher that forwards to `_ovpp_m_<name>_<arity>` |
| Overloading by type | ❌ | Needs runtime type dispatch, which would defeat the "compile-time checking" goal; to distinguish types, use different names or a `When` branch |
| Overloading + `Optional`/`ParamArray` | ✅ | Argument-count ranges take part in matching |

## 5. Not Applicable (VBScript Has No Such Thing)

These keywords do not exist in VBS, so we **do not add** them — adding them would mean inventing new syntax that needs a design of its own:

| Keyword | Origin | Our position |
|---|---|---|
| `Abstract` / `MustOverride` | VB.NET | ❌ Not applicable: VBS has no abstract classes |
| `Overridable` / `Overrides` | VB.NET / VBA | ❌ Not applicable: our methods are **overridable by default** and need no modifier |
| `Static` (procedure-level persistent variables) | VBA | ❌ Not applicable: VBS does not support it; use module-level variables when you need persistent state |
| `Shared` (class static members) | VB.NET | 🚧 Planned: can be emulated with module-level variables + `Public` procedures |
| `Friend` / `Protected` | VB.NET | ❌ Not applicable: only the two levels `Public` / `Private` exist |
| `Implements` (interfaces) | VBA / VB.NET | 🚧 Planned: inheritance already provides polymorphism, interfaces will be done when a real need appears |
| `Event` / `RaiseEvent` | VBA / VB.NET | ❌ Not applicable: VBS has no event model; use procedure values (`AddressOf`) when you need callbacks |

## 6. Summary of Semantic Differences from VBS

1. **`Property Set` is merged with `Let`**: in VBS, `Set obj.Prop = x` goes through `Property Set` while `obj.Prop = x` goes through `Property Let`. We route both through the same setter (Python's `@x.setter`).
   Impact: a single property cannot carry two separate `Let` and `Set` code paths. This is extremely rare in practice.
2. **The timing of `Class_Terminate`**: it depends on Python reference counting/GC and is not guaranteed to be called immediately when the scope ends.
3. **Method parameters default to `ByRef`**: same as VBS, implemented through the unified `Ref` protocol (class methods included).
4. **`Private` members are not visible in subclasses**: same as VBS (VBS `Private` is private to the class as well).

## 7. Test Coverage

| Test file | Coverage |
|---|---|
| `tests/test_objects.py` | Classes, fields, methods, properties, `Me`, constructor/destructor, visibility, default members, `GetObject`, method `ByRef` |
| `tests/test_try_catch.py` | Structured exceptions inside class methods |
| `tests/test_overloading.py` | Procedure/method overloading, dispatch by count, error messages |
| `tests/test_option_typing.py` | Type checking on class fields/properties/inheritance chains |
| `tests/test_vbs_superset.py` | Object-related syntax verified item by item in the "superset" matrix |

## 8. Future Extensions (as needed)

1. `Shared` static members (emulated with a module-level variable + a class-name dictionary);
2. `Implements` interfaces (polymorphism is already satisfied by inheritance; waiting for a real need);
3. separating `Property Set` from `Let` (if a compatibility problem shows up that requires the distinction);
4. constructor arguments (`New Foo(1, 2)` → arguments passed to `Class_Initialize`; VBS does not support this, so it is an extension).
