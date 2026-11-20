# OpenVBS++ object model design report

> English | **[中文](OBJECT_MODEL_DESIGN.zh-CN.md)**

> Version: v0.4 (P1 + P2 landed), reviewed at v0.9.20
> Audience: the project author and future maintainers
> Status: **P1 (collections / `Set` / `CreateObject`) and P2 (`Class` / `Property` / `Me`) are implemented**;
> differences between this document and the implementation are in §8 "Implementation status".
> P3 (super-set extensions such as inheritance) is still a plan.
> User-facing documentation lives in `doc/OBJECT_MODEL.md` (Chinese: `doc/OBJECT_MODEL.zh-CN.md`);
> this file records design trade-offs and the evolution plan only, and must not duplicate that document.

---

## 1. Goals and hard constraints

### 1.1 Prerequisites that must hold

| Constraint | Explanation |
|---|---|
| **VBS superset** | Every legal VBScript program (including `Class` / `Property` / `Set` / `CreateObject`) must compile. Object syntax used to be the largest gap; after P1/P2 it is covered (see `doc/VBS_SUPERSET.md`). What is still missing is `Declare` and the `Collection` built-in class. |
| **Do not break existing semantics** | Current implementation: procedures are renamed to `_ovpp_fn_*` / `_ovpp_sub_*`; ByRef uses `Ref` boxes and write-back at the call site; `Empty` / `Null` / `Nothing` are runtime sentinels; `Option Explicit` is checked at compile time; `On Error Resume Next` works by wrapping statements in try/except plus `_ovpp_err.capture`. The object model must stay **orthogonal** to all of this — no rewrite. |
| **Zero third-party dependencies** | The runtime may only use the standard library (a hard rule of this educational project). |
| **Portable** | No Windows COM launch. VBS objects are IDispatch in the real world; we do not follow that route and instead provide **equivalent Python objects plus same-named built-in classes**. |
| **Friendly errors** | Line number + source line + VBS error number, consistent with the existing `OVPP*` error family. |

### 1.2 Design stance

- **Semantics first**: guarantee "the same VBS program behaves the same in our implementation" before performance or elegance.
- **Decide at compile time when possible**: `Set`/`Let` distinction, default-member resolution, `Option Explicit` and field visibility are all resolved during compilation where they can be.
- **No COM emulation**: `CreateObject("Scripting.Dictionary")` goes through the built-in registry; COM ProgIDs such as `CreateObject("Excel.Application")` produce a clear error.

---

## 2. VBScript object semantics at a glance

| Concept | VBS semantics | Key trap |
|---|---|---|
| Value vs object | A variable may hold a scalar (Integer/String/…) or an object reference | `obj = other` is **default-property assignment**, not reference assignment |
| `Let` | Explicit value assignment (usually omitted) | `Let obj = x` writes the default property |
| `Set` | Reference assignment, objects only | `Set x = 5` is a runtime error |
| Default member | A class may declare `Attribute Value`, so `obj = v` / `v = obj` access it | Without a default member, `obj = v` raises "object doesn't support this property" |
| `Property Get/Let/Set` | Property procedures, may take parameters | the `Set` property procedure is `Property Set`, same word as the `Set` statement but unrelated |
| `Me` | Reference to the current object | unavailable without a class context |
| `Class_Initialize` / `Class_Terminate` | constructor / destructor | destruction timing is unspecified and must not raise |
| `Nothing` | Empty object reference | compare with `Is`, not `=` |
| `Is` | Object identity comparison | unrelated to the `Case Is` syntax despite the name |
| `For Each` | Iterate a collection/array with iterator semantics | modifying the collection while iterating is undefined |
| `With` | Evaluate the expression once; `.member` inside refers to it | cannot rebind inside the block |
| `CreateObject` | Create a COM object by ProgID | our implementation only knows built-in class names |
| Lifetime | reference counting (COM); the last release triggers `Class_Terminate` | reference cycles leak |

---

## 3. Mapping onto Python

### 3.1 Overall representation

**Every VBS `Class` compiles to one Python class**, named like procedures:

```
VBS:  Class Person
Python: class _ovpp_cls_person: ...
```

Class fields are stored in an **instance dictionary** (`self._fields`) rather than mapped to Python attributes. Reasons:

1. VBS fields are `Variant`s of any type; a dict makes the `Empty` initial value uniform;
2. a `Property` may share a name with a field, and dict storage keeps "field slot" and "property procedure" clearly separate;
3. `With` / default members / `For Each` all need one uniform member lookup entry point.

```python
class _ovpp_cls_person:
    _field_names = ("name", "age")          # known at compile time: init + Option Explicit checks
    _default_member = None                  # default member name (from Attribute Value)

    def __init__(self):
        self._fields = {"name": _ovpp_empty, "age": _ovpp_empty}
        self._init()                        # calls Class_Initialize

    def _init(self):
        # compiled Class_Initialize body (if any)
        ...

    # property procedures
    def _get_name(self):
        return self._fields["name"]
    def _let_name(self, value):
        self._fields["name"] = value

    # methods (VBS Sub/Function)
    def _ovpp_sub_say(self, msg):
        ...
```

### 3.2 Member lookup order (single accessor generated by the compiler)

Every member access (`obj.x`, `With .x`, `Me.x`) compiles to a call on **the same** accessor, which decides the semantics at runtime:

```python
_ovpp_obj.get_member(obj, "x")     # read
_ovpp_obj.set_member(obj, "x", v)  # write (Let)
_ovpp_obj.set_object(obj, "x", v)  # write (Set)
```

Lookup order (same as VBS):

| Order | Target | Behaviour on hit |
|---|---|---|
| 1 | `Property Get x` | call `_get_x()` |
| 2 | field `x` | return `_fields["x"]` |
| 3 | method `x` | return the bound method (for `obj.x(...)`) |
| 4 | built-in member (`Dictionary` etc.) | the built-in class's Python implementation |
| 5 | none of the above | `OVPPObjectNotSupportedError` (438) |

For writes: look for `Property Let/Set` first, then the field, otherwise raise (VBS also forbids adding fields dynamically).

### 3.3 Distinguishing `Let` from `Set`

- Syntax level: `Set x = expr` keeps the `Set` marker; `Let x = expr` and `x = expr` are ordinary assignments.
- Compile time:
  - the target is statically known to be an object (assigned with `Set`, or first `Set` after `Dim`) → reference assignment;
  - the target is a class field → emit `set_member`;
  - **cannot be decided statically** (frequent in VBS, e.g. `x = obj` meaning the default property) → emit a runtime dispatch:

```python
_ovpp_obj.let_assign(current_value, new_value)
# if current_value is an object with a default member -> set_member(default, new_value)
# otherwise -> plain assignment
```

That runtime dispatch is the key to correct `Let`/`Set` semantics; the cost is one type check (acceptable).

### 3.4 `Property Get/Let/Set`

| VBS | Python |
|---|---|
| `Property Get X()` | `def _get_x(self):` |
| `Property Let X(v)` | `def _let_x(self, v):` |
| `Property Set X(v)` | `def _setobj_x(self, v):` |
| parameterised `Property Get X(i)` | `def _get_x(self, i):` (parameters passed as usual) |

- When `Let` and `Set` property procedures share the name `X`, both implementations coexist and the caller's `let` or `set` decides which runs — exactly VBS behaviour.
- `X = value` inside a property body means "set the return value" (for `Get`) per VBS rules: **inside a Get procedure `X = v` sets the return value**, compiled to `ret_x = v` followed by `return ret_x` (reusing the function return-value machinery).

### 3.5 `Me`

`Me` compiles to the first implicit parameter of the method (Python's `self`). `Me` inside a `With` block resolves normally; `Me` outside any class context → compile-time error (same as VBS).

### 3.6 Default member

VBS declares a default member with `Attribute Value`. The compiler records `_default_member`, and the runtime `let_assign` / read paths consult it.

Using a class without a default member as a value → runtime error 438 (`Object doesn't support this property or method`).

### 3.7 `CreateObject` / `GetObject`

```vb
Set d = CreateObject("Scripting.Dictionary")
```

Compilation strategy: `CreateObject` is a built-in function; when the argument is a **string literal** the class is resolved at compile time, when it is a variable the runtime registry is used.

| ProgID | Our implementation |
|---|---|
| `Scripting.Dictionary` | `ovpprt.collections.Dictionary` |
| `Scripting.FileSystemObject` | implemented (filesystem subset) |
| `WScript.Shell` and other host objects | implemented on Windows through the `com-support` package; a clear error elsewhere |
| others (`Excel.Application` …) | real COM on Windows via `com-support`; otherwise `OVPPObjectNotSupportedError` mentioning "COM objects are not supported" |

`GetObject` works the same way: only objects we can provide, otherwise an error.

### 3.8 Built-in collections

| VBS | Python implementation |
|---|---|
| `Dictionary`: `Add` / `Exists` / `Item` / `Remove` / `RemoveAll` / `Count` / `Keys` / `Items` | dict wrapper, insertion order preserved (Python 3.7+ is ordered) |
| `Collection`: `Add` / `Item` / `Remove` / `Count` | list wrapper, **1-based** index (VBS `Collection` is 1-based!) |
| arrays | the existing list |

`For Each` iterates **Keys** for a `Dictionary` (VBS behaviour), elements for a `Collection`, elements for an array — which is exactly the Python iteration protocol, so `ast.For` works as-is.

### 3.9 Interaction with `With` / `For Each` / `Err`

- `With obj`: `obj` is evaluated once into a temporary (already the case), and `.x` inside goes through the same `get_member` / `set_member`.
- `For Each x In obj`: `iter(obj)`; object classes may implement `_iterate()` returning keys or elements.
- `Err`: a built-in singleton, not a script class; `Err.Number` etc. are Python `@property` (current implementation).
- Errors raised inside object methods: caught by `_ovpp_err`'s try/except and mapped to the matching `Err.Number`; a method-local `On Error Resume Next` only affects that method's frame (existing `enter`/`leave` machinery).

### 3.10 Lifetime and `Class_Terminate`

| Option | Pros | Cons |
|---|---|---|
| `__del__` | simple | timing unspecified, exceptions swallowed, cycles never trigger |
| `weakref.finalize` | more controllable than `__del__` | still needs GC, watch closure captures |
| explicit `Destroy` / scope end | deterministic | not VBS semantics (VBS is reference counted) |

**Recommendation**: approximate with `weakref.finalize(self, _ovpp_term, self._fields)` by default, and state in the docs that "destruction timing is not guaranteed; do not rely on it to release critical resources". Exceptions raised in `Class_Terminate` are swallowed and recorded in `Err` (VBS does not allow throwing out of it either).

### 3.11 `Option Explicit` and objects

- Script-level / procedure-level variables: existing checks unchanged.
- Object fields: `Option Explicit` does not apply to fields (they are declared inside the class), but **undeclared variables inside a class** must be checked — when compiling a `Class`, add field names and method parameters to the scope declaration set.

---

## 4. Generated code sketch

```vb
' input
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
# output (sketch; the real one carries line numbers and error handling)
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

## 5. Phased delivery plan

| Phase | Content | Acceptance | Estimate |
|---|---|---|---|
| **P1 syntax and built-in collections** | `Set`/`Let` dispatch, `CreateObject`/`GetObject` (built-in classes), `Dictionary`/`Collection`, `For Each` over collections, `Is Nothing` | the `dictionary` / `create_object` expectedFailures in the superset tests turn green; ≥20 new cases | 1–1.5 days |
| **P2 classes and properties** | `Class ... End Class`, fields, `Property Get/Let/Set`, `Me`, `Class_Initialize`/`Class_Terminate`, default member | `class_definition` / `property_procedures` pass; ≥30 object-model cases | 2–3 days |
| **P3 extensions (non-VBS)** | inheritance (`Inherits`), interfaces, operator overloading | must **not** change any existing VBS semantics; extension syntax only after new keywords | later |

After each phase: update `doc/VBS_SUPERSET.md`, move the corresponding expectedFailure into `VBS_FEATURES`, run the full suite.

---

## 6. Risks and trade-offs

| # | Risk | Explanation | Recommendation |
|---|---|---|---|
| 1 | **`Let`/`Set` not statically decidable** | whether `x = obj` writes the default property or assigns depends on the runtime type | runtime dispatch (§3.3); do not attempt pure static resolution |
| 2 | default-member ambiguity | an object may have a default member and a same-named property | fixed lookup order: property procedure → field → method (§3.2), documented |
| 3 | `Property Set` vs the `Set` statement | syntactically unrelated but identically named, easy to confuse | already distinguished lexically (keyword + context); tests cover both |
| 4 | missing COM (mitigated) | many real VBS scripts depend on `Excel.Application` etc. | since v0.9.14 the `com-support` package connects to real COM on Windows (wrapping pywin32, no import in scripts); on other platforms or when disabled it errors clearly and never pretends to work |
| 5 | `Class_Terminate` timing | Python GC differs from VBS reference counting | use `weakref.finalize` and document "not guaranteed"; critical resources need an explicit Close |
| 6 | `Private` fields bypassed by `With` | `.n` is visible inside the class, not outside | the accessor checks visibility per compilation unit; outside access to a private field raises 438 |
| 7 | property and method with the same name | VBS does not actually allow `Property Get X` and `Sub X` together | compile-time duplicate-name error beats strange runtime behaviour |
| 8 | mutating a collection inside `For Each` | behaviour undefined | iterate over a snapshot (`list(obj)`) to avoid RuntimeError |
| 9 | `Dim` inside a class | in VBS `Dim x` in a class is a field | treat it as a field and add it to `_field_names` |
| 10 | reference-cycle leaks | Python GC collects, but `Class_Terminate` will not fire | document it; offer an explicit `Destroy` |
| 11 | `TypeName` / `IsObject` on class instances | must return `"Counter"`, not the Python class name | keep a class-name registry |
| 12 | interaction with `Option Explicit` | should undeclared class variables error? | class fields count as declared; method locals are checked normally |

---

## 7. Relationship to the "VBS superset" principle

- After this design lands, of the 8 expectedFailures in `tests/test_vbs_superset.py`, the four
  `class_definition`, `property_procedures`, `create_object` and `dictionary` should turn green;
  `execute`, `eval`, `format_functions` and `option_compare_text` are independent features scheduled separately.
- The superset principle requires: **new object syntax must not change the behaviour of existing VBS programs.**
  For example, after `Class` exists, `Set x = Nothing`, `x Is Nothing`, `With` and `Err` must keep working exactly as before.

---

## 8. Implementation status (v0.4)

The implementation simplified several decisions to control complexity without violating VBS semantics:

| Design | Implementation | Reason / impact |
|---|---|---|
| fields in a `self._fields` dict + unified accessor | **plain Python attributes** (`self.name`) | shorter generated code, easier debugging; field names lower-cased at compile time |
| `get_member` / `set_member` runtime dispatch | reads go through `_ovpp_obj.member_get`, calls with arguments through `member_call`, writes are plain attribute assignment | one runtime decision only where a name may be a method or a field |
| `Property Get/Let/Set` → `_get_x` / `_let_x` | → Python `@property` / `@x.setter` | property and field names differ so they coexist; `Let` and `Set` property procedures merge into one setter |
| automatic default-member value | syntax parsed (`Attribute X.VB_UserMemId = 0` recorded as `_default_member`), **no automatic get/set** | low value, high ambiguity; scheduled for v0.5 |
| `Class_Terminate` via `weakref.finalize` | `__del__` with exceptions swallowed | CPython reference counting fires immediately, observable in tests |
| method parameters support ByRef | ✅ since v0.6 a unified `Ref` protocol: `call_method(obj, name, [Ref...])`, ByVal unpacked at method entry | the call site need not know the receiver's class |
| `CreateObject` via ProgID registry | literals resolved at compile time, dynamic strings via the `_ovpp_classes` runtime registry | zero cost for the common spelling |
| `Collection` built-in class | not implemented | VBScript itself has no built-in `Collection` (that is VBA) |

Implemented and covered by tests: `Class`, fields, methods, properties, `Me`, `New`,
`CreateObject`, `Dictionary`, `With`, `Is Nothing`, `TypeName`, `Class_Initialize`,
`Class_Terminate`. See `tests/test_objects.py`; the superset cases moved from
`TestKnownGaps` into `VBS_FEATURES`.
