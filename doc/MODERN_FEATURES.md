> English | **[中文](MODERN_FEATURES.zh-CN.md)**

# Modern Scripting-Language Features at a Glance

> **Summary** — A checklist of modern scripting-language features
> (first-class functions, static members, modules, pattern matching,
> async, generics, tooling) mapped to ✅ implemented / 🚧 planned / ❌
> not applicable, so contributors can see where help is wanted.

OpenVBS++ is positioned as a **superset of VBScript** (100% syntax-compatible) that also has the capabilities a modern scripting language is expected to offer.
This document uses one table per area to show "modern feature → current status → how it lands".

Legend: ✅ implemented · 🚧 planned · ❌ not doing (conflicts with VBS or is meaningless there)

## 1. Language Core

| Modern feature | Status | Notes |
|---|---|---|
| Package/module-level functions and procedures | ✅ | `Sub` / `Function`; may be used before they are defined |
| First-class functions (procedures as values) | 🚧 | Today a procedure name can only be called; v0.7 plans to introduce a `Function` value type |
| By-value/by-reference parameters | ✅ | `ByVal` / `ByRef`, including class methods (unified Ref protocol) |
| Default arguments | 🚧 | `Optional` is not implemented yet (VBS has it, so this is a superset gap) |
| Variadic arguments | 🚧 | `ParamArray` is planned |
| Overloading | ✅ | **Dispatched by argument count**, resolved at compile time (v0.6.1) |
| Recursion / closures | ✅ / ❌ | Recursion is supported; closures are not VBS semantics, so we do not do them |
| Classes and objects | ✅ | `Class` / `Property` / `Me` / constructor and destructor |
| Inheritance and overriding | ✅ | `Inherits` (a superset extension), method overriding goes through the MRO |
| Interfaces / polymorphism | 🚧 | Inheritance already provides polymorphism; interfaces are planned |
| Exception handling | ✅ | `On Error Resume Next` / `GoTo 0` / `Err.Raise` |
| Structured exceptions (try/catch) | 🚧 | `Try ... Catch ... End Try` is planned as a superset extension |
| Compile-time type checking | ✅ | `Option Typing` + `->` annotations, reported **at compile time** |
| Type inference | ✅ | Literals, operations, built-in function returns, class members, array elements |
| String interpolation $"...{x}..." | ✅ | Expanded into & concatenation at compile time (v0.9.1) |
| Optional chaining ?.b / null coalescing ?? b | ✅ | Safe for Nothing/Empty/Null (v0.9.1) |
| Compound assignment += -= *= /= &= | ✅ | Purely new syntax (v0.9.1) |
| Generics | 🚧 | Depends on a type-system extension, planned for later |
| Pattern matching | ❌ | `Select Case` already covers the VBS semantics |

## 2. Engineering Capabilities

| Modern feature | Status | Notes |
|---|---|---|
| Modularity (multiple files) | 🚧 | Single file for now; `Include` is planned |
| Package manager | 🚧 | Called `over` in the design, see the README roadmap |
| Static analysis / type checking | ✅ | `Option Typing`, completed at compile time |
| Debug information (line numbers, source lines) | ✅ | Both syntax and runtime errors carry `Line N` plus the source line |
| AST introspection / decompilation | ✅ | `-v` prints tokens/AST/generated Python; `ast2str` |
| Code formatting | 🚧 | Planned (an AST-based printer already has a prototype) |
| Language server (LSP) | 🚧 | v1.0 goal |

## 3. Interop and Extension

| Modern feature | Status | Notes |
|---|---|---|
| Embedding the host language | ✅ | `#python ... #endpython`, line numbers stay unchanged. Requires an explicit import at the top of the script: `<script src="python-embed" />` |
| Sharing objects across languages | ✅ | The `COP` public object pool holds the same object on the Python side. **`Option COP` was removed in 0.9.20**: Python interop (both `#python` blocks and the `COP` pool) now requires `<script src="python-embed" />` at the top of the script, and writing `Option COP` is a syntax error whose message points at that import |
| Dynamic execution | ✅ | `Execute` / `Eval` (they reuse the compiler, global scope) |
| External library calls | 🚧 | Indirectly available through `#python`; a native FFI is not planned |
| Standard library | ✅ | `ovpprt`: strings/arrays/dates/formatting/collections/JSON/regular expressions/filesystem |
| Regular expressions | ✅ | The `RegExp` object: `Test` / `Replace` / `Execute` (v0.9.2) |

## 4. Runtime Quality

| Modern feature | Status | Notes |
|---|---|---|
| Immutable values / value semantics | ✅ | numbers/strings/booleans; objects are references |
| Null safety | ✅ | Three states `Empty` / `Null` / `Nothing`, semantics identical to VBS |
| Garbage collection | ✅ | Left to CPython; `Class_Terminate` uses `__del__` |
| Iterator protocol | ✅ | `For Each` supports arrays, `Dictionary` and classes |
| Unicode | ✅ | UTF-8 end to end, Chinese identifiers supported |
| Portable | ✅ | Pure standard library, no platform binding |

## 5. Next Steps (v0.7 proposals)

1. `Optional` / `ParamArray` (VBS has them, so this is a superset gap and the highest priority);
2. `Try ... Catch ... End Try` (a superset extension, coexisting with `On Error`);
3. procedures as values / simple closures;
4. multi-file `Include` + a prototype of the `over` package manager;
5. the `RegExp` object.
