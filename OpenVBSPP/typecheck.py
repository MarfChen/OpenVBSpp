'''
Copyright 2027 Marf Chen(MarfChen)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Compile-time static checking, active only when `Option Typing` is declared.

Design principles:
  * it runs only when `Option Typing` appears; without it behaviour is plain VBScript;
  * type annotations use `->`, because the colon already separates statements;
  * the checks are conservative: report what is certain and let the rest pass, so a
    weak inference engine never rejects a legal program (`Variant` matches everything).

中文：`Option Typing` 打开时的编译期静态类型检查。只在它出现时才运行（不开启时行为与
plain VBScript. Type annotations use `->` because the colon already separates statements,
and the check is conservative: complain only when certain.
中文：普通 VBS 完全一致）；类型注释用 `->`（冒号已被语句分隔符占用）；只做「能确定就报错、
Uncertain programs pass, so a weak inference never rejects a valid script where `Variant`
matches anything.
中文：不确定就放过」的保守检查，不会因为类型推断能力不足而拒绝合法程序（`Variant` 通配一切）。
'''

from OpenVBSPP.ast_nodes import (TryStmt,Program, Assign, MemberAssign, AssignIndex, Print,
                                 Var, Member, Number, Float, String, Boolean,
                                 EmptyValue, NullValue, NothingValue,
                                 Compare, BinOp, UnaryOp, Call, NewExpr,
                                 While, If, DoLoop, With, ForEach,
                                 SelectCase, For,
                                 DimDecl, ReDimDecl, ConstDecl, EraseStmt,
                                 OptionExplicit, OptionTyping,
                                 OnErrorResumeNext, OnErrorGoToZero,
                                 SubDef, FunctionDef, PropertyDef, ClassDef,
                                 FieldDecl, CallStmt)
from OpenVBSPP.ast_nodes._control import ExitSub, ExitFunction, ExitDo, ExitFor
from OpenVBSPP.exceptions import OVPPError, OVPPTypeError
from OpenVBSPP.ovpprt import builtins as vb_builtins


NUMERIC_RANK = {
    "Byte": 0,
    "Integer": 1,
    "Long": 2,
    "Single": 3,
    "Double": 4,
    "Currency": 5,
}

NUMERIC = frozenset(NUMERIC_RANK)
VAGUE = frozenset({"", "Variant", "Empty", "Null", "Object", "Array"})

# Return types of the built-in functions. / 中文：内置函数的返回类型。
BUILTIN_RETURN = {
    "len": "Integer", "instr": "Integer", "instrrev": "Integer", "asc": "Integer",
    "ubound": "Integer", "lbound": "Integer", "clng": "Long", "cint": "Integer",
    "cbyte": "Byte", "cbool": "Boolean", "cdbl": "Double", "abs": "Double",
    "int": "Integer", "fix": "Integer", "sgn": "Integer", "sqr": "Double",
    "round": "Double", "rnd": "Double", "timer": "Double",
    "isnumeric": "Boolean", "isempty": "Boolean", "isnull": "Boolean",
    "isobject": "Boolean", "isarray": "Boolean", "isdate": "Boolean",
    "left": "String", "right": "String", "mid": "String", "ucase": "String",
    "lcase": "String", "ltrim": "String", "rtrim": "String", "trim": "String",
    "strreverse": "String", "space": "String", "string": "String",
    "replace": "String", "join": "String", "cstr": "String", "chr": "String",
    "hex": "String", "oct": "String", "typename": "String",
    "split": "Array", "array": "Array", "filter": "Array",
    "now": "Date", "date": "Date", "time": "Date",
    "year": "Integer", "month": "Integer", "day": "Integer",
    "hour": "Integer", "minute": "Integer", "second": "Integer",
    "weekday": "Integer", "dateadd": "Date", "datediff": "Integer",
    "strcomp": "Integer", "msgbox": "Integer", "inputbox": "String",
    "createobject": "Object",
}

# Element type of a subscripted assignment like `a(1) = ...`: Variant by default.
# 中文：`a(1) = ...` 这类下标的元素类型，默认 Variant。
ELEMENT_TYPE = "Variant"


def _type_of(node):
    return getattr(node, "type_name", None) or "Variant"


def _array_type(dims, element_type):
    """The type of a Dim'd array: `Array(Integer)`; without an element type it is `Array`.

    中文：Dim 的数组类型：`Array(Integer)`；没写元素类型就是 `Array`。
    """
    if dims is None:
        return element_type or "Variant"
    if element_type:
        return f"Array({element_type})"
    return "Array"


def _element_of(type_name):
    """Pull the element type out of `Array(Integer)`.

    中文：从 `Array(Integer)` 里取出 Integer。
    """
    if type_name and type_name.startswith("Array(") and type_name.endswith(")"):
        return type_name[6:-1]
    return ELEMENT_TYPE


class TypeInfo:
    """The signature of one procedure, property or method.

    中文：一个过程 / 属性 / 方法的签名。
    """

    __slots__ = ("params", "return_type", "name")

    def __init__(self, name, params, return_type):
        self.name = name
        self.params = params              # list[(name, type)]
        self.return_type = return_type or "Variant"


class ClassInfo:
    __slots__ = ("name", "fields", "props", "methods", "default_member", "base")

    def __init__(self, name):
        self.name = name
        self.fields = {}                  # canon -> type
        self.props = {}                   # canon -> (get_type, let_type)
        self.methods = {}                 # canon -> TypeInfo
        self.default_member = None
        self.base = None


def _canon(name):
    return name.lower()


class _Scope:
    def __init__(self, parent=None):
        self.vars = {}
        self.parent = parent
        self.return_var = None
        self.return_type = None
        self.consts = set()

    def declare_const(self, name):
        key = _canon(name)
        self.vars[key] = "Variant"
        self.consts.add(key)

    def in_consts(self, name):
        key = _canon(name)
        scope = self
        while scope is not None:
            if key in scope.consts:
                return True
            scope = scope.parent
        return False

    def declare(self, name, type_name):
        self.vars[_canon(name)] = type_name or "Variant"

    def lookup(self, name):
        key = _canon(name)
        scope = self
        while scope is not None:
            if key in scope.vars:
                return scope.vars[key]
            scope = scope.parent
        return None


class TypeChecker:
    def __init__(self, program):
        self.program = program
        self.procs = {}
        self.classes = {}
        self.option_typing = False

    # ------------------------------------------------------------------ entry / 中文：入口
    def check(self):
        self._collect(self.program.body)
        if not self.option_typing:
            return

        module_scope = _Scope()
        for stmt in self.program.body:
            if isinstance(stmt, DimDecl):
                for name, dims, type_name in stmt.items:
                    module_scope.declare(name, _array_type(dims, type_name))
            elif isinstance(stmt, ConstDecl):
                module_scope.declare(stmt.name, "Variant")

        for stmt in self.program.body:
            self._check_stmt(stmt, module_scope)

    # ------------------------------------------------------------------ collection / 中文：收集
    def _collect(self, body):
        for stmt in body:
            if isinstance(stmt, OptionTyping):
                self.option_typing = True

            elif isinstance(stmt, (SubDef, FunctionDef)):
                key = _canon(stmt.name)
                overloads = self.procs.setdefault(key, [])
                if any(len(item.params) == len(stmt.params) for item in overloads):
                    raise OVPPError(
                        f"Procedure '{stmt.name}' is already defined with "
                        f"{len(stmt.params)} parameter(s)", line_number=stmt.line)
                overloads.append(TypeInfo(
                    stmt.name,
                    [(p.name, _type_of(p)) for p in stmt.params],
                    stmt.return_type if isinstance(stmt, FunctionDef) else None))

            elif isinstance(stmt, ClassDef):
                self._collect_class(stmt)

    def _collect_class(self, node):
        info = ClassInfo(node.name)
        info.base = _canon(node.base) if getattr(node, "base", None) else None
        for member in node.body:
            if isinstance(member, FieldDecl):
                info.fields[_canon(member.name)] = (
                    "Array" if member.dims is not None else _type_of(member))
            elif isinstance(member, PropertyDef):
                if member.kind == "get":
                    info.props.setdefault(_canon(member.name), [None, None])[0] = _type_of(member)
                else:
                    param_type = _type_of(member.params[0]) if member.params else "Variant"
                    info.props.setdefault(_canon(member.name), [None, None])[1] = param_type
            elif isinstance(member, (SubDef, FunctionDef)):
                key = _canon(member.name)
                info.methods.setdefault(key, {})[len(member.params)] = TypeInfo(
                    member.name,
                    [(p.name, _type_of(p)) for p in member.params],
                    member.return_type if isinstance(member, FunctionDef) else None)
        self.classes[_canon(node.name)] = info

    # ------------------------------------------------------------------ statements / 中文：语句
    def _check_stmt(self, node, scope):
        if isinstance(node, (SubDef, FunctionDef)):
            self._check_procedure(node, scope)
            return

        if isinstance(node, ClassDef):
            for member in node.body:
                if isinstance(member, (SubDef, FunctionDef, PropertyDef)):
                    self._check_procedure(member, scope, owner=node)
            return

        if isinstance(node, DimDecl):
            for name, dims, type_name in node.items:
                scope.declare(name, _array_type(dims, type_name))
            return

        if isinstance(node, ConstDecl):
            scope.declare_const(node.name)
            self.infer(node.expr, scope)
            return

        if isinstance(node, Assign):
            self._check_assignment(node.name, node.expr, scope, node)
            return

        if isinstance(node, MemberAssign):
            target_type = self._member_type(node.base, node.name, scope, for_write=True)
            self._check_value(target_type, node.expr, scope, node, "member")
            return

        if isinstance(node, AssignIndex):
            element = _element_of(scope.lookup(node.name) or "")
            self._check_value(element, node.expr, scope, node, "array element")
            return

        if isinstance(node, Print):
            self.infer(node.expr, scope)
            return

        if isinstance(node, If):
            self.infer(node.cond, scope)
            for stmt in node.then_body:
                self._check_stmt(stmt, scope)
            for stmt in (node.else_body or []):
                self._check_stmt(stmt, scope)
            return

        if isinstance(node, (While, DoLoop)):
            if getattr(node, "cond", None) is not None:
                self.infer(node.cond, scope)
            if getattr(node, "pre_cond", None) is not None:
                self.infer(node.pre_cond, scope)
            if getattr(node, "post_cond", None) is not None:
                self.infer(node.post_cond, scope)
            for stmt in node.body:
                self._check_stmt(stmt, scope)
            return

        if isinstance(node, For):
            self._check_value("Double", node.start, scope, node, "For start")
            self._check_value("Double", node.end, scope, node, "For end")
            if node.step is not None:
                self._check_value("Double", node.step, scope, node, "For step")
            for stmt in node.body:
                self._check_stmt(stmt, scope)
            return

        if isinstance(node, ForEach):
            self.infer(node.iterable, scope)
            scope.declare(node.var, "Variant")
            for stmt in node.body:
                self._check_stmt(stmt, scope)
            return

        if isinstance(node, With):
            self.infer(node.obj, scope)
            for stmt in node.body:
                self._check_stmt(stmt, scope)
            return

        if isinstance(node, SelectCase):
            self.infer(node.expr, scope)
            for conds, body in node.cases:
                for cond in conds:
                    value = cond.value
                    if isinstance(value, tuple):
                        for item in value:
                            if hasattr(item, "line"):
                                self.infer(item, scope)
                    elif hasattr(value, "line"):
                        self.infer(value, scope)
                for stmt in body:
                    self._check_stmt(stmt, scope)
            for stmt in (node.else_body or []):
                self._check_stmt(stmt, scope)
            return

        if isinstance(node, TryStmt):
            return self._check_try(node, scope)

        if isinstance(node, CallStmt):
            self._check_call(node, scope)

    # ------------------------------------------------------------------ procedures / 中文：过程
    def _check_procedure(self, node, parent_scope, owner=None):
        scope = _Scope(parent_scope)
        for param in node.params:
            scope.declare(param.name, _type_of(param))

        for name, dims, type_name in _iter_dims(node.body):
            scope.declare(name, _array_type(dims, type_name))

        is_function = isinstance(node, FunctionDef) or (
            isinstance(node, PropertyDef) and node.kind == "get")
        if is_function:
            scope.return_var = _canon(node.name)
            scope.return_type = (getattr(node, "return_type", None)
                                 or getattr(node, "type_name", None)
                                 or "Variant")

        for stmt in node.body:
            self._check_stmt(stmt, scope)

        if (is_function and scope.return_type not in VAGUE
                and not _assigns_name(node.body, _canon(node.name))):
            raise OVPPTypeError(
                f"Function '{node.name}' declares return type "
                f"{scope.return_type} but never assigns '{node.name}'",
                line_number=getattr(node, "line", None))

    # ------------------------------------------------------------------ assignment / 中文：赋值
    def _check_try(self, node, scope):
        for stmt in node.body:
            self._check_stmt(stmt, scope)
        for var, when, body in node.handlers:
            if var:
                scope.declare(var, "Variant")
            if when is not None:
                self.infer(when, scope)
            for stmt in body:
                self._check_stmt(stmt, scope)
        for stmt in (node.finally_body or []):
            self._check_stmt(stmt, scope)

    def _check_assignment(self, name, expr, scope, node):
        if scope.in_consts(name):
            raise OVPPTypeError(f"Cannot assign to constant '{name}'",
                                line_number=getattr(node, "line", None))
        target = scope.lookup(name)
        if target is None and scope.return_var == _canon(name):
            target = scope.return_type
        actual = self.infer(expr, scope)
        self._require(target or "Variant", actual, node,
                      f"assignment to '{name}'")

    def _check_value(self, target_type, expr, scope, node, what):
        actual = self.infer(expr, scope)
        self._require(target_type or "Variant", actual, node, what)

    def _require(self, target, actual, node, what):
        if self._compatible(target, actual):
            return
        raise OVPPTypeError(
            f"Type mismatch in {what}: expected {target}, got {actual}",
            line_number=getattr(node, "line", None),
            description="Option Typing turns on compile-time type checking")

    @staticmethod
    def _compatible(target, actual):
        target = target or "Variant"
        actual = actual or "Variant"

        if target in VAGUE or actual in VAGUE:
            return True
        if target == actual:
            return True
        if target in NUMERIC and actual in NUMERIC:
            return True
        if target == "Boolean" and actual in NUMERIC:
            return True
        if target.startswith("Array(") or actual.startswith("Array("):
            if target.startswith("Array(") and actual.startswith("Array("):
                return (target == actual
                        or _element_of(target) in VAGUE
                        or _element_of(actual) in VAGUE)
            return False
        if actual == "Array" or target == "Array":
            return target == actual
        # A class name: the same class, Object and Nothing are allowed.
        # 中文：类名 —— 允许同类、Object、Nothing。
        if target in ("Object",) or actual in ("Object",):
            return True
        return False

    # ------------------------------------------------------------------ calls / 中文：调用
    def _check_call(self, node, scope):
        callee = node.callee
        if not isinstance(callee, Var):
            for arg in node.args:
                self.infer(arg, scope)
            return

        key = _canon(callee.name)
        info = self._resolve(key, len(node.args))
        if info is not None:
            self._check_args_types(info.params, node.args, scope, node, info.name)
            return

        if key in BUILTIN_RETURN:
            for arg in node.args:
                self.infer(arg, scope)
            return

        # Unknown: it may be a class method or an array.
        # 中文：未知 —— 可能是类方法或数组。
        for arg in node.args:
            self.infer(arg, scope)

    # ------------------------------------------------------------------ inference / 中文：推断
    def infer(self, node, scope):
        if node is None:
            return "Variant"

        if isinstance(node, Number):
            return "Integer"
        if isinstance(node, Float):
            return "Double"
        if isinstance(node, String):
            return "String"
        if isinstance(node, Boolean):
            return "Boolean"
        if isinstance(node, EmptyValue):
            return "Empty"
        if isinstance(node, NullValue):
            return "Null"
        if isinstance(node, NothingValue):
            return "Object"
        if isinstance(node, NewExpr):
            return node.name
        if isinstance(node, Var):
            declared = scope.lookup(node.name)
            return declared or "Variant"

        if isinstance(node, Member):
            return self._member_type(node.base, node.name, scope)

        if isinstance(node, BinOp):
            return self._binop_type(node, scope)

        if isinstance(node, UnaryOp):
            inner = self.infer(node.operand, scope)
            if node.op == "NOT":
                return "Boolean" if inner == "Boolean" else "Integer"
            return inner

        if isinstance(node, Compare):
            self.infer(node.left, scope)
            self.infer(node.right, scope)
            return "Boolean"

        if isinstance(node, Call):
            return self._call_type(node, scope)

        return "Variant"

    def _binop_type(self, node, scope):
        left = self.infer(node.left, scope)
        right = self.infer(node.right, scope)

        if node.op == "AMP":
            return "String"
        if node.op in ("AND", "OR", "XOR", "EQV", "IMP"):
            if left == "Boolean" and right == "Boolean":
                return "Boolean"
            return "Integer"
        if node.op == "DIV":
            return "Double"
        if node.op in ("INTDIV", "MOD"):
            return "Integer" if left in NUMERIC and right in NUMERIC else "Double"
        if node.op == "POW":
            return "Double"
        if left in NUMERIC and right in NUMERIC:
            rank = max(NUMERIC_RANK.get(left, 1), NUMERIC_RANK.get(right, 1))
            for name, value in NUMERIC_RANK.items():
                if value == rank:
                    return name
        return "Variant"

    def _call_type(self, node, scope):
        func = node.func
        if isinstance(func, Member):
            for arg in node.args:
                self.infer(arg, scope)
            return self._member_type(func.base, func.name, scope)

        if not isinstance(func, Var):
            return "Variant"

        key = _canon(func.name)
        info = self._resolve(key, len(node.args))
        if info is not None:
            self._check_args_types(info.params, node.args, scope, node, info.name)
            return info.return_type

        if key in BUILTIN_RETURN:
            for arg in node.args:
                self.infer(arg, scope)
            return BUILTIN_RETURN[key]

        # Array subscript: a(i) yields the element type, exact if declared, else Variant.
        # 中文：数组下标 a(i) → 元素类型（声明过就精确，否则 Variant）。
        for arg in node.args:
            self.infer(arg, scope)
        return _element_of(scope.lookup(func.name) or "")

    def _resolve(self, key, count):
        for info in self.procs.get(key, []):
            if len(info.params) == count:
                return info
        return None

    def _check_args_types(self, params, args, scope, node, name):
        if len(params) != len(args):
            return
        for index, ((pname, ptype), arg) in enumerate(zip(params, args)):
            actual = self.infer(arg, scope)
            self._require(ptype, actual, node,
                          f"argument {index + 1} of '{name}'")

    def _member_type(self, base, name, scope, for_write=False):
        if base is None:
            return "Variant"
        owner_type = self.infer(base, scope)
        key = _canon(name)
        current = _canon(owner_type)
        seen = set()
        while current and current not in seen:
            seen.add(current)
            info = self.classes.get(current)
            if info is None:
                return "Variant"
            if key in info.fields:
                return info.fields[key]
            if key in info.props:
                get_type, let_type = info.props[key]
                if for_write and let_type:
                    return let_type
                return get_type or "Variant"
            if key in info.methods:
                returns = {item.return_type for item in info.methods[key].values()}
                return returns.pop() if len(returns) == 1 else "Variant"
            current = info.base
        return "Variant"


def _assigns_name(body, target):
    """Whether the body assigns to target, the function name.

    中文：函数体里是否给 target（函数名）赋过值。
    """
    def walk(nodes):
        for node in nodes:
            if isinstance(node, Assign) and _canon(node.name) == target:
                return True
            if isinstance(node, If):
                if walk(node.then_body) or walk(node.else_body or []):
                    return True
            elif isinstance(node, (While, DoLoop, For, ForEach, With)):
                if walk(node.body):
                    return True
            elif isinstance(node, SelectCase):
                for _conds, branch in node.cases:
                    if walk(branch):
                        return True
                if walk(node.else_body or []):
                    return True
        return False

    return walk(body)


def _iter_dims(body):
    """Collect Dim declarations of a procedure (or class) recursively, for pre-declaring scopes.

    中文：递归收集过程（或类）里的 Dim 声明，供作用域预声明使用。
    """
    def walk(nodes):
        for node in nodes:
            if isinstance(node, DimDecl):
                for item in node.items:
                    yield item
            elif isinstance(node, If):
                yield from walk(node.then_body)
                yield from walk(node.else_body or [])
            elif isinstance(node, (While, DoLoop, For, ForEach, With)):
                yield from walk(node.body)
            elif isinstance(node, SelectCase):
                for _conds, branch in node.cases:
                    yield from walk(branch)
                yield from walk(node.else_body or [])

    yield from walk(body)


def check(program, source=None):
    """Type-check a Program; it only takes effect when Option Typing is on.

    中文：对 Program 做静态类型检查；只在 Option Typing 打开时生效。
    """
    checker = TypeChecker(program)
    checker.check()
    return checker
