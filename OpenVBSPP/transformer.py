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
'''


import ast

from OpenVBSPP.ast_nodes import (Program, Assign, MemberAssign, AssignIndex, UnaryOp,
                                 Print, Var, Member, Number, Float, DateLit, String, Boolean,
                                 EmptyValue, NullValue, NothingValue,
                                 Compare, BinOp, AddressOf,
                                 While, If, DoLoop, Break, ExitDo, ExitFor,
                                 ExitSub, ExitFunction,
                                 With, ForEach, TryStmt,
                                 LabelStmt, GoToStmt, OnErrorGoToLabel,
                                 ResumeStmt,
                                 ScriptEnd, StopStmt,
                                 SelectCase, For,
                                 OptionExplicit, OptionTyping, OptionCompare,
                                 DimDecl, ReDimDecl, ConstDecl,
                                 EraseStmt, OnErrorResumeNext, OnErrorGoToZero,
                                 SubDef, FunctionDef, CallStmt, Call,
                                 ClassDef, FieldDecl, PropertyDef, NewExpr,
                                 AttributeDecl)
from OpenVBSPP.exceptions import (OVPPError, OVPPSyntaxError, OVPPTypeError,
                                  OVPPArgumentError, OVPPProcedureNotDefinedError,
                                  OVPPVariableNotDefinedError,
                                  OVPPObjectNotSupportedError)
from OpenVBSPP.ovpprt import builtins as vb_builtins
from OpenVBSPP.plugins.python_embed import is_placeholder, placeholder_index
from OpenVBSPP.plugins.modules import (is_placeholder as is_module_placeholder,)

# Compatibility alias: a module's placeholder statement / 中文：兼容别名：模块占位语句
_python_placeholder = is_placeholder


# Aliases of the runtime modules; the underscore prefix keeps them away from user variables. / 中文：运行时模块的别名。用下划线前缀是为了避免和用户变量重名。
BOOL_MODULE = "_ovpp_bool"
STR_MODULE = "_ovpp_str"
ARR_MODULE = "_ovpp_arr"
BUILTIN_MODULE = "_ovpp_builtin"
REF_MODULE = "_ovpp_ref"
ERR_MODULE = "_ovpp_err"
OBJ_MODULE = "_ovpp_obj"
COLL_MODULE = "_ovpp_collections"
DYNAMIC_MODULE = "_ovpp_dynamic"
COP_MODULE = "_ovpp_cop"
PYTHON_FEATURE = "python-embed"
REGEXP_MODULE = "_ovpp_regexp"
JSON_MODULE = "_ovpp_json"
IDE_MODULE = "_ovpp_idebridge"
CLASSES_NAME = "_ovpp_classes"
EMPTY_NAME = "_ovpp_empty"
NULL_NAME = "_ovpp_null"
NOTHING_NAME = "_ovpp_nothing"

BUILTIN_NAMES = frozenset(vb_builtins.REGISTRY)

# Member names are lower-cased; a few hit Python keywords and get another spelling. / 中文：成员名统一小写；个别名字在 Python 里是关键字，需要换个写法
MEMBER_ALIASES = {
    "raise": "raise_error",
    "class": "class_value",
}

# The predefined VBScript constants / 中文：VBS 的预定义常量
VBS_CONSTANTS = {
    "vbcrlf": "\r\n",
    "vbcr": "\r",
    "vblf": "\n",
    "vbtab": "\t",
    "vbback": "\b",
    "vbformfeed": "\f",
    "vbverticaltab": "\v",
    "vbnullchar": "\0",
    "vbnullstring": "",
 # The named formats of FormatDateTime
 # 中文：FormatDateTime 的命名格式
  
    "vbgeneraldate": 0,
    "vblongdate": 1,
    "vbshortdate": 2,
    "vblongtime": 3,
    "vbshorttime": 4,
}

# These names never need an Option Explicit declaration / 中文：这些名字不需要 Option Explicit 声明
SPECIAL_NAMES = frozenset(VBS_CONSTANTS) | {"err", "vbtrue", "vbfalse", "cop"}

_FROM_OVPPRT = (
    ("boolean", BOOL_MODULE),
    ("string", STR_MODULE),
    ("array", ARR_MODULE),
    ("builtins", BUILTIN_MODULE),
    ("ref", REF_MODULE),
    ("err", ERR_MODULE),
    ("object", OBJ_MODULE),
    ("collections", COLL_MODULE),
    ("dynamic", DYNAMIC_MODULE),
    ("cop", COP_MODULE),
    ("regexp", REGEXP_MODULE),
    ("jsonobj", JSON_MODULE),
    ("ide_bridge", IDE_MODULE),
)

op_map = {
    "PLUS": ast.Add(),
    "MINUS": ast.Sub(),
    "MUL": ast.Mult(),
    "DIV": ast.Div(),
    "POW": ast.Pow(),
}

cmp_map = {
    "EQ": ast.Eq(),
    "NE": ast.NotEq(),
    "LT": ast.Lt(),
    "GT": ast.Gt(),
    "LE": ast.LtE(),
    "GE": ast.GtE(),
    "IS": ast.Is(),
}

bool_call_map = {
    "AND": "ovpp_and",
    "OR": "ovpp_or",
    "XOR": "ovpp_xor",
    "EQV": "ovpp_eqv",
    "IMP": "ovpp_imp",
}

# VBScript's \ and Mod truncate toward zero, so the runtime does the work. / 中文：VBS 的整除 \ 与 Mod 需要"向零截断"语义，交给运行时函数处理
builtin_call_map = {
    "INTDIV": "ovpp_intdiv",
    "MOD": "ovpp_mod",
}

unary_call_map = {
    "NOT": "ovpp_not",
    "UMINUS": ast.USub(),
    "UPLUS": ast.UAdd(),
}


def canon(name):
    """VBScript identifiers are case-insensitive, so lower-case them.

    中文：VBS 标识符大小写不敏感，统一转小写。
    """
    return name.lower()


def _var_names_in(node):
    """Collect every variable name in an expression, conservatively counting a Member's base too.

    中文：收集表达式里出现的所有变量名（保守：Member 的 base 也算进去）。
    """
    found = set()
    stack = [node]
    seen = set()
    while stack:
        current = stack.pop()
        if current is None or isinstance(current, (str, bytes, int, float, bool)):
            continue
        if isinstance(current, (list, tuple)):
            stack.extend(current)
            continue
        if isinstance(current, dict):
            stack.extend(current.values())
            continue
        values = getattr(current, "__dict__", None)
        if values is None:
            continue
        if id(current) in seen:
            continue
        seen.add(id(current))
        if isinstance(current, Var):
            found.add(canon(current.name))
        stack.extend(values.values())
    return found


def _ref_param_names(node):
    """Work out which ByRef parameters really need a Ref box.

    Only two kinds do: those assigned inside the body, whose new value must travel back,
    and those passed on as arguments to another call, which may write them back. A
    read-only ByRef parameter can pass by value, saving the caller a `Ref(...)` and a
    write-back and sparing the callee a `.v` on every access.

    This is pure static analysis and errs conservative: anything uncertain keeps its Ref.

    中文：算出哪些 ByRef 参数必须真的用 Ref 盒子。只需要两类 —— 过程体里被赋过值的
    (whose new value must go back to the caller), and to a parameter passed on to another call that may write it back. A read-only ByRef
    中文：（新值要回写给调用方），以及作为实参传给别的调用的（对方可能再回写）。只读的 ByRef
    parameter can simply pass by value: the caller saves a `Ref(...)` and a write-back, and the callee never has to go through `.v`.
    中文：参数按值传即可：调用方省掉一次 `Ref(...)` 与一次回写，被调方也不用每次访问走 `.v`。
    The check is purely static and errs on the safe side: when it cannot be sure, the Ref stays.
    中文：纯静态判断，宁可保守，算不出确定就保留 Ref。
    """
    names = frozenset(canon(p.name) for p in node.params if p.byref)
    if not names:
        return frozenset()

    needed = set()
    stack = list(node.body)
    seen = set()
    while stack:
        current = stack.pop()
        if current is None or isinstance(current, (str, bytes, int, float, bool)):
            continue
        if isinstance(current, (list, tuple)):
            stack.extend(current)
            continue
        values = getattr(current, "__dict__", None)
        if values is None:
            continue
        if id(current) in seen:
            continue
        seen.add(id(current))
        if isinstance(current, (Assign, AssignIndex, ReDimDecl)):
            needed.add(canon(current.name))
        elif isinstance(current, (For, ForEach)):
            needed.add(canon(current.var))
        elif isinstance(current, (Call, CallStmt)):
            for arg in current.args:
                needed.update(_var_names_in(arg))
        stack.extend(values.values())

    return frozenset(needed & names)


def _needs_error_frame(body):
    """Whether a procedure body really needs an error frame (enter / leave).

    Only bodies using an On Error statement, Resume, or reading `Err` need one. The rest
    cannot catch an error under VBScript semantics, and dropping the frame drops two
    function calls per invocation.

    中文：过程体是否真的需要错误处理帧（enter / leave）。只有出现 On Error 系列语句、
    Resume or refers to `Err` directly; any other procedure cannot catch an error under VBScript semantics,
    中文：Resume，或者直接引用了 `Err` 时才需要；其余过程按 VBS 语义根本不可能捕获错误，
    so skipping the frame skips two function calls per invocation.
    中文：省掉帧就省掉每次调用的两次函数调用。
    """
    stack = list(body)
    seen = set()
    while stack:
        current = stack.pop()
        if current is None or isinstance(current, (str, bytes, int, float, bool)):
            continue
        if isinstance(current, (list, tuple)):
            stack.extend(current)
            continue
        if isinstance(current, (OnErrorResumeNext, OnErrorGoToZero,
                                OnErrorGoToLabel, ResumeStmt)):
            return True
        if isinstance(current, Var) and canon(current.name) == "err":
            return True
        values = getattr(current, "__dict__", None)
        if values is None:
            continue
        if id(current) in seen:
            continue
        seen.add(id(current))
        stack.extend(values.values())
    return False


def _writes_name(body, name):
    """Whether a variable is assigned somewhere in this procedure or loop body.

    中文：过程体（或循环体）里是否给某个变量赋过值。
    """
    stack = list(body)
    seen = set()
    while stack:
        current = stack.pop()
        if current is None or isinstance(current, (str, bytes, int, float, bool)):
            continue
        if isinstance(current, (list, tuple)):
            stack.extend(current)
            continue
        values = getattr(current, "__dict__", None)
        if values is None:
            continue
        if id(current) in seen:
            continue
        seen.add(id(current))
        if isinstance(current, (Assign, AssignIndex, ReDimDecl)):
            if canon(current.name) == name:
                return True
        elif isinstance(current, (For, ForEach)):
            if canon(current.var) == name:
                return True
        stack.extend(values.values())
    return False


def _const_step(node):
    """Return the value of `Step` when it is a literal constant, else None.

    中文：`Step` 是字面量常量时返回它的数值，否则返回 None。
    """
    step = node.step
    if step is None:
        return 1
    if isinstance(step, (Number, Float)):
        return step.val
    if isinstance(step, UnaryOp) and step.op == "-":
        inner = step.operand
        if isinstance(inner, (Number, Float)):
            return -inner.val
    return None


class _ProcInfo:
    __slots__ = ("name", "kind", "params", "py_name", "ref_params")

    def __init__(self, node, py_name):
        self.name = node.name
        self.kind = node.kind
        self.params = list(node.params)
        self.py_name = py_name
        self.ref_params = _ref_param_names(node)

    @property
    def arity(self):
        return len(self.params)

    def min_arity(self):
        params = self.params
        if params and params[-1].paramarray:
            params = params[:-1]
        return sum(1 for p in params if not p.optional)

    def accepts(self, count):
        params = self.params
        if params and params[-1].paramarray:
            return count >= self.min_arity()
        return self.min_arity() <= count <= len(params)

    def describe_arity(self):
        params = self.params
        low = self.min_arity()
        if params and params[-1].paramarray:
            return f"{low}+"
        high = len(params)
        return str(low) if low == high else f"{low}-{high}"


class Transformer:
    """Turn the OpenVBS++ AST into the standard library ast.

    Two passes:
      1. `_collect()` scans the whole script first, registering procedure signatures,
         module-level variables and Option Explicit;
      2. statements are then translated one by one, every node carrying its source line.

    中文：把 OpenVBS++ AST 转成标准库 ast。两遍处理 —— 先 `_collect()` 扫描整个脚本，
    It records procedure signatures, module-level variables and Option Explicit, then emits a Python AST statement by statement, with every node carrying
    中文：登记过程签名、模块级变量、Option Explicit；再逐语句生成 Python AST，每个节点都带上
    its source line number.
    中文：源码行号。
    """

    def __init__(self, python_blocks=None, features=None):
        self._tmp_counter = 0
        self.python_blocks = python_blocks or {}
        self.features = frozenset(features or ())
        self.class_methods = {}         # canon(cls) -> {canon(method) -> [Param]}
        # static inference result: the class each variable belongs to
        # 中文：静态推断出的变量所属类
        self._var_classes = {}
        self.procs = {}                 # canon(name) -> _ProcInfo
        # Module-level names that are visible: declared, or assigned somewhere
        # 中文：模块级可见的名字（声明过或被赋值过）
        self.module_vars = set()
        # Names explicitly declared at module level (Dim / Const / ReDim)
        # 中文：模块级显式声明的名字（Dim/Const/ReDim）
        self.module_declared = set()
        self.option_explicit = False
        # For Option Explicit: names already declared in each scope
        # 中文：Option Explicit 用：各作用域已声明名字
        self._declared_stack = []
        self.byref_params = frozenset()
        self.current = None
        self.ret_var = None
     # A ByRef call inside an expression queues its box-building and write-back statements
     # 中文：表达式里出现 ByRef 调用时，产生的"建盒子 / 写回"语句先挂在这里，
      
     # here; the enclosing statement splices them in, since a Python expression holds no statements.
     # 中文：由外层语句统一插入（Python 的表达式里不能放语句）。
      
        self._pending_pre = []
        self._pending_post = []
     # The With-block stack holds already evaluated object expressions
     # 中文：With 块栈：存的是已经求值好的对象表达式
      
        self._with_stack = []
     # Whether the procedure sits under On Error Resume Next, approximated at compile time
     # 中文：当前过程是否处于 On Error Resume Next 保护之下（编译期近似）
      
        self._on_error_resume = False
     # Classes: canon(name) -> ClassDef
     # 中文：类：canon(name) -> ClassDef
      
        self.classes = {}
        self.option_typing = False
        self._in_class_method = False
        self._class_fields = frozenset()
        self._class_methods = frozenset()
        self._class_props = frozenset()
        self._method_locals = frozenset()
        # Field name -> Python attribute name; private fields are renamed
        # 中文：字段名 -> Python 属性名（私有字段会改名）
        self._field_attrs = {}
        # Method name -> Python method name
        # 中文：方法名 -> Python 方法名
        self._method_attrs = {}
        self.option_compare_text = False
        # The handler section of On Error GoTo <label>
        # 中文：On Error GoTo <label> 的处理段
        self._error_label_tail = None
        # canon(label) -> statement index
        # 中文：canon(label) -> 语句下标
        self._proc_labels = {}
        # The current procedure's statements, for GoTo
        # 中文：当前过程的语句列表（GoTo 用）
        self._proc_body = []
        self._goto_seen = set()
        self._in_handler_tail = False
        self._handler_uses_resume = False
        self._state_machine = False
        self._sm_labels = {}
        self._current_class = None
        self._shared_fields = frozenset()
        self._delegates = set()
        self._param_names = frozenset()

    # entry / 中文：入口
    def collect(self, program):
        """Collect signatures only: across files collect first, then generate code one by one.

        中文：只收集签名（多文件：先合并收集，再逐个生成代码）。
        """
        self._collect(program)

    def transform(self, node):
        if node is None:
            return None

        if isinstance(node, Program):
            self._collect(node)
            self._declared_stack.append(set(self.module_declared))
            try:
                return self._transform_module(node)
            finally:
                self._declared_stack.pop()

        return self._stamp(node, self._dispatch(node))

    def _stamp(self, node, py_node):
        if isinstance(py_node, ast.AST):
            line = getattr(node, "line", None)
            if line and getattr(py_node, "lineno", None) is None:
                py_node.lineno = max(1, int(line))
                py_node.col_offset = max(0, int(getattr(node, "col", 0) or 0))
        return py_node

    def _new_tmp(self, prefix):
        self._tmp_counter += 1
        return f"_{prefix}_{self._tmp_counter}"

    def _single(self, stmts):
        """Fold several statements into one node, wrapping them in `if True:` when there are more.

        中文：把多条语句收成一个节点（多条时包一层 if True:）。
        """
        stmts = [s for s in stmts if s is not None]
        if not stmts:
            return ast.Pass()
        if len(stmts) == 1:
            return stmts[0]
        return ast.If(test=ast.Constant(value=True), body=stmts, orelse=[])

    def _take_pending(self):
        pre, post = self._pending_pre, self._pending_post
        self._pending_pre, self._pending_post = [], []
        return pre, post

    def _guard(self, py_node):
        """Wrap one statement in try/except to implement On Error Resume Next.

        中文：把一条语句包进 try/except，实现 On Error Resume Next。
        """
        record = ast.Expr(value=ast.Call(
            func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                               attr=("record" if self._error_label_tail else "capture"),
                               ctx=ast.Load()),
            args=[ast.Name(id="_ovpp_exc", ctx=ast.Load())],
            keywords=[]))

        if self._error_label_tail and self._handler_uses_resume:
         # The handler contains Resume or Resume Next, so a retry loop wraps this statement.
         # 中文：处理段里有 Resume / Resume Next：用重试循环包住这条语句
          
            clear = ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                   attr="set_resume", ctx=ast.Load()),
                args=[ast.Constant(value="")], keywords=[]))
            tail_loop = ast.While(
                test=ast.Constant(value=True),
                body=list(self._error_label_tail) + [ast.Break()],
                orelse=[])
            take = ast.Assign(
                targets=[ast.Name(id="_ovpp_act", ctx=ast.Store())],
                value=ast.Call(
                    func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                       attr="take_resume", ctx=ast.Load()),
                    args=[], keywords=[]))
            branch_retry = ast.If(
                test=ast.Compare(left=ast.Name(id="_ovpp_act", ctx=ast.Load()),
                                 ops=[ast.Eq()],
                                 comparators=[ast.Constant(value="retry")]),
                body=[ast.Continue()], orelse=[])
            branch_next = ast.If(
                test=ast.Compare(left=ast.Name(id="_ovpp_act", ctx=ast.Load()),
                                 ops=[ast.Eq()],
                                 comparators=[ast.Constant(value="next")]),
                body=[ast.Break()], orelse=[])
            inner = ast.Try(
                body=[py_node],
                handlers=[ast.ExceptHandler(
                    type=ast.Name(id="Exception", ctx=ast.Load()),
                    name="_ovpp_exc",
                    body=[record, clear, tail_loop, take, branch_retry,
                          branch_next, ast.Return(value=None)])],
                orelse=[], finalbody=[])
            return ast.While(test=ast.Constant(value=True), body=[inner], orelse=[])

        if self._error_label_tail:
            body = [record] + list(self._error_label_tail)
            body.append(ast.Return(value=None))
        else:
            body = [record]

        return ast.Try(
            body=[py_node],
            handlers=[ast.ExceptHandler(
                type=ast.Name(id="Exception", ctx=ast.Load()),
                name="_ovpp_exc",
                body=body)],
            orelse=[],
            finalbody=[],
        )

    def _emit(self, stmts, out):
        for stmt in stmts:
            guard = self._on_error_resume or bool(self._error_label_tail)
            out.append(self._guard(stmt) if guard else stmt)

    def _flush_nodes(self, nodes):
        """Convert a statement list item by item, splicing each item's pre and post statements in order.

        中文：把语句列表逐个转换，并把它产生的 pre/post 语句按顺序插进结果。
        """
        out = []
        for node in nodes:
            py_node = self.transform(node)
            if py_node is None:
                continue
            pre, post = self._take_pending()
            self._emit(pre, out)
            self._emit([py_node], out)
            self._emit(post, out)
        return out

    def _stmts(self, nodes):
        return self._flush_nodes(nodes) or [ast.Pass()]

    def _hoist(self, value):
        """Hoist a value expression into a temporary so pre and post statements land at the right place.

        中文：把值表达式提到临时变量，让 pre/post 语句能插在正确位置。
        """
        pre, post = self._take_pending()
        if not pre and not post:
            return [], value
        tmp = self._new_tmp("val")
        stmts = pre + [ast.Assign(targets=[ast.Name(id=tmp, ctx=ast.Store())],
                                  value=value)] + post
        return stmts, ast.Name(id=tmp, ctx=ast.Load())

    # first pass / 中文：第一遍
    def _collect(self, program):
        for stmt in program.body:
            self._collect_stmt(stmt)

    def _collect_stmt(self, node):
        if isinstance(node, OptionExplicit):
            self.option_explicit = True
            return

        if isinstance(node, OptionTyping):
            self.option_typing = True
            return

        if isinstance(node, OptionCompare):
            self.option_compare_text = (node.mode == "text")
            return

        if isinstance(node, ClassDef):
            key = canon(node.name)
            if key in self.classes:
                raise OVPPError(f"Class '{node.name}' is defined more than once",
                                line_number=node.line)
            self.classes[key] = node
            methods = {}
            for member in node.body:
                if isinstance(member, (SubDef, FunctionDef)):
                    methods[canon(member.name)] = list(member.params)
            self.class_methods[key] = methods
            return

        if isinstance(node, (SubDef, FunctionDef)):
            key = canon(node.name)
            prefix = "_ovpp_sub_" if node.kind == "sub" else "_ovpp_fn_"
            overloads = self.procs.setdefault(key, [])
            arity = len(node.params)
            if any(info.arity == arity for info in overloads):
                raise OVPPError(
                    f"Procedure '{node.name}' is already defined with {arity} parameter(s)",
                    line_number=node.line)
         # With overloads the generated name carries the argument count
         # 中文：有重载时，生成的名字带上参数个数
          
            py_name = prefix + key if not overloads else f"{prefix}{key}_{arity}"
            overloads.append(_ProcInfo(node, py_name))
            return

        if isinstance(node, DimDecl):
            for name, _dims, _type in node.items:
                key = canon(name)
                self.module_vars.add(key)
                self.module_declared.add(key)
            return

        if isinstance(node, ConstDecl):
            key = canon(node.name)
            self.module_vars.add(key)
            self.module_declared.add(key)
            return

        if isinstance(node, ReDimDecl):
            key = canon(node.name)
            self.module_vars.add(key)
            self.module_declared.add(key)
            return

        if isinstance(node, (Assign, AssignIndex)):
            self.module_vars.add(canon(node.name))
            return

        if isinstance(node, For):
            self.module_vars.add(canon(node.var))
            for stmt in node.body:
                self._collect_stmt(stmt)
            return

        if isinstance(node, ForEach):
            self.module_vars.add(canon(node.var))
            for stmt in node.body:
                self._collect_stmt(stmt)
            return

        if isinstance(node, If):
            for stmt in node.then_body:
                self._collect_stmt(stmt)
            for stmt in (node.else_body or []):
                self._collect_stmt(stmt)
            return

        if isinstance(node, (While, DoLoop, With)):
            for stmt in node.body:
                self._collect_stmt(stmt)
            return

        if isinstance(node, SelectCase):
            for _conds, body in node.cases:
                for stmt in body:
                    self._collect_stmt(stmt)
            for stmt in (node.else_body or []):
                self._collect_stmt(stmt)
            return

    def _local_declarations(self, body):
        """Collect the names a procedure introduces with Dim, Const and ReDim.

        中文：收集一个过程里 Dim / Const / ReDim 出来的名字。
        """
        found = set()

        def walk(nodes):
            for node in nodes:
                if isinstance(node, DimDecl):
                    for name, _dims, _type in node.items:
                        found.add(canon(name))
                elif isinstance(node, ConstDecl):
                    found.add(canon(node.name))
                elif isinstance(node, ReDimDecl):
                    found.add(canon(node.name))
                elif isinstance(node, If):
                    walk(node.then_body)
                    walk(node.else_body or [])
                elif isinstance(node, (While, DoLoop, For, ForEach, With)):
                    walk(node.body)
                elif isinstance(node, SelectCase):
                    for _conds, branch in node.cases:
                        walk(branch)
                    walk(node.else_body or [])

        walk(body)
        return found

    def _assigned_names(self, body):
        """Collect the names assigned inside a procedure, which decides what must be global.

        中文：收集一个过程里被赋值的名字（用来决定哪些需要 global）。
        """
        found = set()

        def walk(nodes):
            for node in nodes:
                if isinstance(node, (Assign, AssignIndex)):
                    found.add(canon(node.name))
                elif isinstance(node, ReDimDecl):
                    found.add(canon(node.name))
                elif isinstance(node, For):
                    found.add(canon(node.var))
                    walk(node.body)
                elif isinstance(node, ForEach):
                    found.add(canon(node.var))
                    walk(node.body)
                elif isinstance(node, If):
                    walk(node.then_body)
                    walk(node.else_body or [])
                elif isinstance(node, (While, DoLoop, With)):
                    walk(node.body)
                elif isinstance(node, SelectCase):
                    for _conds, branch in node.cases:
                        walk(branch)
                    walk(node.else_body or [])

        walk(body)
        return found

    # module / 中文：模块
    def _runtime_imports(self):
        stmts = [
            ast.ImportFrom(
                module="OpenVBSPP.ovpprt",
                names=[ast.alias(name=module, asname=alias)],
                level=0,
            )
            for module, alias in _FROM_OVPPRT
        ]
        for module, name, alias in (
            ("empty", "EMPTY", EMPTY_NAME),
            ("null", "NULL", NULL_NAME),
            ("nothing", "NOTHING", NOTHING_NAME),
        ):
            stmts.append(ast.ImportFrom(
                module=f"OpenVBSPP.ovpprt.{module}",
                names=[ast.alias(name=name, asname=alias)],
                level=0,
            ))
        return stmts

    def _transform_module(self, node):
        imports = self._runtime_imports()

        classes = [self._stamp(stmt, self._transform_class(stmt))
                   for stmt in node.body if isinstance(stmt, ClassDef)]

        registry = self._classes_registry()

        procs = [self._stamp(stmt, self._transform_proc(stmt))
                 for stmt in node.body
                 if isinstance(stmt, (SubDef, FunctionDef))]

        statements = [stmt for stmt in node.body
                      if not isinstance(stmt, (SubDef, FunctionDef, ClassDef,
                                               OptionExplicit, OptionTyping,
                                               OptionCompare))]
        body = self._flush_nodes(statements)

        return ast.Module(body=imports + classes + [registry] + procs + body,
                          type_ignores=[])

    def _classes_registry(self):
        """Emit `_ovpp_classes = {"foo": _ovpp_cls_foo, ...}` for CreateObject.

        中文：生成 `_ovpp_classes = {"foo": _ovpp_cls_foo, ...}`，供 CreateObject 使用。
        """
        keys = [ast.Constant(value=key) for key in sorted(self.classes)]
        values = [ast.Name(id="_ovpp_cls_" + key, ctx=ast.Load())
                  for key in sorted(self.classes)]
        return ast.Assign(
            targets=[ast.Name(id=CLASSES_NAME, ctx=ast.Store())],
            value=ast.Dict(keys=keys, values=values))

    def _transform_proc(self, node, as_method=False, py_name=None):
        if as_method:
         # Methods stay out of self.procs; each signature is built from its own node
         # 中文：方法不登记在 self.procs 里，按节点自身构造签名
          
            info = _ProcInfo(node, py_name or canon(node.name))
        else:
            overloads = self.procs.get(canon(node.name), [])
            info = next((item for item in overloads
                         if item.arity == len(node.params)), None)
            if info is None:
                raise OVPPError(f"Procedure '{node.name}' was not registered",
                                line_number=node.line)

        saved_with = self._with_stack
        saved_error = self._on_error_resume
        saved_method = self._in_class_method
        saved_var_classes = self._var_classes
        self._var_classes = {}
        self.current = info
        self.ret_var = canon(node.name) if info.kind == "function" else None
        self._param_names = frozenset(canon(p.name) for p in info.params
                                      if not p.is_array)
        if as_method:
            self._in_class_method = True
         # Method arguments always arrive as Refs; ByVal is unwrapped at the top of the body
         # 中文：方法参数统一通过 Ref 传入，ByVal 在函数体开头解包
          
            self.byref_params = frozenset(canon(p.name) for p in info.params if p.byref)
        else:
         # Only ByRef parameters that get written back travel in a Ref box; others pass by value
         # 中文：只有"会被写回"的 ByRef 参数才走 Ref 盒子，其余按值传
          
            self.byref_params = frozenset(canon(p.name) for p in info.params
                                          if p.byref and canon(p.name) in info.ref_params)
        self._with_stack = []
        self._on_error_resume = False

        declared = {canon(p.name) for p in info.params}
        declared |= self._local_declarations(node.body)
        if self.ret_var:
            declared.add(self.ret_var)

        saved_locals = self._method_locals
        if as_method:
            self._method_locals = frozenset(declared)

        self._declared_stack.append(declared)
        try:
            assigned = self._assigned_names(node.body)
            globals_needed = sorted(n for n in assigned
                                    if n in self.module_vars and n not in declared)

            prologue = []
            if globals_needed:
                prologue.append(ast.Global(names=globals_needed))
            for param in info.params:
                if not param.paramarray:
                    continue
                pname = canon(param.name)
                if as_method:
                 # Method arguments are Refs, so read .v
                 # 中文：方法的实参都是 Ref，取 .v
                  
                    items = ast.ListComp(
                        elt=ast.Attribute(value=ast.Name(id="_pa", ctx=ast.Load()),
                                          attr="v", ctx=ast.Load()),
                        generators=[ast.comprehension(
                            target=ast.Name(id="_pa", ctx=ast.Store()),
                            iter=ast.Name(id=pname, ctx=ast.Load()),
                            ifs=[], is_async=0)])
                else:
                    items = ast.Call(func=ast.Name(id="list", ctx=ast.Load()),
                                     args=[ast.Name(id=pname, ctx=ast.Load())],
                                     keywords=[])
             # At runtime a ParamArray is an ordinary array, i.e. a list
             # 中文：ParamArray 在运行时就是一个普通数组（list）
              
                prologue.append(ast.Assign(
                    targets=[ast.Name(id=pname, ctx=ast.Store())],
                    value=items))
            if as_method:
                for param in info.params:
                    if param.byref or param.paramarray:
                        continue
                    pname = canon(param.name)
                    prologue.append(ast.Assign(
                        targets=[ast.Name(id=pname, ctx=ast.Store())],
                        value=ast.Attribute(value=ast.Name(id=pname, ctx=ast.Load()),
                                            attr="v", ctx=ast.Load())))
            else:
                # read-only ByRef parameters pass by value, while cross-module and delegate calls still pass a Ref,
                # 中文：只读的 ByRef 参数按值传；但跨模块 / 委托调用仍然会传 Ref，
                # so unpack uniformly here and accept either convention.
                # 中文：这里统一拆包，两种调用约定都能接住。
                for param in info.params:
                    if not param.byref or param.paramarray:
                        continue
                    pname = canon(param.name)
                    if pname in info.ref_params:
                        continue
                    prologue.append(ast.If(
                        test=ast.Compare(
                            left=ast.Call(func=ast.Name(id="type", ctx=ast.Load()),
                                          args=[ast.Name(id=pname, ctx=ast.Load())],
                                          keywords=[]),
                            ops=[ast.Is()],
                            comparators=[ast.Attribute(
                                value=ast.Name(id=REF_MODULE, ctx=ast.Load()),
                                attr="Ref", ctx=ast.Load())]),
                        body=[ast.Assign(
                            targets=[ast.Name(id=pname, ctx=ast.Store())],
                            value=ast.Attribute(value=ast.Name(id=pname, ctx=ast.Load()),
                                                attr="v", ctx=ast.Load()))],
                        orelse=[]))
            if info.kind == "function":
             # The return value rides in a local named after the function, starting as Empty
             # 中文：返回值用"函数名"这个局部变量承载，先初始化成 Empty
              
                prologue.insert(0, ast.Assign(
                    targets=[ast.Name(id=self.ret_var, ctx=ast.Store())],
                    value=ast.Name(id=EMPTY_NAME, ctx=ast.Load())))

            handler_index, all_stmts = self._prepare_error_labels(node.body)
            if handler_index is None:
                body = self._flush_nodes(node.body)
            elif self._contains_resume(node.body):
             # Resume appears, so the procedure compiles into a state machine dispatching on an index
             # 中文：含 Resume：整个过程编译成状态机（下标分派）
              
                body = self._build_state_machine(node, handler_index)
            else:
                tail_stmts = all_stmts[handler_index + 1:]
                self._in_handler_tail = True
                try:
                    tail_py = self._flush_nodes(tail_stmts)
                finally:
                    self._in_handler_tail = False
                self._error_label_tail = tail_py
                try:
                    main_py = self._flush_nodes(all_stmts[:handler_index])
                finally:
                    self._error_label_tail = None
                body = main_py + tail_py
            if info.kind == "function":
                body.append(ast.Return(value=ast.Name(id=self.ret_var, ctx=ast.Load())))
        finally:
            self._declared_stack.pop()
            self.current = None
            self.ret_var = None
            self.byref_params = frozenset()
            self._with_stack = saved_with
            self._on_error_resume = saved_error
            self._in_class_method = saved_method
            self._method_locals = saved_locals
            self._var_classes = saved_var_classes

        # By default a procedure gets an error-handling frame: enter plus try/finally leave.
        # 中文：过程默认要一个错误处理帧：enter + try/finally leave。
        # Yet without On Error, Resume or Err the body cannot catch anything,
        # 中文：但过程体里没有 On Error / Resume / Err 时根本不会捕获错误，
        # and omitting the frame saves two calls per invocation, which call-heavy code badly wants.
        # 中文：省略帧可以省掉每次调用的两次函数调用（函数调用密集的代码很吃这个）。
        statements = prologue + (body or [ast.Pass()])
        if _needs_error_frame(node.body):
            leave_call = ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                   attr="leave", ctx=ast.Load()),
                args=[], keywords=[]))
            enter_call = ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                   attr="enter", ctx=ast.Load()),
                args=[], keywords=[]))
            guarded = ast.Try(body=statements, handlers=[], orelse=[],
                              finalbody=[leave_call])
            statements = [enter_call, guarded]

        arg_nodes = [ast.arg(arg="self")] if as_method else []
        defaults = []
        vararg = None
        for param in info.params:
            pname = canon(param.name)
            if param.paramarray:
                vararg = ast.arg(arg=pname)
                continue
            arg_nodes.append(ast.arg(arg=pname))
            if param.optional:
                if param.default is None:
                    default_value = ast.Name(id=EMPTY_NAME, ctx=ast.Load())
                else:
                    default_value = self.transform(param.default)
             # All method arguments are Refs; a procedure's ByRef parameters are Refs only on demand
             # 中文：方法的所有参数都是 Ref；过程的 ByRef 参数按需是 Ref
              
                if as_method or (param.byref and canon(param.name) in info.ref_params):
                    default_value = self._ref_call(default_value)
                defaults.append(default_value)

        arguments = ast.arguments(
            posonlyargs=[],
            args=arg_nodes,
            vararg=vararg,
            kwonlyargs=[],
            kw_defaults=[],
            kwarg=None,
            defaults=defaults,
        )

        if as_method:
            method_name = py_name or self._method_attrs.get(canon(info.name),
                                                            canon(info.name))
        else:
            method_name = info.py_name
        return ast.FunctionDef(
            name=method_name,
            args=arguments,
            body=statements,
            decorator_list=[],
            returns=None,
            type_comment=None,
        )

    # classes / 中文：类
    def _transform_class(self, node):
        fields = [m for m in node.body if isinstance(m, FieldDecl)]
        methods = [m for m in node.body if isinstance(m, (SubDef, FunctionDef))]
        props = [m for m in node.body if isinstance(m, PropertyDef)]

        default_member = None
        for member in node.body:
            if isinstance(member, AttributeDecl) and canon(member.key) == "vb_usermemid":
                if isinstance(member.value, Number) and member.value.val == 0:
                    default_member = canon(member.member)

        field_attrs = {}
        for field in fields:
            key = canon(field.name)
            public = getattr(field, "visibility", "public") == "public"
            field_attrs[key] = key if public else "_priv_" + key

        method_groups = {}
        for method in methods:
            method_groups.setdefault(canon(method.name), []).append(method)

        method_attrs = {}
        for method in methods:
            key = canon(method.name)
            if key in ("class_initialize", "class_terminate"):
                # the constructor and destructor keep their names and __init__ / __del__ call them
                # 中文：构造/析构保持原名，由 __init__/__del__ 调用
                method_attrs[key] = key
                continue
            public = getattr(method, "visibility", "public") == "public"
            method_attrs[key] = key if public else "_priv_" + key

        base_key = canon(node.base) if getattr(node, "base", None) else None
        if base_key is not None and base_key not in self.classes:
            raise OVPPObjectNotSupportedError(
                f"Base class '{node.base}' is not defined", line_number=node.line)
        if base_key == canon(node.name):
            raise OVPPObjectNotSupportedError(
                f"Class '{node.name}' cannot inherit from itself", line_number=node.line)

        own_attrs = ([ast.Constant(value=name) for name in sorted(field_attrs.values())]
                     + [ast.Constant(value=canon(p.name)) for p in props])
        if base_key:
            attrs_value = ast.BinOp(
                left=ast.Attribute(value=ast.Name(id="_ovpp_cls_" + base_key, ctx=ast.Load()),
                                   attr="_ovpp_attrs", ctx=ast.Load()),
                op=ast.Add(),
                right=ast.List(elts=own_attrs, ctx=ast.Load()))
        else:
            attrs_value = ast.List(elts=own_attrs, ctx=ast.Load())

        class_body = [
            ast.Assign(targets=[ast.Name(id="_vbs_name", ctx=ast.Store())],
                       value=ast.Constant(value=node.name)),
            ast.Assign(targets=[ast.Name(id="_ovpp_vbs_class", ctx=ast.Store())],
                       value=ast.Constant(value=True)),
            ast.Assign(targets=[ast.Name(id="_default_member", ctx=ast.Store())],
                       value=ast.Constant(value=default_member)),
            ast.Assign(targets=[ast.Name(id="_ovpp_attrs", ctx=ast.Store())],
                       value=attrs_value),
        ]

     # A Shared field is a class attribute shared by all instances, so it stays out of __init__
     # 中文：Shared 字段是类属性（所有实例共享），不进 __init__
      
        for field in fields:
            if not getattr(field, "shared", False):
                continue
            class_body.append(ast.Assign(
                targets=[ast.Name(id=field_attrs[canon(field.name)],
                                  ctx=ast.Store())],
                value=ast.Name(id=EMPTY_NAME, ctx=ast.Load())))

        init_stmts = []
        if base_key:
            init_stmts.append(ast.Expr(value=ast.Call(
                func=ast.Attribute(
                    value=ast.Call(func=ast.Name(id="super", ctx=ast.Load()),
                                   args=[], keywords=[]),
                    attr="__init__", ctx=ast.Load()),
                args=[], keywords=[])))
        for field in fields:
            if getattr(field, "shared", False):
                continue
            target = ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                   attr=field_attrs[canon(field.name)], ctx=ast.Store())
            init_stmts.append(ast.Assign(targets=[target],
                                         value=self._array_or_empty(field.dims)))

        initializer = next((m for m in methods if canon(m.name) == "class_initialize"), None)
        terminator = next((m for m in methods if canon(m.name) == "class_terminate"), None)

        if initializer is not None:
            init_stmts.append(ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                   attr="class_initialize", ctx=ast.Load()),
                args=[], keywords=[])))

     # Method arguments follow the Ref protocol, so Class_Initialize's arguments are boxed too
     # 中文：方法参数走 Ref 协议，所以 Class_Initialize 的实参也要包成 Ref
      
        init_call_args = ([ast.Starred(
            value=ast.ListComp(
                elt=ast.Call(
                    func=ast.Attribute(value=ast.Name(id=REF_MODULE, ctx=ast.Load()),
                                       attr="Ref", ctx=ast.Load()),
                    args=[ast.Name(id="_ovpp_a", ctx=ast.Load())], keywords=[]),
                generators=[ast.comprehension(
                    target=ast.Name(id="_ovpp_a", ctx=ast.Store()),
                    iter=ast.Name(id="args", ctx=ast.Load()),
                    ifs=[], is_async=0)]),
            ctx=ast.Load())] if initializer is not None else [])
        if initializer is not None:
            init_stmts[-1] = ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                   attr="class_initialize", ctx=ast.Load()),
                args=init_call_args, keywords=[]))

        class_body.append(ast.FunctionDef(
            name="__init__",
            args=ast.arguments(posonlyargs=[], args=[ast.arg(arg="self")],
                               vararg=ast.arg(arg="args"),
                               kwonlyargs=[], kw_defaults=[], kwarg=None, defaults=[]),
            body=init_stmts or [ast.Pass()],
            decorator_list=[], returns=None, type_comment=None))

        class_body.append(self._transform_setattr())

        saved_fields = self._class_fields
        saved_methods = self._class_methods
        resolved_field_attrs = dict(field_attrs)
        inherited_methods = set()
        for ancestor_key in self._class_chain(canon(node.name))[1:]:
            ancestor = self.classes.get(ancestor_key)
            if ancestor is None:
                continue
            for member in ancestor.body:
                if isinstance(member, FieldDecl) and getattr(member, 'visibility', 'public') == 'public':
                    resolved_field_attrs.setdefault(canon(member.name), canon(member.name))
            inherited_methods |= set(self.class_methods.get(ancestor_key, {}))

        inherited_props = set()
        for ancestor_key in self._class_chain(canon(node.name))[1:]:
            ancestor = self.classes.get(ancestor_key)
            if ancestor is None:
                continue
            for member in ancestor.body:
                if isinstance(member, PropertyDef):
                    inherited_props.add(canon(member.name))

        self._check_interfaces(node, method_groups)

        self._current_class = canon(node.name)
        self._shared_fields = frozenset(
            canon(f.name) for f in fields if getattr(f, "shared", False))
        saved_field_attrs = self._field_attrs
        saved_method_attrs = self._method_attrs
        saved_props = self._class_props
        self._class_props = (frozenset(canon(p.name) for p in props)
                             | frozenset(inherited_props))
        self._class_fields = frozenset(resolved_field_attrs)
        self._class_methods = (frozenset(method_attrs) | frozenset(inherited_methods)
                               | frozenset(canon(p.name) for p in props))
        self._field_attrs = resolved_field_attrs
        self._method_attrs = method_attrs
        try:
            for key, group in method_groups.items():
                if len(group) == 1:
                    only = group[0]
                    if getattr(only, "shared", False):
                        func = self._transform_proc(
                            only, as_method=True,
                            py_name=method_attrs.get(key, key))
                        # drop self
                        # 中文：去掉 self
                        func.args.args = func.args.args[1:]
                        func.decorator_list = [ast.Name(id="staticmethod",
                                                        ctx=ast.Load())]
                        class_body.append(func)
                        continue
                    class_body.append(self._transform_proc(only, as_method=True))
                    continue
                arities = []
                for overload in group:
                    arity = len(overload.params)
                    arities.append(arity)
                    class_body.append(self._transform_proc(
                        overload, as_method=True,
                        py_name=f"_ovpp_m_{key}_{arity}"))
                class_body.append(self._overload_dispatcher(
                    method_attrs[key], key, arities, group[0].name))
            self._emit_properties(class_body, props)
            if terminator is not None:
                class_body.append(self._transform_terminate(terminator))
        finally:
            self._class_fields = saved_fields
            self._class_methods = saved_methods
            self._field_attrs = saved_field_attrs
            self._method_attrs = saved_method_attrs
            self._class_props = saved_props

        bases = ([ast.Name(id="_ovpp_cls_" + base_key, ctx=ast.Load())]
                 if base_key else [])
        return ast.ClassDef(name="_ovpp_cls_" + canon(node.name),
                            bases=bases, keywords=[], body=class_body,
                            decorator_list=[])

    def _overload_dispatcher(self, py_name, vbs_key, arities, display_name):
        """Overloads inside a class: generate an *args method dispatching on the argument count.

        中文：类内重载：生成 *args 方法，按参数个数转发到具体实现。
        """
        body = []
        for arity in arities:
            body.append(ast.If(
                test=ast.Compare(
                    left=ast.Call(func=ast.Name(id="len", ctx=ast.Load()),
                                  args=[ast.Name(id="args", ctx=ast.Load())],
                                  keywords=[]),
                    ops=[ast.Eq()],
                    comparators=[ast.Constant(value=arity)]),
                body=[ast.Return(value=ast.Call(
                    func=ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                       attr=f"_ovpp_m_{vbs_key}_{arity}",
                                       ctx=ast.Load()),
                    args=[ast.Starred(value=ast.Name(id="args", ctx=ast.Load()),
                                      ctx=ast.Load())],
                    keywords=[]))],
                orelse=[]))

        body.append(ast.Expr(value=ast.Call(
            func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                               attr="arity_error", ctx=ast.Load()),
            args=[ast.Constant(value=display_name),
                  ast.Call(func=ast.Name(id="len", ctx=ast.Load()),
                           args=[ast.Name(id="args", ctx=ast.Load())],
                           keywords=[]),
                  ast.List(elts=[ast.Constant(value=n) for n in arities],
                           ctx=ast.Load())],
            keywords=[])))
        body.append(ast.Raise(exc=ast.Call(
            func=ast.Name(id="RuntimeError", ctx=ast.Load()),
            args=[ast.Constant(value="unreachable")], keywords=[]), cause=None))

        return ast.FunctionDef(
            name=py_name,
            args=ast.arguments(
                posonlyargs=[],
                args=[ast.arg(arg="self")],
                vararg=ast.arg(arg="args"),
                kwonlyargs=[], kw_defaults=[], kwarg=None, defaults=[]),
            body=body,
            decorator_list=[], returns=None, type_comment=None)

    def _transform_setattr(self):
        """Add a __setattr__ guard so only declared fields and properties can be written.

        中文：给类加一道 __setattr__ 检查：只允许写已声明的字段/属性。
        """
        guard = ast.If(
            test=ast.UnaryOp(
                op=ast.Not(),
                operand=ast.Compare(
                    left=ast.Name(id="name", ctx=ast.Load()),
                    ops=[ast.In()],
                    comparators=[ast.Attribute(
                        value=ast.Name(id="self", ctx=ast.Load()),
                        attr="_ovpp_attrs", ctx=ast.Load())])),
            body=[ast.Raise(exc=ast.Call(
                func=ast.Name(id="AttributeError", ctx=ast.Load()),
                args=[ast.JoinedStr(values=[
                    ast.Constant(value="object has no settable member '"),
                    ast.FormattedValue(value=ast.Name(id="name", ctx=ast.Load()),
                                       conversion=-1, format_spec=None),
                    ast.Constant(value="'")])],
                keywords=[]), cause=None)],
            orelse=[])
        return ast.FunctionDef(
            name="__setattr__",
            args=ast.arguments(posonlyargs=[],
                               args=[ast.arg(arg="self"), ast.arg(arg="name"),
                                     ast.arg(arg="value")],
                               vararg=None, kwonlyargs=[], kw_defaults=[],
                               kwarg=None, defaults=[]),
            body=[guard,
                  ast.Expr(value=ast.Call(
                      func=ast.Attribute(value=ast.Name(id="object", ctx=ast.Load()),
                                         attr="__setattr__", ctx=ast.Load()),
                      args=[ast.Name(id="self", ctx=ast.Load()),
                            ast.Name(id="name", ctx=ast.Load()),
                            ast.Name(id="value", ctx=ast.Load())],
                      keywords=[]))],
            decorator_list=[], returns=None, type_comment=None)

    def _check_interfaces(self, node, method_groups):
        """The Implements contract, checked at compile time: method name plus argument count.

        中文：Implements：编译期契约检查（方法名 + 参数个数）。"""
        for interface in getattr(node, "interfaces", []) or []:
            key = canon(interface)
            iface = self.classes.get(key)
            if iface is None:
                raise OVPPSyntaxError(
                    f"Interface '{interface}' is not defined",
                    line_number=getattr(node, "line", None))
            for member in iface.body:
                if not isinstance(member, (SubDef, FunctionDef)):
                    continue
                if getattr(member, "visibility", "public") != "public":
                    continue
                name = canon(member.name)
                arity = len(member.params)
                group = method_groups.get(name) or []
                inherited = [len(self.class_methods[key2][name])
                             for key2 in self._class_chain(canon(node.name))[1:]
                             if name in self.class_methods.get(key2, {})]
                if not any(len(item.params) == arity for item in group) \
                        and arity not in inherited:
                    raise OVPPSyntaxError(
                        f"Class '{node.name}' does not implement "
                        f"'{member.name}' required by interface '{interface}'",
                        line_number=getattr(node, "line", None))

    def _emit_properties(self, class_body, props):
        """Group by property name and emit Get / Let / Set.

        中文：按属性名分组生成 Get / Let / Set。

        * Get -> @property; Let -> @name.setter
        * Set becomes a separate _ovpp_pset_<name>, called for Set obj.P = x
        中文：Set -> 独立的 _ovpp_pset_<name>（Set obj.P = x 时调用）
        * without a Get, add a write-only getter that raises when read, as VBScript write-only properties do
        中文：没有 Get -> 补一个"只写"getter，读取时抛错（VBS 写属性语义）
        """
        groups = {}
        for prop in props:
            groups.setdefault(canon(prop.name), []).append(prop)

        for key, members in groups.items():
            if not any(getattr(item, "kind", "get") == "get" for item in members):
                class_body.append(self._write_only_getter(key))

            for item in members:
                entries = self._transform_property(item)
                if getattr(item, "kind", None) == "set":
                    for entry in entries:
                        if isinstance(entry, ast.FunctionDef):
                            entry.name = "_ovpp_pset_" + key
                            entry.decorator_list = []
                class_body.extend(entries)

    def _write_only_getter(self, key):
        """Placeholder getter for a write-only property, so @key.setter has something to attach to.

        中文：只写属性的占位 getter：让 @key.setter 有东西可挂。"""
        raise_stmt = ast.Expr(value=ast.Call(
            func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                               attr="write_only_error", ctx=ast.Load()),
            args=[ast.Constant(value=f"Property '{key}' is write-only")],
            keywords=[]))
        return ast.FunctionDef(
            name=key,
            args=ast.arguments(posonlyargs=[], args=[ast.arg(arg="self")],
                               vararg=None, kwonlyargs=[], kw_defaults=[],
                               kwarg=None, defaults=[]),
            body=[raise_stmt],
            decorator_list=[ast.Name(id="property", ctx=ast.Load())],
            returns=None, type_comment=None)

    def _transform_property(self, node):
        name = canon(node.name)
        is_get = node.kind == "get"

        saved = (self._with_stack, self._on_error_resume, self._in_class_method,
                 self.ret_var, self.byref_params)
        self._with_stack = []
        self._on_error_resume = False
        self._in_class_method = True
        self.byref_params = frozenset()
        self.ret_var = name if is_get else None

        declared = {"self"} | {canon(p.name) for p in node.params}
        declared |= self._local_declarations(node.body)
        if is_get:
            declared.add(name)

        prologue = []
        if is_get:
            prologue.append(ast.Assign(
                targets=[ast.Name(id=name, ctx=ast.Store())],
                value=ast.Name(id=EMPTY_NAME, ctx=ast.Load())))

        self._declared_stack.append(declared)
        saved_locals = self._method_locals
        self._method_locals = frozenset(declared)
        try:
            handler_index, all_stmts = self._prepare_error_labels(node.body)
            if handler_index is None:
                body = self._flush_nodes(node.body)
            else:
                tail_py = self._flush_nodes(all_stmts[handler_index + 1:])
                self._error_label_tail = tail_py
                try:
                    main_py = self._flush_nodes(all_stmts[:handler_index])
                finally:
                    self._error_label_tail = None
                body = main_py + tail_py
            if is_get:
                body.append(ast.Return(value=ast.Name(id=name, ctx=ast.Load())))
        finally:
            self._declared_stack.pop()
            self._method_locals = saved_locals
            (self._with_stack, self._on_error_resume, self._in_class_method,
             self.ret_var, self.byref_params) = saved

        args = [ast.arg(arg="self")] + [ast.arg(arg=canon(p.name)) for p in node.params]
        decorators = ([ast.Name(id="property", ctx=ast.Load())] if is_get
                      else [ast.Attribute(value=ast.Name(id=name, ctx=ast.Load()),
                                          attr="setter", ctx=ast.Load())])

        return [ast.FunctionDef(
            name=name,
            args=ast.arguments(posonlyargs=[], args=args, vararg=None, kwonlyargs=[],
                               kw_defaults=[], kwarg=None, defaults=[]),
            body=prologue + (body or [ast.Pass()]),
            decorator_list=decorators,
            returns=None,
            type_comment=None)]

    def _transform_terminate(self, node):
        """Class_Terminate becomes __del__, and any exception there is swallowed.

        中文：Class_Terminate → __del__（异常一律吞掉）。"""
        method = self._transform_proc(node, as_method=True)
        guarded = ast.Try(
            body=method.body,
            handlers=[ast.ExceptHandler(type=ast.Name(id="Exception", ctx=ast.Load()),
                                        name="_ovpp_exc", body=[ast.Pass()])],
            orelse=[], finalbody=[])
        return ast.FunctionDef(name="__del__", args=method.args, body=[guarded],
                               decorator_list=[], returns=None, type_comment=None)

    # names / 中文：名字
    def _name_load(self, key):
        if key in self.byref_params:
            return ast.Attribute(value=ast.Name(id=key, ctx=ast.Load()),
                                 attr="v", ctx=ast.Load())
        return ast.Name(id=key, ctx=ast.Load())

    def _name_store(self, key):
        if key in self.byref_params:
            return ast.Attribute(value=ast.Name(id=key, ctx=ast.Load()),
                                 attr="v", ctx=ast.Store())
        return ast.Name(id=key, ctx=ast.Store())

    def _read_name(self, name, node=None):
        key = canon(name)
        if key == "me" and self._in_class_method:
            return ast.Name(id="self", ctx=ast.Load())
        if self._in_class_method and key in self._class_props:
            return ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                 attr=key, ctx=ast.Load())
        shared = self._shared_target(key)
        if shared is not None:
            return ast.Attribute(value=shared,
                                 attr=self._field_attrs[key], ctx=ast.Load())
        if self._is_field(key):
            return ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                 attr=self._field_attrs[key], ctx=ast.Load())
        in_scope = (self._declared_stack and key in self._declared_stack[-1])             or key in self.module_vars
        if key in self.classes and not in_scope:
            return ast.Name(id="_ovpp_cls_" + key, ctx=ast.Load())
        if key == "ide":
            return ast.Attribute(value=ast.Name(id=IDE_MODULE, ctx=ast.Load()),
                                 attr="IDE", ctx=ast.Load())
        if key == "json":
            return ast.Attribute(value=ast.Name(id=JSON_MODULE, ctx=ast.Load()),
                                 attr="JSON", ctx=ast.Load())
        if key == "cop":
            if PYTHON_FEATURE not in self.features:
                raise OVPPError(
                    "Python interop needs an explicit import: put "
                    '<script src="python-embed" /> at the top of the script')
            return ast.Attribute(value=ast.Name(id=COP_MODULE, ctx=ast.Load()),
                                 attr="COP", ctx=ast.Load())
        if key == "err":
            return ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                 attr="ERR", ctx=ast.Load())
        if key in VBS_CONSTANTS:
            return ast.Constant(value=VBS_CONSTANTS[key])
        if key == "vbtrue":
            return self._call_bool("to_bool", [ast.Constant(value=True)])
        if key == "vbfalse":
            return self._call_bool("to_bool", [ast.Constant(value=False)])
        self._check_scope(key, name, node, writing=False)
        return self._name_load(key)

    def _write_name(self, name, node=None):
        key = canon(name)
        if key == "me":
            raise OVPPSyntaxError("'Me' is read-only: you cannot assign to it",
                                  line_number=getattr(node, "line", None))
        shared = self._shared_target(key)
        if shared is not None:
            return ast.Attribute(value=shared,
                                 attr=self._field_attrs[key], ctx=ast.Store())
        if self._is_field(key):
            return ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                 attr=self._field_attrs[key], ctx=ast.Store())
        self._check_scope(key, name, node, writing=True)
        return self._name_store(key)

    def _shared_target(self, key):
        """A Shared field reads and writes on the class object instead of the instance.

        中文：Shared 字段：读写落在类对象上，而不是实例上。"""
        if not (self._current_class and key in self._shared_fields):
            return None
        return ast.Name(id="_ovpp_cls_" + self._current_class, ctx=ast.Load())

    def _is_field(self, key):
        return (self._in_class_method
                and key in self._field_attrs
                and key not in self._method_locals)

    def _check_scope(self, key, original, node, writing):
        if not self.option_explicit or not self._declared_stack:
            return
        if key in SPECIAL_NAMES or key in ("me", "self"):
            return
        if key in self._declared_stack[-1] or key in self.module_declared:
            return
        if key in self._field_attrs:
            return
        where = "assigned to" if writing else "used"
        raise OVPPVariableNotDefinedError(
            f"Variable '{original}' is {where} before it is declared "
            f"(Option Explicit is on); add a Dim statement",
            line_number=getattr(node, "line", None))

    def _subscript_chain(self, base, index_args, store):
        node = base
        count = len(index_args)
        for index, arg in enumerate(index_args):
            last = index == count - 1
            node = ast.Subscript(
                value=node,
                slice=self.transform(arg),
                ctx=ast.Store() if (store and last) else ast.Load(),
            )
        return node

    # member access / 中文：成员访问
    def _member_base(self, base, node=None):
        if base is None:
            if not self._with_stack:
                raise OVPPSyntaxError(
                    "'.member' can only be used inside a With block",
                    line_number=getattr(node, "line", None))
            return self._with_stack[-1]
        return self.transform(base)

    def _transform_member(self, node):
        if getattr(node, "safe", False):
            return ast.Call(
                func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                                   attr="safe_member", ctx=ast.Load()),
                args=[self._member_base(node.base, node),
                      ast.Constant(value=self._member_attr(node.name))],
                keywords=[])
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                               attr="member_get", ctx=ast.Load()),
            args=[self._member_base(node.base, node),
                  ast.Constant(value=self._member_attr(node.name))],
            keywords=[])

    def _member_store(self, node):
        return ast.Attribute(value=self._member_base(node.base, node),
                             attr=self._member_attr(node.name), ctx=ast.Store())

    def _member_attr(self, name):
        key = canon(name)
        if self._in_class_method and key in self._method_attrs:
            return self._method_attrs[key]
        return MEMBER_ALIASES.get(key, key)

    def _let_value(self, value, node):
        """On a plain assignment (Let), an object on the right contributes its default member's value.

        中文：普通赋值（Let）时，如果右边是对象且有默认成员，取默认属性值。"""
        if getattr(node, "is_set", False):
            return value
        if not isinstance(node.expr, (Var, Member)):
            return value
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                               attr="default_value", ctx=ast.Load()),
            args=[value], keywords=[])

    def _member_call_expr(self, member, args, statement=False):
        """obj.name(args) follows the Ref protocol; whether it is a method or a subscript is decided at run time.

        中文：obj.name(args)：统一走 Ref 协议，运行期再决定方法还是数组下标。
        """
        base_expr = self._member_base(member.base, member)
        attr = self._member_attr(member.name)

        refs, setups, writebacks = [], [], []
        for arg in args:
            tmp = self._new_tmp("ref")
            setups.append(ast.Assign(
                targets=[ast.Name(id=tmp, ctx=ast.Store())],
                value=self._ref_call(self.transform(arg))))
            refs.append(ast.Name(id=tmp, ctx=ast.Load()))
            target = self._writeback_target(arg)
            if target is not None:
                writebacks.append(ast.Assign(
                    targets=[target],
                    value=ast.Attribute(value=ast.Name(id=tmp, ctx=ast.Load()),
                                        attr="v", ctx=ast.Load())))

        call = ast.Call(
            func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                               attr=("safe_call" if getattr(member, "safe", False)
                                     else "call_method"),
                               ctx=ast.Load()),
            args=[base_expr, ast.Constant(value=attr),
                  ast.List(elts=refs, ctx=ast.Load())],
            keywords=[])

        if statement:
            pre, post = self._take_pending()
            return self._single(pre + setups + post + [ast.Expr(value=call)] + writebacks)

        self._pending_pre.extend(setups)
        self._pending_post.extend(writebacks)
        return call

    # procedure calls / 中文：过程调用
    def _arity_ok(self, info, count):
        return info.accepts(count)

    def _resolve_proc(self, key, count, node):
        """Pick an overload by argument count, with a clear error when nothing matches.

        中文：按参数个数在重载里挑一个；找不到时给出清晰的错误。
        """
        overloads = self.procs.get(key) or []
        for info in overloads:
            if info.arity == count:
                return info
        for info in overloads:
            if self._arity_ok(info, count):
                return info
        if not overloads:
            return None
        arities = ", ".join(info.describe_arity() for info in overloads)
        raise OVPPArgumentError(
            f"'{overloads[0].name}' has no overload taking {count} argument(s) "
            f"(available: {arities})",
            line_number=getattr(node, "line", None))

    def _builtin_call(self, key, args):
        func = vb_builtins.REGISTRY[key]
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                               attr=func.__name__, ctx=ast.Load()),
            args=[self.transform(arg) for arg in args],
            keywords=[],
        )

    def _ref_call(self, value):
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=REF_MODULE, ctx=ast.Load()),
                               attr="Ref", ctx=ast.Load()),
            args=[value], keywords=[])

    def _call_with_refs(self, func_expr, params, args, node, ref_names=None):
        """Build a call that handles ByRef according to the signature; procedures and
        object methods share it.

        A `ref_names` of None means every ByRef parameter needs a Ref, which is what a
        method call does since method parameters are Refs anyway. When a set is passed,
        only the parameters it contains get a Ref.

        中文：通用的「按签名处理 ByRef」调用构造器，普通过程和对象方法共用。`ref_names`
        A ref_names of None means every ByRef parameter needs a Ref, which is a method call since method parameters are Refs already;
        中文：为 None 表示「所有 ByRef 参数都要 Ref」（方法调用，方法的形参本来就是 Ref）；
        pass a set and only the parameters it names get a Ref.
        中文：传入集合时只有命中的参数才建 Ref。
        """
        py_args, setups, writebacks = [], [], []
        for index, arg in enumerate(args):
            param = params[index] if index < len(params) else None
            wants_ref = (param is not None and getattr(param, "byref", False)
                         and (ref_names is None or canon(param.name) in ref_names))
            if wants_ref:
                tmp = self._new_tmp("ref")
                setups.append(ast.Assign(
                    targets=[ast.Name(id=tmp, ctx=ast.Store())],
                    value=self._ref_call(self.transform(arg))))
                py_args.append(ast.Name(id=tmp, ctx=ast.Load()))
                target = self._writeback_target(arg)
                if target is not None:
                    writebacks.append(ast.Assign(
                        targets=[target],
                        value=ast.Attribute(value=ast.Name(id=tmp, ctx=ast.Load()),
                                            attr="v", ctx=ast.Load())))
            else:
                py_args.append(self.transform(arg))

        call = ast.Call(func=func_expr, args=py_args, keywords=[])
        return call, setups, writebacks

    def _proc_call(self, info, args, node):
        return self._call_with_refs(
            ast.Name(id=info.py_name, ctx=ast.Load()), info.params, args, node,
            info.ref_params)

    # object methods / 中文：对象方法
    def _class_chain(self, key):
        chain = []
        seen = set()
        while key and key not in seen:
            seen.add(key)
            chain.append(key)
            node = self.classes.get(key)
            key = (canon(node.base)
                   if node is not None and getattr(node, "base", None) else None)
        return chain

    def _method_params(self, cls_key, method_key):
        for key in self._class_chain(cls_key):
            methods = self.class_methods.get(key, {})
            if method_key in methods:
                return methods[method_key]
        return None

    def _known_class(self, base):
        if isinstance(base, Var):
            return self._var_classes.get(canon(base.name))
        if isinstance(base, NewExpr):
            return canon(base.name)
        return None

    def _method_call_with_refs(self, member, args, node, statement):
        """ByRef for object methods: when the receiver's class can be inferred statically, build Refs from the signature and write back.

        中文：对象方法的 ByRef：接收者类可静态推断时，按签名建 Ref 并写回。
        """
        cls = self._known_class(member.base)
        if not cls:
            return None
        params = self._method_params(cls, canon(member.name))
        if not params or not any(getattr(p, "byref", False) for p in params):
            return None

        func_expr = ast.Attribute(value=self._member_base(member.base, member),
                                  attr=self._member_attr(member.name),
                                  ctx=ast.Load())
        call, setups, writebacks = self._call_with_refs(func_expr, params, args, node)

        if statement:
            pre, post = self._take_pending()
            return self._single(pre + setups + post + [ast.Expr(value=call)] + writebacks)

        self._pending_pre.extend(setups)
        self._pending_post.extend(writebacks)
        return call

    def _self_method_call(self, name, args, statement):
        """A class calling its own method: temporarily push self onto the With stack and reuse the same call path.

        中文：类内部调用自身方法：临时把 self 放到 With 栈上复用同一套调用逻辑。
        """
        self._with_stack.append(ast.Name(id="self", ctx=ast.Load()))
        try:
            return self._member_call_expr(Member(None, name), args, statement=statement)
        finally:
            self._with_stack.pop()

    def _track_class(self, name, expr):
        """Record a variable's static type so object-method ByRef calls can use it.

        中文：记录变量的静态类型，供对象方法 ByRef 使用。
        """
        key = canon(name)
        cls = None
        if isinstance(expr, NewExpr):
            cls = canon(expr.name)
        elif (isinstance(expr, Call) and isinstance(expr.func, Var)
              and canon(expr.func.name) in ("createobject", "getobject")
              and expr.args and isinstance(expr.args[0], String)):
            progid = expr.args[0].val.strip().lower()
            if progid in self.classes:
                cls = progid

        if cls and cls in self.classes:
            self._var_classes[key] = cls
        else:
            self._var_classes.pop(key, None)

    def _writeback_target(self, arg):
        """The write-back target of a ByRef argument, or None when it is not assignable.

        中文：ByRef 实参的写回目标；不是可赋值表达式时返回 None。
        """
        if isinstance(arg, Var):
            return self._write_name(arg.name, arg)
        if isinstance(arg, Call):
            func = arg.func
            if not isinstance(func, Var):
                return None
            key = canon(func.name)
            if key in self.procs or key in BUILTIN_NAMES:
                return None
            if self._in_class_method and key in self._class_methods:
                return None
            base = self._read_name(func.name, arg)
            return self._subscript_chain(base, arg.args, store=True)
        return None

    # statements / 中文：语句
    def _transform_call_stmt(self, node):
        callee = node.callee

        if isinstance(callee, Var) and is_module_placeholder(callee.name):
            # a multi-file declaration, handled by the compiler
            # 中文：多文件声明：由 compiler 处理
            return ast.Pass()

        if isinstance(callee, Var) and is_placeholder(callee.name):
            source = self.python_blocks.get(placeholder_index(callee.name), "")
            return ast.Expr(value=ast.Call(
                func=ast.Name(id="exec", ctx=ast.Load()),
                args=[ast.Constant(value=source),
                      self._globals_call(), self._globals_call()],
                keywords=[]))

        if isinstance(callee, Member):
            return self._member_call_expr(callee, node.args, statement=True)

        if not isinstance(callee, Var):
            raise OVPPSyntaxError("Only named procedures or object members can be called",
                                  line_number=node.line)

        key = canon(callee.name)

        if key in ("execute", "executeglobal"):
            return self._transform_execute(node)
        if key == "getobject":
            return ast.Expr(value=self._transform_create_object(node, "get_object"))

        info = self._resolve_proc(key, len(node.args), node) if key in self.procs else None

        if info is None:
            if self._in_class_method and key in self._class_methods:
                return self._self_method_call(key, node.args, statement=True)
            if key in BUILTIN_NAMES:
                return ast.Expr(value=self._builtin_call(key, node.args))
            hint = ""
            if key == "print":
                hint = " (OpenVBS++ uses 'Echo' to print a value)"
            raise OVPPProcedureNotDefinedError(
                f"Sub or Function '{callee.name}' is not defined{hint}",
                line_number=node.line)

        call, setups, writebacks = self._proc_call(info, node.args, node)
        pre, post = self._take_pending()
        return self._single(pre + setups + post + [ast.Expr(value=call)] + writebacks)

    def _transform_goto(self, node):
        """GoTo label: run the statements after the label, then leave the procedure.

        中文：GoTo label：执行标签之后的语句，然后离开当前过程。
        """
        key = canon(node.label)
        index = self._proc_labels.get(key)
        if index is None:
            raise OVPPSyntaxError(f"Label '{node.label}' is not defined in this procedure",
                                  line_number=getattr(node, "line", None))
        if key in self._goto_seen:
            raise OVPPSyntaxError(f"Recursive GoTo to '{node.label}' is not supported",
                                  line_number=getattr(node, "line", None))
        self._goto_seen.add(key)
        try:
            stmts = self._flush_nodes(self._proc_body[index + 1:])
        finally:
            self._goto_seen.discard(key)
        stmts.append(ast.Return(value=None))
        return self._single(stmts)

    def _prepare_error_labels(self, body):
        """Scan the procedure body: register labels and prepare the handler for On Error GoTo <label>.

        中文：扫描过程体：登记标签，并为 On Error GoTo <label> 准备处理段。
        """
        labels = {}
        handler_label = None
        for index, stmt in enumerate(body):
            if isinstance(stmt, LabelStmt):
                key = canon(stmt.name)
                if key in labels:
                    raise OVPPSyntaxError(
                        f"Label '{stmt.name}' is defined more than once",
                        line_number=getattr(stmt, "line", None))
                labels[key] = index
            elif isinstance(stmt, OnErrorGoToLabel):
                handler_label = canon(stmt.label)

        self._proc_labels = labels
        self._proc_body = body

        if handler_label is None:
            return None, body
        index = labels.get(handler_label)
        if index is None:
            raise OVPPSyntaxError(
                f"Label '{handler_label}' is not defined in this procedure",
                line_number=getattr(body[0], "line", None))
        return index, body

    def _transform_resume(self, node):
        """`Resume` retries this statement, `Resume Next` moves to the next one.

        In state-machine mode both compile to an assignment on _pc, so no break or continue
        has to escape a block and nested If / For inside a handler stays ordered.

        中文：`Resume`（重试本句）/ `Resume Next`（下一条）。在状态机模式下编译成
        an assignment to _pc, so nothing has to break or continue out and a nested If or For inside the handler stays ordered.
        中文：「改写 _pc」——不需要 break/continue 逃逸，所以处理段里嵌套 If/For 也不会乱。
        """
        if not self._state_machine:
            raise OVPPSyntaxError(
                "'Resume' is only allowed inside an 'On Error GoTo <label>' handler "
                "that is compiled as a state machine",
                line_number=getattr(node, "line", None))
        if node.target is None:
            target = ast.Name(id="_ovpp_fail", ctx=ast.Load())
        elif str(node.target).lower() == "next":
            target = ast.BinOp(
                left=ast.Name(id="_ovpp_fail", ctx=ast.Load()),
                op=ast.Add(), right=ast.Constant(value=1))
        else:
            raise OVPPSyntaxError(
                f"'Resume {node.target}' (label form) is not supported yet",
                line_number=getattr(node, "line", None))

        return self._single([
            ast.Assign(targets=[ast.Name(id="_ovpp_pc", ctx=ast.Store())],
                       value=target),
            ast.Continue(),
        ])

    def _contains_resume(self, nodes):
        """Recursively check for Resume, which may sit inside an If or a For.

        中文：递归检查是否有 Resume（可能在 If/For 里）。
        """
        for node in nodes:
            if isinstance(node, ResumeStmt):
                return True
            for attribute in ("body", "then_body", "else_body", "finally_body"):
                inner = getattr(node, attribute, None)
                if isinstance(inner, list) and self._contains_resume(inner):
                    return True
            for entry in getattr(node, "cases", []) or []:
                if isinstance(entry, (list, tuple)) and len(entry) == 2:
                    if self._contains_resume(entry[1] or []):
                        return True
        return False

    def _build_state_machine(self, node, handler_index):
        """Compile the body into a loop dispatching on a statement index.

        Every statement is an `if _pc == N` branch:
          * _pc defaults to N+1; on error it becomes the handler index and _ovpp_fail = N;
          * inside the handler, Resume sets _pc back to _ovpp_fail and Resume Next to +1;
          * the handler itself is not wrapped in try, so its errors still propagate.

        中文：把过程体编译成「语句下标分派」循环。每条语句是一个 `if _pc == N` 分支：
        _pc starts at N+1; on error it becomes the handler index and _ovpp_fail records N; inside the handler
        中文：默认先 `_pc = N+1`，出错时 `_pc = 处理段下标` 并记下 `_ovpp_fail = N`；处理段里
        Resume sets _pc back to _ovpp_fail and Resume Next sets it to N+1; the handler is not wrapped in try itself,
        中文：`Resume` 把 _pc 设回 _ovpp_fail，`Resume Next` 设成 +1；处理段自身不再套 try，
        so an error inside the handler propagates as usual.
        中文：所以处理段里的错误照常向上抛。
        """
        order = []
        handler_pos = None
        label_pos = {}
        for index, stmt in enumerate(node.body):
            if isinstance(stmt, LabelStmt):
                if index == handler_index:
                    handler_pos = len(order)
                label_pos[canon(stmt.name)] = len(order)
                continue
            order.append(stmt)

        total = len(order)
        if handler_pos is None:
            handler_pos = total

        saved = (self._state_machine, self._in_handler_tail,
                 self._error_label_tail, self._on_error_resume,
                 self._sm_labels)
        self._state_machine = True
        self._error_label_tail = None
        self._on_error_resume = False
        self._sm_labels = label_pos

        try:
            chain = [ast.Return(value=None)]
            for position in range(total - 1, -1, -1):
                self._in_handler_tail = position >= handler_pos
                inner = self._flush_nodes([order[position]])
                if position < handler_pos:
                    handler_body = [
                        ast.Expr(value=ast.Call(
                            func=ast.Attribute(
                                value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                attr="record", ctx=ast.Load()),
                            args=[ast.Name(id="_ovpp_exc", ctx=ast.Load())],
                            keywords=[])),
                        ast.Assign(
                            targets=[ast.Name(id="_ovpp_fail", ctx=ast.Store())],
                            value=ast.Constant(value=position)),
                        ast.Assign(
                            targets=[ast.Name(id="_ovpp_pc", ctx=ast.Store())],
                            value=ast.Constant(value=handler_pos)),
                    ]
                    inner = [ast.Try(
                        body=inner or [ast.Pass()],
                        handlers=[ast.ExceptHandler(
                            type=ast.Name(id="Exception", ctx=ast.Load()),
                            name="_ovpp_exc", body=handler_body)],
                        orelse=[], finalbody=[])]

                branch = [
                    ast.Assign(
                        targets=[ast.Name(id="_ovpp_pc", ctx=ast.Store())],
                        value=ast.Constant(value=position + 1)),
                ] + inner
                chain = [ast.If(
                    test=ast.Compare(
                        left=ast.Name(id="_ovpp_pc", ctx=ast.Load()),
                        ops=[ast.Eq()],
                        comparators=[ast.Constant(value=position)]),
                    body=branch, orelse=chain)]
        finally:
            (self._state_machine, self._in_handler_tail, self._error_label_tail,
             self._on_error_resume, self._sm_labels) = saved

        init = [
            ast.Assign(targets=[ast.Name(id="_ovpp_pc", ctx=ast.Store())],
                       value=ast.Constant(value=0)),
            ast.Assign(targets=[ast.Name(id="_ovpp_fail", ctx=ast.Store())],
                       value=ast.Constant(value=0)),
        ]
        return init + [ast.While(test=ast.Constant(value=True),
                                 body=[chain[0]], orelse=[])]

    # ------------------------------------------------------------------ Try
    def _handler_stmts(self, handler):
        var, _when, body = handler
        stmts = self._flush_nodes(body)
        if var:
            stmts.insert(0, ast.Assign(
                targets=[ast.Name(id=canon(var), ctx=ast.Store())],
                value=ast.Call(
                    func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                       attr="snapshot", ctx=ast.Load()),
                    args=[], keywords=[])))
        return stmts or [ast.Pass()]

    def _transform_try(self, node):
        """Try ... Catch e [When condition] ... Finally ... End Try

        It compiles to a push_handler() followed by an ordinary try/except/finally. An
        On Error Resume Next inside the block hands exceptions to this handler, see
        err.capture.

        中文：编译成 push_handler() 之后一个标准 try/except/finally；块内的
        An On Error Resume Next inside the block hands its exceptions to this handler; see err.capture.
        中文：On Error Resume Next 会把异常交给这里（见 err.capture）。
        """
     # Exceptions inside the block belong to its Catch, so the procedure-level On Error GoTo handler is masked for now
     # 中文：块内的异常归本块的 Catch 管：暂时屏蔽过程级的 On Error GoTo 处理段
      
        saved_tail = self._error_label_tail
        self._error_label_tail = None
        saved_resume = self._on_error_resume
        self._on_error_resume = False
        try:
            body = self._flush_nodes(node.body) or [ast.Pass()]
        finally:
            self._error_label_tail = saved_tail
            self._on_error_resume = saved_resume

        push = ast.Expr(value=ast.Call(
            func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                               attr="push_handler", ctx=ast.Load()),
            args=[], keywords=[]))
        pop = ast.Expr(value=ast.Call(
            func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                               attr="pop_handler", ctx=ast.Load()),
            args=[], keywords=[]))

        handlers = []
        if node.handlers:
            inner = [ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                   attr="record", ctx=ast.Load()),
                args=[ast.Name(id="_ovpp_exc", ctx=ast.Load())], keywords=[]))]

         # Bind the snapshot to the Catch variable first, so the When clause can use e too
         # 中文：先把快照绑定到 Catch 变量，这样 When 子句里也能用 e
          
            first_var = next((entry[0] for entry in node.handlers if entry[0]), None)
            if first_var:
                inner.append(ast.Assign(
                    targets=[ast.Name(id=canon(first_var), ctx=ast.Store())],
                    value=ast.Call(
                        func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                           attr="snapshot", ctx=ast.Load()),
                        args=[], keywords=[])))

            wildcard = None
            conditioned = []
            for entry in node.handlers:
                if entry[1] is None:
                    wildcard = entry
                else:
                    conditioned.append(entry)
            else_part = (self._handler_stmts(wildcard) if wildcard else [ast.Raise()])

            for entry in reversed(conditioned):
                else_part = [ast.If(test=self.transform(entry[1]),
                                    body=self._handler_stmts(entry),
                                    orelse=else_part)]
            inner.extend(else_part)

            handlers = [ast.ExceptHandler(
                type=ast.Name(id="Exception", ctx=ast.Load()),
                name="_ovpp_exc", body=inner)]

        finalbody = [pop] + self._flush_nodes(node.finally_body or [])
        return self._single([push, ast.Try(body=body, handlers=handlers,
                                          orelse=[], finalbody=finalbody)])

    def _transform_dim(self, node):
        stmts = []
        for name, dims, type_name in node.items:
            cls = canon(type_name) if type_name else None
            if cls and cls in self.classes and dims is None:
                self._var_classes[canon(name)] = cls
            else:
                self._var_classes.pop(canon(name), None)
            stmts.append(ast.Assign(targets=[self._write_name(name, node)],
                                    value=self._array_or_empty(dims)))
        return self._single(stmts)

    def _array_or_empty(self, dims):
        if dims is None:
            return ast.Name(id=EMPTY_NAME, ctx=ast.Load())
        if not dims:
            return ast.List(elts=[], ctx=ast.Load())
        sizes = ast.List(
            elts=[ast.BinOp(left=self.transform(d), op=ast.Add(),
                            right=ast.Constant(value=1)) for d in dims],
            ctx=ast.Load())
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=ARR_MODULE, ctx=ast.Load()),
                               attr="make", ctx=ast.Load()),
            args=[sizes], keywords=[])

    def _transform_redim(self, node):
        key = canon(node.name)
        sizes = ast.List(
            elts=[ast.BinOp(left=self.transform(d), op=ast.Add(),
                            right=ast.Constant(value=1)) for d in node.dims],
            ctx=ast.Load())

        if node.preserve:
            value = ast.Call(
                func=ast.Attribute(value=ast.Name(id=ARR_MODULE, ctx=ast.Load()),
                                   attr="resize", ctx=ast.Load()),
                args=[self._read_name(node.name, node), sizes], keywords=[])
        else:
            value = ast.Call(
                func=ast.Attribute(value=ast.Name(id=ARR_MODULE, ctx=ast.Load()),
                                   attr="make", ctx=ast.Load()),
                args=[sizes], keywords=[])

        return ast.Assign(targets=[self._write_name(node.name, node)], value=value)

    def _transform_erase(self, node):
        stmts = []
        for name in node.names:
            key = canon(name)
            stmts.append(ast.Assign(
                targets=[self._write_name(name, node)],
                value=ast.Call(
                    func=ast.Attribute(value=ast.Name(id=ARR_MODULE, ctx=ast.Load()),
                                       attr="erase", ctx=ast.Load()),
                    args=[self._read_name(name, node)], keywords=[])))
        return self._single(stmts)

    def _transform_with(self, node):
        value = self.transform(node.obj)
        pre, post = self._take_pending()
        tmp = self._new_tmp("with")

        self._with_stack.append(ast.Name(id=tmp, ctx=ast.Load()))
        try:
            body = self._stmts(node.body)
        finally:
            self._with_stack.pop()

        prologue = pre + [ast.Assign(targets=[ast.Name(id=tmp, ctx=ast.Store())],
                                     value=value)] + post
        return self._single(prologue + body)

    def _transform_if(self, node):
        test = self.transform(node.cond)
     # Take the condition's pre and post statements first, or the first statement of the then branch swallows them
     # 中文：先取走条件产生的 pre/post，否则会被 then 分支的第一条语句吃掉
      
        pre, post = self._take_pending()

        if_node = ast.If(
            test=test,
            body=self._stmts(node.then_body),
            orelse=self._stmts(node.else_body) if node.else_body else []
        )

        if not pre and not post:
            return if_node

     # The condition calls a ByRef procedure, so evaluate it into a temporary, write back, then branch
     # 中文：条件里有 ByRef 调用：先把条件算进临时变量并写回，再分支
      
        tmp = self._new_tmp("cond")
        return self._single(pre + [ast.Assign(
            targets=[ast.Name(id=tmp, ctx=ast.Store())], value=if_node.test,
        )] + post + [ast.If(test=ast.Name(id=tmp, ctx=ast.Load()),
                            body=if_node.body, orelse=if_node.orelse)])

    def _transform_while(self, node):
        test = self.transform(node.cond)
     # Take the condition's pre and post statements first, before the loop body eats them
     # 中文：先取走条件产生的 pre/post，避免被循环体第一条语句吃掉
      
        pre, post = self._take_pending()
        body = self._stmts(node.body)

        if not pre and not post:
            return ast.While(test=test, body=body, orelse=[])

        tmp = self._new_tmp("cond")
        check = pre + [ast.Assign(targets=[ast.Name(id=tmp, ctx=ast.Store())],
                                  value=test)] + post
        check.append(ast.If(
            test=ast.UnaryOp(op=ast.Not(), operand=ast.Name(id=tmp, ctx=ast.Load())),
            body=[ast.Break()], orelse=[]))
        return ast.While(test=ast.Constant(value=True), body=check + body, orelse=[])

    def _transform_do_loop(self, node):
        body = self._stmts(node.body)

     # Do ... Loop While/Until cond runs at least once and tests at the end
     # 中文：Do ... Loop While/Until cond：至少执行一次，在尾部判断
      
        if node.post_cond is not None:
            cond = self.transform(node.post_cond)
            if node.post_cond_type == "UNTIL":
                check = ast.If(test=cond, body=[ast.Break()], orelse=[])
            else:
                check = ast.If(test=ast.UnaryOp(op=ast.Not(), operand=cond),
                               body=[ast.Break()], orelse=[])
            return ast.While(test=ast.Constant(value=True),
                             body=body + [check], orelse=[])

        if node.pre_cond is None:
            return ast.While(test=ast.Constant(value=True), body=body, orelse=[])

        if node.pre_cond_type == "UNTIL":
            test = ast.UnaryOp(op=ast.Not(), operand=self.transform(node.pre_cond))
        else:
            test = self.transform(node.pre_cond)

        pre, post = self._take_pending()
        if not pre and not post:
            return ast.While(test=test, body=body, orelse=[])

        tmp = self._new_tmp("cond")
        check = pre + [ast.Assign(targets=[ast.Name(id=tmp, ctx=ast.Store())],
                                  value=test)] + post
        check.append(ast.If(
            test=ast.UnaryOp(op=ast.Not(), operand=ast.Name(id=tmp, ctx=ast.Load())),
            body=[ast.Break()], orelse=[]))
        return ast.While(test=ast.Constant(value=True), body=check + body, orelse=[])

    def _transform_for(self, node):
        """For i = a To b [Step s] ... Next"""
        end_tmp = self._new_tmp("for_end")
        step_tmp = self._new_tmp("for_step")
        key = canon(node.var)

     # Fast path: the step is a literal constant and the body does not touch the counter, so use range
     # 中文：快路径：步长是字面量常量、循环体不改计数器 —— 直接用 range，
      
     # and skip the per-iteration step comparison and increment.
     # 中文：省掉每次迭代的 Step 比较与自增。
      
        constant_step = _const_step(node)
        if (constant_step is not None and constant_step != 0
                and not _writes_name(node.body, key)):
            return self._transform_for_seq(node, key, constant_step)

        end_value = self.transform(node.end)
        step_value = (self.transform(node.step) if node.step is not None
                      else ast.Constant(value=1))
        start_value = self.transform(node.start)
        pre, post = self._take_pending()

        prologue = pre + [
            ast.Assign(targets=[ast.Name(id=end_tmp, ctx=ast.Store())],
                       value=end_value),
            ast.Assign(targets=[ast.Name(id=step_tmp, ctx=ast.Store())],
                       value=step_value),
            ast.Assign(targets=[self._name_store(key)],
                       value=start_value),
        ] + post

        test = ast.IfExp(
            test=ast.Compare(
                left=ast.Name(id=step_tmp, ctx=ast.Load()),
                ops=[ast.GtE()],
                comparators=[ast.Constant(value=0)]
            ),
            body=ast.Compare(
                left=self._name_load(key),
                ops=[ast.LtE()],
                comparators=[ast.Name(id=end_tmp, ctx=ast.Load())]
            ),
            orelse=ast.Compare(
                left=self._name_load(key),
                ops=[ast.GtE()],
                comparators=[ast.Name(id=end_tmp, ctx=ast.Load())]
            ),
        )

        body = self._stmts(node.body)
        body.append(
            ast.Assign(targets=[self._name_store(key)],
                       value=ast.BinOp(
                           left=self._name_load(key),
                           op=ast.Add(),
                           right=ast.Name(id=step_tmp, ctx=ast.Load())
                       ))
        )

        loop = ast.While(test=test, body=body, orelse=[])
        return ast.If(test=ast.Constant(value=True), body=prologue + [loop], orelse=[])

    def _transform_for_seq(self, node, key, step):
        """The For fast path: `for i in for_seq(a, b, step): ...`.

        中文：For 的快路径：`for i in for_seq(a, b, step): ...`。
        """
        seq_tmp = self._new_tmp("for_seq")

        start_value = self.transform(node.start)
        end_value = self.transform(node.end)
        pre, post = self._take_pending()

        prologue = pre + [
            ast.Assign(targets=[self._name_store(key)],
                       value=start_value),
            ast.Assign(
                targets=[ast.Name(id=seq_tmp, ctx=ast.Store())],
                value=ast.Call(
                    func=ast.Attribute(value=ast.Name(id=ARR_MODULE, ctx=ast.Load()),
                                       attr="for_seq", ctx=ast.Load()),
                    args=[self._name_load(key), end_value, ast.Constant(value=step)],
                    keywords=[])),
        ] + post

        body = self._stmts(node.body)
     # After a normal VBScript finish the counter holds the first out-of-range value; for...else runs only when no
     # 中文：VBS 正常结束后计数器停在"第一个越界值"；for...else 只在没有
      
     # Exit For happened, and `if seq` keeps an empty range from adding one step out of nowhere.
     # 中文：Exit For 时执行，`if seq` 保证空区间不会凭空加一次 step。
      
        tail = ast.If(
            test=ast.Name(id=seq_tmp, ctx=ast.Load()),
            body=[ast.Assign(
                targets=[self._name_store(key)],
                value=ast.BinOp(left=self._name_load(key), op=ast.Add(),
                                right=ast.Constant(value=step)))],
            orelse=[])

        loop = ast.For(
            target=self._name_store(key),
            iter=ast.Name(id=seq_tmp, ctx=ast.Load()),
            body=body,
            orelse=[tail],
            type_comment=None,
        )
        return ast.If(test=ast.Constant(value=True),
                      body=prologue + [loop], orelse=[])

    def _transform_for_each(self, node):
        key = canon(node.var)
        value = self.transform(node.iterable)
        pre, post = self._take_pending()

        if pre or post:
            tmp = self._new_tmp("seq")
            prologue = pre + [ast.Assign(targets=[ast.Name(id=tmp, ctx=ast.Store())],
                                         value=value)] + post
            iterable = ast.Name(id=tmp, ctx=ast.Load())
            return self._single(prologue + [self._make_for_each(node, key, iterable)])

        return self._make_for_each(node, key, value)

    def _make_for_each(self, node, key, iterable):
        return ast.For(
            target=self._name_store(key),
            iter=iterable,
            body=self._stmts(node.body),
            orelse=[],
            type_comment=None,
        )

    def _case_test(self, select_var, cond):
        sv = ast.Name(id=select_var, ctx=ast.Load())

        if cond.kind == "EXPR":
            return ast.Compare(left=sv, ops=[ast.Eq()],
                               comparators=[self.transform(cond.value)])

        if cond.kind == "TO":
            low, high = cond.value
            return ast.BoolOp(op=ast.And(), values=[
                ast.Compare(left=sv, ops=[ast.GtE()],
                            comparators=[self.transform(low)]),
                ast.Compare(left=sv, ops=[ast.LtE()],
                            comparators=[self.transform(high)]),
            ])

        if cond.kind == "IS":
            op_str, value = cond.value
            return ast.Compare(left=sv, ops=[cmp_map[op_str]],
                               comparators=[self.transform(value)])

        raise NotImplementedError(f"Unsupported CaseCond kind: {cond.kind}")

    def _transform_select_case(self, node):
        select_var = self._new_tmp("select")

        value = self.transform(node.expr)
        pre, post = self._take_pending()
        stmts = pre + [ast.Assign(
            targets=[ast.Name(id=select_var, ctx=ast.Store())],
            value=value
        )] + post

        root_if = None
        current_if = None

        for case_exprs, body in node.cases:
            tests = [self._case_test(select_var, ce) for ce in case_exprs]
            test = tests[0] if len(tests) == 1 else ast.BoolOp(op=ast.Or(), values=tests)

            if_node = ast.If(test=test, body=self._stmts(body), orelse=[])
            if root_if is None:
                root_if = if_node
                current_if = if_node
            else:
                current_if.orelse = [if_node]
                current_if = if_node

        if root_if is None:
            if node.else_body:
                stmts.extend(self._stmts(node.else_body))
            return ast.If(test=ast.Constant(value=True), body=stmts, orelse=[])

        if node.else_body:
            current_if.orelse = self._stmts(node.else_body)

        stmts.append(root_if)
        return ast.If(test=ast.Constant(value=True), body=stmts, orelse=[])

    # expressions / 中文：表达式
    def _call_bool(self, func_name, args):
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=BOOL_MODULE, ctx=ast.Load()),
                               attr=func_name, ctx=ast.Load()),
            args=args, keywords=[])

    def _transform_address_of(self, node):
        """AddressOf Foo becomes a passable, callable procedure value.

        中文：AddressOf Foo -> 一个可传递、可调用的过程值。
        """
        key = canon(node.name)
        overloads = self.procs.get(key)
        cls_attr = None
        if overloads:
            py_name = overloads[0].py_name
        elif self._in_class_method and key in self._class_methods:
            py_name = self._method_attrs.get(key, key)
            cls_attr = True
        else:
            raise OVPPProcedureNotDefinedError(
                f"Procedure '{node.name}' is not defined",
                line_number=getattr(node, "line", None))

        if cls_attr:
            target = ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                   attr=py_name, ctx=ast.Load())
            flags = []
        else:
            target = ast.Name(id=py_name, ctx=ast.Load())
            flags = [ast.Constant(value=bool(p.byref))
                     for p in overloads[0].params]
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                               attr="delegate", ctx=ast.Load()),
            args=[target, ast.Constant(value=node.name),
                  ast.List(elts=flags, ctx=ast.Load())], keywords=[])

    def _transform_binop(self, node):
        if node.op == "COALESCE":
            return ast.Call(
                func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                                   attr="coalesce", ctx=ast.Load()),
                args=[self.transform(node.left), self.transform(node.right)],
                keywords=[])
        left = self.transform(node.left)
        right = self.transform(node.right)

        if node.op in op_map:
            return ast.BinOp(left=left, op=op_map[node.op], right=right)

        if node.op == "AMP":
            return ast.Call(
                func=ast.Attribute(value=ast.Name(id=STR_MODULE, ctx=ast.Load()),
                                   attr="ovpp_str_concat", ctx=ast.Load()),
                args=[left, right], keywords=[])

        if node.op in bool_call_map:
            return self._call_bool(bool_call_map[node.op], [left, right])

        if node.op in builtin_call_map:
            return ast.Call(
                func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                                   attr=builtin_call_map[node.op], ctx=ast.Load()),
                args=[left, right], keywords=[])

        raise NotImplementedError(f"Unsupported BinOp type: {node.op}")

    def _transform_unaryop(self, node):
        operand = self.transform(node.operand)
        if node.op == "NOT":
            return self._call_bool(unary_call_map[node.op], [operand])
        return ast.UnaryOp(op=unary_call_map[node.op], operand=operand)

    def _transform_compare(self, node):
     # Comparison results become VBScript booleans: true is -1 and false is 0
     # 中文：比较结果统一成 VBS 风格的布尔（真 = -1，假 = 0）
      
        left = self.transform(node.left)
        right = self.transform(node.right)

        if self.option_compare_text and node.op != "IS":
            compare = ast.Call(
                func=ast.Attribute(value=ast.Name(id=STR_MODULE, ctx=ast.Load()),
                                   attr="vb_compare", ctx=ast.Load()),
                args=[ast.Constant(value=node.op), left, right],
                keywords=[])
            return self._call_bool("to_bool", [compare])

        cmp_node = ast.Compare(
            left=left,
            ops=[cmp_map[node.op]],
            comparators=[right])
        return self._call_bool("to_bool", [cmp_node])

    def _transform_new_with_args(self, node):
        """New Foo(a, b) passes its arguments to Class_Initialize.

        中文：New Foo(a, b)：把参数传给 Class_Initialize。
        """
        key = canon(node.name)
        if key not in self.classes:
            raise OVPPObjectNotSupportedError(
                f"Class '{node.name}' is not defined",
                line_number=getattr(node, "line", None))
        return ast.Call(func=ast.Name(id="_ovpp_cls_" + key, ctx=ast.Load()),
                        args=[self.transform(arg) for arg in node.args],
                        keywords=[])

    def _transform_create_object(self, node, factory="create_object"):
        """CreateObject / GetObject: a literal ProgID is resolved at compile time.

        中文：CreateObject / GetObject：字面量 ProgID 在编译期解析。
        """
        if not node.args:
            raise OVPPArgumentError("CreateObject expects a ProgID string",
                                    line_number=node.line)

        if len(node.args) > 2:
            raise OVPPArgumentError(
                "CreateObject expects at most two arguments: "
                "CreateObject(class, [servername])",
                line_number=getattr(node, "line", None))

        if len(node.args) == 2:
         # VBScript's second argument names a remote server for DCOM activation, which this local implementation does not support
         # 中文：VBS 的第二个参数是远程服务器名（DCOM 激活），本地实现不支持
          
            raise OVPPObjectNotSupportedError(
                "Remote activation is not supported: CreateObject(class, "
                "servername) cannot be implemented locally",
                line_number=getattr(node, "line", None))

        first = node.args[0]
        rest = []

        if isinstance(first, String):
            progid = first.val.strip().lower()
            if progid == "vbscript.regexp":
                return ast.Call(
                    func=ast.Attribute(value=ast.Name(id=REGEXP_MODULE, ctx=ast.Load()),
                                       attr="RegExp", ctx=ast.Load()),
                    args=rest, keywords=[])
            if progid == "scripting.dictionary":
                return ast.Call(
                    func=ast.Attribute(value=ast.Name(id=COLL_MODULE, ctx=ast.Load()),
                                       attr="Dictionary", ctx=ast.Load()),
                    args=rest, keywords=[])
            if progid in self.classes:
                return ast.Call(func=ast.Name(id="_ovpp_cls_" + progid, ctx=ast.Load()),
                                args=rest, keywords=[])

        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                               attr=factory, ctx=ast.Load()),
            args=[self.transform(first), ast.Name(id=CLASSES_NAME, ctx=ast.Load())] + rest,
            keywords=[])

    def _globals_call(self):
        return ast.Call(func=ast.Name(id="globals", ctx=ast.Load()),
                        args=[], keywords=[])

    def _transform_execute(self, node):
        if len(node.args) != 1:
            raise OVPPArgumentError("Execute expects exactly one string argument",
                                    line_number=node.line)
        return ast.Expr(value=ast.Call(
            func=ast.Attribute(value=ast.Name(id=DYNAMIC_MODULE, ctx=ast.Load()),
                               attr="execute", ctx=ast.Load()),
            args=[self.transform(node.args[0]), self._globals_call()], keywords=[]))

    def _transform_eval(self, node):
        if len(node.args) != 1:
            raise OVPPArgumentError("Eval expects exactly one string argument",
                                    line_number=node.line)
        return ast.Call(
            func=ast.Attribute(value=ast.Name(id=DYNAMIC_MODULE, ctx=ast.Load()),
                               attr="evaluate", ctx=ast.Load()),
            args=[self.transform(node.args[0]), self._globals_call()], keywords=[])

    def _transform_call(self, node):
        func = node.func

        if isinstance(func, Member):
            return self._member_call_expr(func, node.args)

        if not isinstance(func, Var):
            raise OVPPSyntaxError("Only named procedures or object members can be called",
                                  line_number=node.line)

        key = canon(func.name)

        if key in ("createobject", "getobject"):
            return self._transform_create_object(
                node, "get_object" if key == "getobject" else "create_object")

        if key == "eval":
            return self._transform_eval(node)

        info = self._resolve_proc(key, len(node.args), node) if key in self.procs else None

        if info is None:
            if self._in_class_method and key in self._class_methods:
                return self._self_method_call(key, node.args, statement=False)
            if key in BUILTIN_NAMES:
                return self._builtin_call(key, node.args)
         # Neither a procedure nor a built-in function, so treat it as the array subscript a(i)
         # 中文：不是过程也不是内置函数：当作数组下标 a(i)
          
            base = self._read_name(func.name, node)
            if not node.args:
                # () on a variable: VBScript calls the procedure reference held by
                # the variable, and an empty subscript list is meaningless anyway.
                # 中文：变量上的 ()：VBScript 视为调用变量持有的过程引用；
                # 中文：空的参数列表本来也没有下标含义。
                return ast.Call(
                    func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                                       attr="call_value", ctx=ast.Load()),
                    args=[base, ast.List(elts=[], ctx=ast.Load())], keywords=[])
            if key in self._delegates or key in self._param_names:
                args = [self.transform(arg) for arg in node.args]
                return ast.Call(
                    func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                                       attr="call_value", ctx=ast.Load()),
                    args=[base, ast.List(elts=args, ctx=ast.Load())], keywords=[])
            return self._subscript_chain(base, node.args, store=False)

        if info.kind == "sub":
            raise OVPPTypeError(
                f"Sub '{func.name}' cannot be used as a function",
                line_number=node.line)

        if info.ref_params:
         # A ByRef function called inside an expression boxes its arguments first,
         # 中文：ByRef 函数在表达式里调用：实参先装进 Ref 盒子，
          
         # and the enclosing statement uses _hoist() to insert the write-backs at the right place.
         # 中文：写回语句由外层语句用 _hoist() 插在正确位置。
          
            py_args = []
            for index, arg in enumerate(node.args):
                param = info.params[index] if index < len(info.params) else None
                if param is None or not param.byref or canon(param.name) not in info.ref_params:
                    py_args.append(self.transform(arg))
                    continue

                tmp = self._new_tmp("ref")
                self._pending_pre.append(ast.Assign(
                    targets=[ast.Name(id=tmp, ctx=ast.Store())],
                    value=self._ref_call(self.transform(arg))))
                py_args.append(ast.Name(id=tmp, ctx=ast.Load()))

                target = self._writeback_target(arg)
                if target is not None:
                    self._pending_post.append(ast.Assign(
                        targets=[target],
                        value=ast.Attribute(value=ast.Name(id=tmp, ctx=ast.Load()),
                                            attr="v", ctx=ast.Load())))

            return ast.Call(
                func=ast.Name(id=info.py_name, ctx=ast.Load()),
                args=py_args, keywords=[])

        return ast.Call(
            func=ast.Name(id=info.py_name, ctx=ast.Load()),
            args=[self.transform(arg) for arg in node.args],
            keywords=[])

    # dispatch / 中文：分发
    def _dispatch(self, node):
        if isinstance(node, BinOp):
            return self._transform_binop(node)

        if isinstance(node, UnaryOp):
            return self._transform_unaryop(node)

        if isinstance(node, Compare):
            return self._transform_compare(node)

        if isinstance(node, Call):
            return self._transform_call(node)

        if isinstance(node, Member):
            return self._transform_member(node)

        if isinstance(node, Assign):
            if isinstance(node.expr, AddressOf):
                self._delegates.add(canon(node.name))
            else:
                self._delegates.discard(canon(node.name))
            self._track_class(node.name, node.expr)
            target = self._write_name(node.name, node)
            value = self._let_value(self.transform(node.expr), node)
            stmts, value = self._hoist(value)
            if not stmts:
                return ast.Assign(targets=[target], value=value)
            stmts.append(ast.Assign(targets=[target], value=value))
            return self._single(stmts)

        if isinstance(node, MemberAssign):
            if getattr(node, "is_set", False) and not node.index_args:
                value = self.transform(node.expr)
                stmts, value = self._hoist(value)
                call = ast.Call(
                    func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                                       attr="set_member", ctx=ast.Load()),
                    args=[self._member_base(node.base, node),
                          ast.Constant(value=self._member_attr(node.name)),
                          value,
                          ast.Constant(value=True)], keywords=[])
                stmts.append(ast.Expr(value=call))
                return self._single(stmts)
            if node.index_args:
                value = self.transform(node.expr)
                stmts, value = self._hoist(value)
                call = ast.Call(
                    func=ast.Attribute(value=ast.Name(id=OBJ_MODULE, ctx=ast.Load()),
                                       attr="set_member_index", ctx=ast.Load()),
                    args=[self._member_base(node.base, node),
                          ast.Constant(value=self._member_attr(node.name)),
                          ast.List(elts=[self.transform(arg)
                                         for arg in node.index_args],
                                   ctx=ast.Load()),
                          value],
                    keywords=[])
                stmts.append(ast.Expr(value=call))
                return self._single(stmts)

            target = self._member_store(node)
            value = self._let_value(self.transform(node.expr), node)
            stmts, value = self._hoist(value)
            if not stmts:
                return ast.Assign(targets=[target], value=value)
            stmts.append(ast.Assign(targets=[target], value=value))
            return self._single(stmts)

        if isinstance(node, AssignIndex):
            base = self._read_name(node.name, node)
            target = self._subscript_chain(base, node.index_args, store=True)
            value = self.transform(node.expr)
            stmts, value = self._hoist(value)
            if not stmts:
                return ast.Assign(targets=[target], value=value)
            stmts.append(ast.Assign(targets=[target], value=value))
            return self._single(stmts)

        if isinstance(node, DimDecl):
            return self._transform_dim(node)

        if isinstance(node, ReDimDecl):
            return self._transform_redim(node)

        if isinstance(node, ConstDecl):
            target = self._write_name(node.name, node)
            value = self.transform(node.expr)
            stmts, value = self._hoist(value)
            if not stmts:
                return ast.Assign(targets=[target], value=value)
            stmts.append(ast.Assign(targets=[target], value=value))
            return self._single(stmts)

        if isinstance(node, EraseStmt):
            return self._transform_erase(node)

        if isinstance(node, OnErrorResumeNext):
            self._on_error_resume = True
            return ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                   attr="resume_next", ctx=ast.Load()),
                args=[], keywords=[]))

        if isinstance(node, OnErrorGoToZero):
            self._on_error_resume = False
            return ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id=ERR_MODULE, ctx=ast.Load()),
                                   attr="goto_zero", ctx=ast.Load()),
                args=[], keywords=[]))

        if isinstance(node, ScriptEnd):
            # `End` unwinds the whole script: it is a BaseException so neither
            # `On Error Resume Next` nor `Try` can swallow it.
            # 中文：`End` 结束整个脚本；它是 BaseException，On Error 与 Try 都接不住。
            return ast.Raise(exc=ast.Call(
                func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                                   attr="ScriptEnd", ctx=ast.Load()),
                args=[], keywords=[]), cause=None)

        if isinstance(node, StopStmt):
            # VBScript's `Stop` is a debugger breakpoint; there is no debugger here.
            # 中文：VBScript 的 `Stop` 是调试断点；这里没有调试器，所以什么都不做。
            return ast.Pass()

        if isinstance(node, CallStmt):
            return self._transform_call_stmt(node)

        if isinstance(node, OptionExplicit):
            return None

        if isinstance(node, OptionTyping):
            return None

        if isinstance(node, OptionCompare):
            return None

        if isinstance(node, LabelStmt):
            # a label emits no code
            # 中文：标签本身不生成代码
            return ast.Pass()

        if isinstance(node, GoToStmt):
            return self._transform_goto(node)

        if isinstance(node, ResumeStmt):
            return self._transform_resume(node)

        if isinstance(node, OnErrorGoToLabel):
            # implemented by the procedure-level handler
            # 中文：由过程级处理段实现
            return ast.Pass()

        if isinstance(node, TryStmt):
            return self._transform_try(node)

        if isinstance(node, AddressOf):
            return self._transform_address_of(node)

        if isinstance(node, ClassDef):
            return None

        if isinstance(node, NewExpr) and node.args:
            return self._transform_new_with_args(node)

        if isinstance(node, NewExpr) and canon(node.name) == "regexp":
            return ast.Call(
                func=ast.Attribute(value=ast.Name(id=REGEXP_MODULE, ctx=ast.Load()),
                                   attr="RegExp", ctx=ast.Load()),
                args=[], keywords=[])
        if isinstance(node, NewExpr):
            key = canon(node.name)
            if key not in self.classes:
                raise OVPPObjectNotSupportedError(
                    f"Class '{node.name}' is not defined", line_number=node.line)
            return ast.Call(func=ast.Name(id="_ovpp_cls_" + key, ctx=ast.Load()),
                            args=[], keywords=[])

        if isinstance(node, Print):
            value = self.transform(node.expr)
            stmts, value = self._hoist(value)
            call = ast.Expr(value=ast.Call(
                func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                                   attr="Echo", ctx=ast.Load()),
                args=[value], keywords=[]))
            if not stmts:
                return call
            stmts.append(call)
            return self._single(stmts)

        if isinstance(node, If):
            return self._transform_if(node)

        if isinstance(node, While):
            return self._transform_while(node)

        if isinstance(node, DoLoop):
            return self._transform_do_loop(node)

        if isinstance(node, For):
            return self._transform_for(node)

        if isinstance(node, ForEach):
            return self._transform_for_each(node)

        if isinstance(node, With):
            return self._transform_with(node)

        if isinstance(node, SelectCase):
            return self._transform_select_case(node)

        if isinstance(node, (ExitDo, ExitFor)):
            return ast.Break()

        if isinstance(node, Break):
            return ast.Break()

        if isinstance(node, ExitSub):
            return ast.Return(value=None)

        if isinstance(node, ExitFunction):
            if self.ret_var is None:
                raise OVPPError("'Exit Function' outside a Function", line_number=node.line)
            return ast.Return(value=ast.Name(id=self.ret_var, ctx=ast.Load()))

        if isinstance(node, Number):
            return ast.Constant(value=node.val)

        if isinstance(node, Float):
            return ast.Constant(value=node.val)

        if isinstance(node, DateLit):
            return ast.Call(
                func=ast.Attribute(value=ast.Name(id=BUILTIN_MODULE, ctx=ast.Load()),
                                   attr="date_literal", ctx=ast.Load()),
                args=[ast.Constant(value=node.val)], keywords=[])

        if isinstance(node, String):
            return ast.Constant(value=node.val)

        if isinstance(node, Boolean):
            return self._call_bool("to_bool", [ast.Constant(value=node.val)])

        if isinstance(node, EmptyValue):
            return ast.Name(id=EMPTY_NAME, ctx=ast.Load())

        if isinstance(node, NullValue):
            return ast.Name(id=NULL_NAME, ctx=ast.Load())

        if isinstance(node, NothingValue):
            return ast.Name(id=NOTHING_NAME, ctx=ast.Load())

        if isinstance(node, Var):
            return self._read_name(node.name, node)

        if isinstance(node, (SubDef, FunctionDef)):
            return None

        raise NotImplementedError(type(node))
