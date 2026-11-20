'''
Copyright 2026 Marf Chen(MarfChen)

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


from OpenVBSPP.ast_nodes._base import AST


class Param(AST):
    """One procedure parameter. VBScript passes by reference (ByRef) by default.

    中文：过程参数。VBS 默认按引用传递（ByRef）。
    """
    def __init__(self, name, byref=True, is_array=False, type_name=None,
                 optional=False, default=None, paramarray=False):
        self.name = name
        # ParamArray is always by value, as VBScript requires.
        # 中文：ParamArray 必须按值（VBS 规定）。
        self.byref = False if paramarray else byref
        self.is_array = is_array
        # The type annotation used under Option Typing.
        # 中文：Option Typing 下的类型注解。
        self.type_name = type_name
        # Whether the parameter is Optional.
        # 中文：Optional 参数。
        self.optional = optional
        # The default-value expression of an Optional parameter.
        # 中文：Optional 的默认值表达式。
        self.default = default
        # ParamArray absorbs the remaining arguments.
        # 中文：ParamArray —— 吸收剩余实参。
        self.paramarray = paramarray
    @property
    def min_arity(self):
        return 0 if (self.optional or self.paramarray) else 1
    def __repr__(self):
        mode = "ByRef" if self.byref else "ByVal"
        extra = " Optional" if self.optional else ""
        extra += " ParamArray" if self.paramarray else ""
        return f"Param({mode}{extra} {self.name}: {self.type_name})"


class ProcedureDef(AST):
    """What Sub and Function share.

    中文：Sub / Function 的公共部分。
    """
    def __init__(self, name, params, body, return_type=None):
        self.name = name
        self.params = params    # list[Param]
        self.body = body        # list[AST]
        self.return_type = return_type
    @property
    def kind(self):
        raise NotImplementedError
    def __repr__(self):
        return f"{type(self).__name__}({self.name}, {self.params})"


class SubDef(ProcedureDef):
    """Sub name(params) ... End Sub"""
    @property
    def kind(self):
        return "sub"


class FunctionDef(ProcedureDef):
    """Function name(params) ... End Function — the return value is assigned to the name.

    中文：Function name(params) ... End Function（用函数名赋值作为返回值）
    """
    @property
    def kind(self):
        return "function"


class CallStmt(AST):
    """Call a procedure or method as a statement: Call Foo(1) / Foo 1 / .Method 1

    callee is an expression node (a Var or a Member); the Transformer decides whether
    it is a procedure, a built-in function or an object member.

    中文：把过程 / 方法当作语句调用 —— `Call Foo(1)`、`Foo 1`、`.Method 1`。callee 是一个
    The expression node (Var or Member); the transformer decides whether it is a procedure, a built-in function or an object member.
    中文：表达式节点（Var 或 Member），由 Transformer 判定它是过程、内置函数还是对象成员。
    """
    def __init__(self, callee, args):
        self.callee = callee
        self.args = args        # list[AST]
    def __repr__(self):
        return f"CallStmt({self.callee}, {self.args})"
