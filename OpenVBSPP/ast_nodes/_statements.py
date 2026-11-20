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


from OpenVBSPP.ast_nodes._base import AST


class Assign(AST):
    def __init__(self, name, expr): self.name = name; self.expr = expr
    def __repr__(self): return f"Assign({self.name}, {self.expr})"
class MemberAssign(AST):
    """a.b = expr; a base of None means .b = expr inside a With block.

    中文：a.b = expr；base 为 None 表示 With 块里的 .b = expr。

    A non-empty index_args means an array field element: a.b(i) = expr.

    中文：index_args 非空时表示数组字段的元素赋值：a.b(i) = expr。
    """
    def __init__(self, base, name, expr, index_args=None):
        self.base = base
        self.name = name
        self.expr = expr
        self.index_args = index_args
    def __repr__(self):
        return f"MemberAssign({self.base}, {self.name}, {self.expr}, {self.index_args})"
class AssignIndex(AST):
    """a(i, j) = expr assigns an array element.

    中文：给数组元素赋值。"""
    def __init__(self, name, index_args, expr):
        self.name = name
        self.index_args = index_args
        self.expr = expr
    def __repr__(self):
        return f"AssignIndex({self.name}[{self.index_args}], {self.expr})"
class Print(AST):
    def __init__(self, expr): self.expr = expr
    def __repr__(self): return f"Print({self.expr})"
class OptionExplicit(AST):
    """Option Explicit: every variable must be declared before use.

    中文：要求所有变量先声明后使用。"""
    def __repr__(self): return "OptionExplicit()"
class OptionCompare(AST):
    """Option Compare Text / Binary controls case sensitivity of string comparison.

    中文：控制字符串比较是否区分大小写。"""
    def __init__(self, mode):
        self.mode = mode
    def __repr__(self): return f"OptionCompare({self.mode})"
class OptionTyping(AST):
    """Option Typing enables type annotations and compile-time type checking.

    中文：启用类型注解与编译期静态类型检查。"""
    def __repr__(self): return "OptionTyping()"
class DimDecl(AST):
    """Dim a, b(10), c(); items is [(name, dims or None), ...].

    中文：Dim a, b(10), c()，items 是 [(name, dims 或 None), ...]。"""
    def __init__(self, items):
        self.items = items
    def __repr__(self):
        return f"DimDecl({self.items})"
class ReDimDecl(AST):
    """ReDim [Preserve] a(10)"""
    def __init__(self, name, dims, preserve=False):
        self.name = name
        self.dims = dims
        self.preserve = preserve
    def __repr__(self):
        return f"ReDimDecl({self.name}, {self.dims}, preserve={self.preserve})"
class ConstDecl(AST):
    """Const name = expr"""
    def __init__(self, name, expr):
        self.name = name
        self.expr = expr
    def __repr__(self):
        return f"ConstDecl({self.name}, {self.expr})"
class EraseStmt(AST):
    """Erase a, b clears arrays.

    中文：Erase a, b 清空数组。"""
    def __init__(self, names):
        self.names = names
    def __repr__(self):
        return f"EraseStmt({self.names})"
class OnErrorResumeNext(AST):
    def __repr__(self): return "OnErrorResumeNext()"
class OnErrorGoToLabel(AST):
    """On Error GoTo <label>"""
    def __init__(self, label):
        self.label = label
    def __repr__(self): return f"OnErrorGoToLabel({self.label})"

class OnErrorGoToZero(AST):
    def __repr__(self): return "OnErrorGoToZero()"


class ResumeStmt(AST):
    """Resume [Next] continues or retries inside an On Error GoTo handler.

    中文：在 On Error GoTo 的处理段里继续或重试。"""
    def __init__(self, target=None):
        # None retries this statement, "next" moves on / 中文：None 重试本句，"next" 下一条
        self.target = target
    def __repr__(self):
        return f"ResumeStmt({self.target!r})"


class ScriptEnd(AST):
    """A bare `End` terminates the script immediately.

    中文：单独一个 `End` 立即结束整个脚本（VBScript 的 End 语句）。"""
    def __repr__(self):
        return "ScriptEnd()"


class StopStmt(AST):
    """`Stop` is a debugger breakpoint in VBScript; we accept and ignore it.

    中文：`Stop` 在 VBScript 里是调试断点；这里接受该语句但不做任何事。"""
    def __repr__(self):
        return "StopStmt()"
