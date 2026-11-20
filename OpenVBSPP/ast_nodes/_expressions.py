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


class Number(AST):
    def __init__(self, val): self.val = int(val)
    def __repr__(self): return f"Number({self.val})"
class Float(AST):
    """A floating-point literal such as 3.5.

    中文：浮点数字面量，例如 3.5。
    """
    def __init__(self, val): self.val = float(val)
    def __repr__(self): return f"Float({self.val})"
class DateLit(AST):
    """A `#...#` date literal, kept as text and parsed when the script runs.

    中文：`#...#` 日期字面量 —— 原样保存文本，运行时再解析成日期值。
    """

    def __init__(self, val):
        self.val = val

    def __repr__(self):
        return f"DateLit({self.val!r})"


class String(AST):
    def __init__(self, val):
        raw = val
        if len(raw) >= 2 and raw.startswith('"') and raw.endswith('"'):
            raw = raw[1:-1]
        # In VBS a doubled "" stands for one literal double quote.
        # 中文：VBS 里用 "" 表示一个字面双引号。
        self.val = raw.replace('""', '"')
    def __repr__(self): return f'String("{self.val}")'
class Boolean(AST):
    """A True / False literal.

    中文：True / False 字面量。
    """
    def __init__(self, val): self.val = bool(val)
    def __repr__(self): return f"Boolean({self.val})"
class EmptyValue(AST):
    """An Empty literal.

    中文：Empty 字面量。
    """
    def __repr__(self): return "EmptyValue()"
class NullValue(AST):
    """A Null literal.

    中文：Null 字面量。
    """
    def __repr__(self): return "NullValue()"
class NothingValue(AST):
    """A Nothing literal.

    中文：Nothing 字面量。
    """
    def __repr__(self): return "NothingValue()"
class Var(AST):
    def __init__(self, name): self.name = name
    def __repr__(self): return f"Var({self.name})"
class Member(AST):
    """Member access a.b; a None base means .b inside a With block.

    中文：成员访问 a.b；base 为 None 表示 With 块里的 .b。
    """
    def __init__(self, base, name):
        self.base = base
        self.name = name
    def __repr__(self):
        return f"Member({self.base}, {self.name})"
class BinOp(AST):
    def __init__(self, left, op, right):
        self.left = left
        self.op = op    # string like '+', '-', etc.
        self.right = right
    def __repr__(self): return f"BinOp({self.left}, '{self.op}', {self.right})"
class UnaryOp(AST):
    def __init__(self, op, operand):
        self.op = op
        self.operand = operand
    def __repr__(self): return f"UnaryOp('{self.op}', {self.operand})"
class Compare(AST):
    def __init__(self, left, op, right):
        self.left = left
        self.op = op   # 'EQ', 'NE', 'LT', 'GT', 'LE', 'GE', 'IS'
        self.right = right
    def __repr__(self): return f"Compare({self.left}, '{self.op}', {self.right})"
class Call(AST):
    """A function or method call, or an array subscript; the Transformer decides which.

    func is an expression node: a Var (procedure, built-in function or array) or a
    Member (method).

    中文：函数 / 方法调用，或数组下标（由 Transformer 区分）。func 是一个表达式节点 ——
    func is an expression node: Var (procedure, built-in function or array) or Member
    (method).

    中文：func 是一个表达式节点：Var（过程 / 内置函数 / 数组）或 Member（方法）。
    """
    def __init__(self, func, args):
        self.func = func
        self.args = args    # list of AST nodes
    def __repr__(self):
        return f"Call({self.func}, {self.args})"


class AddressOf(AST):
    """AddressOf Foo turns a procedure into a passable value (a superset extension).

    中文：AddressOf Foo：把过程变成可传递的值（超集扩展）。
    """
    def __init__(self, name):
        self.name = name
    def __repr__(self):
        return f"AddressOf({self.name})"
