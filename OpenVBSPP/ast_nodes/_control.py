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


class Break(AST):
    pass
class If(AST):
    def __init__(self, cond, then_body, else_body=None):
        self.cond = cond
        self.then_body = then_body
        self.else_body = else_body
    def __repr__(self):
        return f"If({self.cond}, {self.then_body}, {self.else_body})"
class While(AST):
    def __init__(self, cond, body):
        self.cond = cond
        self.body = body
    def __repr__(self):
        return f"While({self.cond}, {self.body})"
class DoLoop(AST):
    def __init__(self, body, pre_cond=None, pre_cond_type=None,
                 post_cond=None, post_cond_type=None):
        self.body = body                    # list[AST]
        # An expression, or None.
        # 中文：表达式或 None。
        self.pre_cond = pre_cond
        self.pre_cond_type = pre_cond_type  # "WHILE" / "UNTIL" / None
        # The expression after Loop While / Loop Until.
        # 中文：Loop While/Until 之后的表达式。
        self.post_cond = post_cond
        self.post_cond_type = post_cond_type
    def __repr__(self):
        return (f"DoLoop({self.body}, {self.pre_cond}, {self.pre_cond_type}, "
                f"{self.post_cond}, {self.post_cond_type})")
class ExitDo(Break):
    pass
class ExitFor(Break):
    pass


class ExitProcedure(AST):
    """Exit Sub / Exit Function: leave a procedure early.

    中文：Exit Sub / Exit Function：提前从过程返回。
    """
    pass
class ExitSub(ExitProcedure):
    pass
class ExitFunction(ExitProcedure):
    pass


class With(AST):
    """With obj ... End With"""
    def __init__(self, obj, body):
        self.obj = obj
        self.body = body
    def __repr__(self):
        return f"With({self.obj}, {self.body})"
class ForEach(AST):
    """For Each x In collection ... Next"""
    def __init__(self, var_name, iterable, body):
        self.var = var_name
        self.iterable = iterable
        self.body = body
    def __repr__(self):
        return f"ForEach({self.var}, {self.iterable}, {self.body})"
class SelectCase(AST):
    def __init__(self, expr, cases, else_body):
        # An expression.
        # 中文：表达式。
        self.expr = expr
        self.cases = cases              # list of ( [CaseCond,...], [stmts...] )
        self.else_body = else_body      # list of stmts
    def __repr__(self):
        return f"SelectCase({self.expr}, {self.cases}, {self.else_body})"
class CaseCond(AST):
    def __init__(self, kind, value):
        self.kind = kind          #kind:"EXPR":value = expr  "TO":value = (low, high) "IS":value = (op, expr)
        self.value = value
    def __repr__(self):
        return f"CaseCond({self.kind}, {self.value})"
class For(AST):
    def __init__(self, var_name, start_expr, end_expr, step_expr, body):
        self.var = var_name
        self.start = start_expr
        self.end = end_expr
        self.step = step_expr
        self.body = body
    def __repr__(self):
        return f"For({self.var}, {self.start}, {self.end}, {self.step}, {self.body})"


class TryStmt(AST):
    """Try ... Catch e [When condition] ... Finally ... End Try (a superset extension).

    中文：Try ... Catch e [When 条件] ... Finally ... End Try（超集扩展）。
    """
    def __init__(self, body, handlers, finally_body):
        self.body = body                  # list[AST]
        self.handlers = handlers          # [(var_name|None, when_expr|None, body)]
        self.finally_body = finally_body  # list[AST] | None
    def __repr__(self):
        return f"TryStmt({len(self.body)} stmts, {len(self.handlers)} handlers)"


class LabelStmt(AST):
    """A line label, as in `handler:` — native VBScript, used with On Error GoTo.

    中文：行标签：handler: （VBS 原生，配合 On Error GoTo 使用）。
    """
    def __init__(self, name):
        self.name = name
    def __repr__(self): return f"LabelStmt({self.name})"


class GoToStmt(AST):
    """GoTo label (used on error-handling paths).

    中文：GoTo label（错误处理路径用）。
    """
    def __init__(self, label):
        self.label = label
    def __repr__(self): return f"GoToStmt({self.label})"
