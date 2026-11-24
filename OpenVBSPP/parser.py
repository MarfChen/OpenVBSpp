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


from OpenVBSPP.lexer import lex
from OpenVBSPP.ast_nodes import (Program, Assign, MemberAssign, AssignIndex, Print,
                              Var, Member, Number, Float, DateLit, String, Boolean,
                              EmptyValue, NullValue, NothingValue,
                              Compare, BinOp, UnaryOp, Call, NewExpr, AddressOf,
                              While, If, DoLoop, ExitDo, ExitFor,
                              ExitSub, ExitFunction, With, ForEach, TryStmt,
                              LabelStmt, GoToStmt, OnErrorGoToLabel, ResumeStmt,
                              ScriptEnd, StopStmt,
                              SelectCase, CaseCond, For,
                              OptionExplicit, OptionTyping, OptionNative,
                              OptionCompare,
                              DimDecl, ReDimDecl, ConstDecl,
                              EraseStmt, OnErrorResumeNext, OnErrorGoToZero,
                              Param, SubDef, FunctionDef, CallStmt,
                              FieldDecl, PropertyDef, ClassDef, AttributeDecl)
from OpenVBSPP.exceptions import OVPPSyntaxError


# These keywords close a block and should never start a statement.
# 中文：这些关键字表示「当前语句块结束」，正常不应出现在语句起始位置。
BLOCK_END_WORDS = ("end", "elseif", "else", "wend", "loop", "next", "case")

# Statement separators: a newline or a colon.
# 中文：语句分隔符 —— 换行与冒号。
SEPARATORS = ("NEWLINE", "COLON")

CMP_OPS = ("EQ", "NE", "LT", "GT", "LE", "GE")

# Keyword-led statements: the generic assignment / call branch never second-guesses
# them once one of these keywords is seen.
# 中文：关键字起始的语句 —— 遇到它们时，通用的「赋值 / 过程调用」分支不会再往下猜。
KEYWORD_STARTERS = (
    "if", "select", "for", "do", "while", "echo", "exit", "call",
    "option", "dim", "redim", "const", "sub", "function", "private", "public",
    "with", "on", "erase", "let", "set", "class", "property", "attribute",
) + BLOCK_END_WORDS


class Parser:
    def __init__(self, tokens, filename=None, source=None):
        self.tokens = list(tokens)
        self.pos = 0
        self.filename = filename
        self.source_lines = source.splitlines() if source is not None else None
        # The enclosing procedure: 'sub' or 'function'.
        # 中文：当前所在过程 —— 'sub' / 'function'。
        self.proc_stack = []
        # The names of the enclosing classes.
        # 中文：当前所在类名。
        self.class_stack = []
        # True once a non-Option statement has been parsed: `Option` statements
        # have to come first. 中文：出现过非 Option 语句后为真 —— Option 必须写在前头。
        self.seen_code = False

    # ------------------------------------------------------------ infrastructure / 中文：基础设施
    def peek(self):
        if self.pos < len(self.tokens):
            return self.tokens[self.pos]
        # At the end of input return an EOF sentinel; its line is the last line, which
        # makes the error message point at something sensible.
        # 中文：没有 token 时返回 EOF 哨兵（行号取最后一行，便于报错）。
        line = self.tokens[-1][2] if self.tokens else 1
        return ("EOF", "", line, 1)

    def at_eof(self):
        return self.pos >= len(self.tokens)

    def eat(self, kind=None):
        tok = self.peek()
        if kind and tok[0] != kind:
            self.error(f"Expected {kind}, got {self.describe(tok)}", tok)
        self.pos += 1
        return tok

    def __repr__(self):
        return f"Parser(pos={self.pos}, tokens={self.tokens})"

    def skip_newlines(self):
        """Skip newlines only, as used inside expressions.

        中文：只跳过换行（表达式内部用）。
        """
        while self.pos < len(self.tokens) and self.tokens[self.pos][0] == "NEWLINE":
            self.pos += 1

    def skip_separators(self):
        """Skip statement separators: newlines and colons.

        中文：跳过语句分隔符：换行与冒号。
        """
        while self.pos < len(self.tokens) and self.tokens[self.pos][0] in SEPARATORS:
            self.pos += 1

    @staticmethod
    def describe(tok):
        kind, value = tok[0], tok[1]
        if kind == "EOF":
            return "end of file"
        if kind == "MISMATCH":
            return f"character {value!r}"
        if kind == "NEWLINE":
            return "end of line"
        return f"'{value}'"

    @staticmethod
    def is_word(tok, word):
        return tok[0] == "IDENT" and tok[1].lower() == word

    def is_block_end(self, tok):
        return tok[0] == "IDENT" and tok[1].lower() in BLOCK_END_WORDS

    def at_statement_end(self, tok=None):
        tok = self.peek() if tok is None else tok
        return tok[0] in ("NEWLINE", "COLON", "EOF") or self.is_block_end(tok)

    def error(self, message, tok=None):
        tok = self.peek() if tok is None else tok
        line = tok[2] if len(tok) > 2 else None
        col = tok[3] if len(tok) > 3 else None
        source_line = None
        if line and self.source_lines and 1 <= line <= len(self.source_lines):
            source_line = self.source_lines[line - 1].rstrip()
        raise OVPPSyntaxError(message, line_number=line, column=col, source_line=source_line)

    # ------------------------------------------------------------ statement blocks / 中文：语句块
    def parse_block(self, stop_words, block_name, missing_hint):
        """Parse a block of statements until a keyword in stop_words appears.

        Anything unexpected — end of input, a terminator that belongs to another block —
        is reported as a syntax error. Returning None would leave the caller spinning,
        which is what used to hang the compiler.

        中文：解析一个语句块，直到遇到 stop_words 中的关键字。任何意外（EOF、不属于本块
        terminator) is reported as a syntax error; nothing ever returns None and leaves the
        caller spinning. That is the lesson of an early hang.
        中文：的结束符）都上报为语法错误，绝不返回 None 让调用方空转，这是早期「编译器卡死」
        中文：的教训。
        """
        body = []
        while True:
            self.skip_separators()
            tok = self.peek()

            if tok[0] == "EOF":
                self.error(f"Missing '{missing_hint}': unexpected end of file "
                           f"while reading the {block_name} block", tok)

            if tok[0] == "IDENT" and tok[1].lower() in stop_words:
                # A bare `End` is the End statement, not the block's `End If`.
                # 中文：单独的 `End` 是 End 语句，而不是本块的 `End If`。
                if not (tok[1].lower() == "end" and self._bare_end_here()):
                    return body

            body.append(self.statement())

    def _bare_end_here(self):
        """True when the current `End` stands alone, i.e. it is the End statement.

        `End If` / `End Sub` / `End With` etc. have a keyword after `End`.

        中文：当前 `End` 是否是单独出现（即 End 语句）。`End If`、`End Sub` 这类
        后面还有关键字，不算。
        """
        index = self.pos + 1
        return (index >= len(self.tokens)
                or self.tokens[index][0] in ("NEWLINE", "COLON", "EOF"))

    def expect_keyword(self, word):
        tok = self.peek()
        if not self.is_word(tok, word):
            self.error(f"Expected '{word.title()}', got {self.describe(tok)}", tok)
        self.pos += 1
        return tok

    def end_block(self, word):
        """Consume 'End <word>', as in End If, End Select, End Sub or End With.

        中文：吃掉 'End <word>'，例如 End If / End Select / End Sub / End With。
        """
        tok = self.peek()
        if not self.is_word(tok, "end"):
            self.error(f"Expected 'End {word.title()}', got {self.describe(tok)}", tok)
        self.pos += 1
        after = self.peek()
        if not self.is_word(after, word):
            self.error(f"Expected '{word.title()}' after 'End', "
                       f"got {self.describe(after)}", after)
        self.pos += 1

    # ------------------------------------------------------------ entry / 中文：入口
    def parse(self):
        body = []
        while True:
            self.skip_separators()
            if self.at_eof():
                break
            body.append(self.statement())
        return Program(body)

    def statement(self):
        tok = self.peek()

        if tok[0] == "EOF":
            self.error("Unexpected end of file", tok)
        if tok[0] == "COLON":
            # An empty statement (run of colons) is swallowed.
            # 中文：空语句（连续冒号）—— 直接吃掉。
            self.eat("COLON")
            return self.statement()
        if tok[0] == "DOT":
            return self.with_member_statement()
        if tok[0] != "IDENT":
            self.error(f"Unexpected {self.describe(tok)} at start of statement", tok)

        word = tok[1].lower()

        # Option statements describe the whole script, so they have to come
        # before anything else — VBScript has the same rule for `Option
        # Explicit`. Without this, a later `Option Native Off` would silently
        # rewrite how the code *above* it was compiled.
        # 中文：Option 语句描述整个脚本，必须写在所有其它语句之前 —— VBS 对
        # `Option Explicit` 也是这个规则。否则后面一句 `Option Native Off`
        # 会悄悄改变它上面代码的编译方式。
        if word == "option":
            if self.seen_code:
                self.error(f"'Option {self._option_name()}' must appear before "
                           "any other statement", tok)
        else:
            self.seen_code = True

        if self._is_label():
            return self.label_statement()
        if word == "goto":
            self.eat("IDENT")
            target = self.peek()
            if target[0] != "IDENT":
                self.error(f"Expected a label after 'GoTo', got {self.describe(target)}",
                           target)
            self.eat("IDENT")
            return GoToStmt(target[1]).at(tok[2], tok[3])
        if word == "resume":
            self.eat("IDENT")
            target = None
            nxt = self.peek()
            if nxt[0] == "IDENT" and nxt[1].lower() == "next":
                self.eat("IDENT")
                target = "next"
            elif nxt[0] == "IDENT":
                self.eat("IDENT")
                target = nxt[1]
            return ResumeStmt(target).at(tok[2], tok[3])
        if word == "try" and self._next_is_newline():
            return self.try_statement()
        if word == "if":
            return self.if_statement()
        if word == "select":
            return self.select_case_statement()
        if word == "for":
            return self.for_statement()
        if word == "do":
            return self.do_loop()
        if word == "while":
            return self.while_statement()
        if word == "with":
            return self.with_statement()
        if word == "echo":
            return self.print_statement()
        if word == "exit":
            return self.exit_statement()
        if word == "on":
            return self.on_error_statement()
        if word == "erase":
            return self.erase_statement()
        if word == "option":
            return self.option_statement()
        if word == "dim":
            return self.dim_statement()
        if word == "redim":
            return self.redim_statement()
        if word == "const":
            return self.const_statement()
        if word == "sub":
            return self.procedure_statement("sub")
        if word == "function":
            return self.procedure_statement("function")
        if word == "call":
            return self.call_statement()
        if word == "class":
            return self.class_statement()
        if word == "property":
            if not self.class_stack:
                self.error("'Property' is only allowed inside a Class", tok)
            return self.property_statement()
        if word == "attribute":
            if not self.class_stack:
                self.error("'Attribute' is only allowed inside a Class", tok)
            return self.attribute_statement()
        if word in ("private", "public"):
            visibility = word
            self.eat("IDENT")
            node = self.statement()
            if node is not None:
                node.visibility = visibility
            return node
        if word == "let":
            # `Let x = ...` is the same as `x = ...`.
            # 中文：`Let x = ...` 等价于 `x = ...`。
            self.eat("IDENT")
            return self.identifier_statement()
        if word == "set":
            # Set obj = New Foo / Set obj = CreateObject(...) / Set obj = Nothing
            self.eat("IDENT")
            node = self.identifier_statement()
            node.is_set = True
            return node
        if word == "stop":
            self.eat("IDENT")
            return StopStmt().at(tok[2], tok[3])
        if word == "end":
            if self._bare_end_here():
                # A bare `End` ends the whole script (VBScript's End statement).
                # 中文：单独的 `End` 结束整个脚本（VBScript 的 End 语句）。
                self.eat("IDENT")
                return ScriptEnd().at(tok[2], tok[3])
            self.error("Unexpected 'End': no matching block is open here", tok)
        if word in BLOCK_END_WORDS:
            self.error(f"Unexpected '{tok[1]}': no matching block is open here", tok)

        return self.identifier_statement()

    # ------------------------------------------------------------ simple statements / 中文：简单语句
    def print_statement(self):
        tok = self.eat("IDENT")   # Echo
        expr = self.expr()
        return Print(expr).at(tok[2], tok[3])

    def exit_statement(self):
        tok = self.eat("IDENT")   # Exit
        target = self.peek()
        if target[0] != "IDENT":
            self.error(f"Expected 'Sub', 'Function', 'Do' or 'For' after 'Exit', "
                       f"got {self.describe(target)}", target)
        what = target[1].lower()
        self.pos += 1

        if what == "do":
            return ExitDo().at(tok[2], tok[3])
        if what == "for":
            return ExitFor().at(tok[2], tok[3])
        if what == "sub":
            if not self.proc_stack or self.proc_stack[-1] != "sub":
                self.error("'Exit Sub' is only allowed inside a Sub", target)
            return ExitSub().at(tok[2], tok[3])
        if what == "function":
            if not self.proc_stack or self.proc_stack[-1] != "function":
                self.error("'Exit Function' is only allowed inside a Function", target)
            return ExitFunction().at(tok[2], tok[3])

        self.error(f"'Exit {target[1]}' is not supported yet", target)

    def on_error_statement(self):
        tok = self.eat("IDENT")   # On
        self.expect_keyword("error")

        nxt = self.peek()
        if nxt[0] != "IDENT":
            self.error(f"Expected 'Resume Next' or 'GoTo 0' after 'On Error', "
                       f"got {self.describe(nxt)}", nxt)

        word = nxt[1].lower()
        if word == "resume":
            self.eat("IDENT")
            self.expect_keyword("next")
            return OnErrorResumeNext().at(tok[2], tok[3])

        if word == "goto":
            self.eat("IDENT")
            target = self.peek()
            if target[0] == "NUMBER" and target[1] == "0":
                self.eat("NUMBER")
                return OnErrorGoToZero().at(tok[2], tok[3])
            if target[0] == "IDENT":
                self.eat("IDENT")
                return OnErrorGoToLabel(target[1]).at(tok[2], tok[3])
            self.error("Expected '0' or a label after 'On Error GoTo', "
                       f"got {self.describe(target)}", target)

        self.error(f"Expected 'Resume Next' or 'GoTo 0' after 'On Error', "
                   f"got '{nxt[1]}'", nxt)

    def erase_statement(self):
        tok = self.eat("IDENT")   # Erase
        names = []
        while True:
            nm = self.peek()
            if nm[0] != "IDENT":
                self.error(f"Expected an array name after 'Erase', "
                           f"got {self.describe(nm)}", nm)
            self.eat("IDENT")
            names.append(nm[1])
            if self.peek()[0] == "COMMA":
                self.eat("COMMA")
                continue
            break

        return EraseStmt(names).at(tok[2], tok[3])

    def _option_name(self):
        """The option word after `Option`, for error messages.

        中文：取 `Option` 后面那个词，用于报错信息。
        """
        index = self.pos + 1
        if index < len(self.tokens) and self.tokens[index][0] == "IDENT":
            return self.tokens[index][1]
        return "?"

    def option_statement(self):
        tok = self.eat("IDENT")   # Option
        word_tok = self.peek()
        if word_tok[0] != "IDENT":
            self.error(f"Expected an option name after 'Option', "
                       f"got {self.describe(word_tok)}", word_tok)
        self.pos += 1
        name = word_tok[1].lower()

        if name == "explicit":
            if self.proc_stack:
                self.error("'Option Explicit' must appear at script level, "
                           "not inside a procedure", tok)
            return OptionExplicit().at(tok[2], tok[3])

        if name == "typing":
            if self.proc_stack:
                self.error("'Option Typing' must appear at script level, "
                           "not inside a procedure", tok)
            return OptionTyping().at(tok[2], tok[3])

        if name == "native":
            if self.proc_stack:
                self.error("'Option Native' must appear at script level, "
                           "not inside a procedure", tok)
            enabled = True
            mode_tok = self.peek()
            if mode_tok[0] == "IDENT" and mode_tok[1].lower() in ("on", "off"):
                self.eat("IDENT")
                enabled = mode_tok[1].lower() == "on"
            return OptionNative(enabled).at(tok[2], tok[3])

        if name == "compare":
            mode_tok = self.peek()
            if mode_tok[0] != "IDENT" or mode_tok[1].lower() not in ("text", "binary"):
                self.error(f"Expected 'Text' or 'Binary' after 'Option Compare', "
                           f"got {self.describe(mode_tok)}", mode_tok)
            self.eat("IDENT")
            return OptionCompare(mode_tok[1].lower()).at(tok[2], tok[3])

        if name == "cop":
            self.error("'Option COP' was removed in 0.9.20; import Python "
                       'interop with <script src="python-embed" />', tok)

        if name == "strict":
            self.error("'Option Strict' was renamed to 'Option Native' in 0.9.26; "
                       "write 'Option Native' instead (Option Strict On -> Option Native On)", tok)

        if name == "base":
            self.error(f"'Option {word_tok[1]}' is planned but not implemented yet", word_tok)

        self.error(f"Unknown option '{word_tok[1]}'", word_tok)

    def parse_type_annotation(self):
        """Parse `-> typename`, or None when there is no annotation.

        中文：解析 `-> 类型名`；没有注解时返回 None。
        """
        if self.peek()[0] != "ARROW":
            return None
        self.eat("ARROW")
        tok = self.peek()
        if tok[0] != "IDENT":
            self.error(f"Expected a type name after '->', "
                       f"got {self.describe(tok)}", tok)
        self.eat("IDENT")
        return tok[1]

    def dim_statement(self):
        tok = self.eat("IDENT")   # Dim
        items = []

        while True:
            name_tok = self.peek()
            if name_tok[0] != "IDENT":
                self.error(f"Expected a variable name after 'Dim', "
                           f"got {self.describe(name_tok)}", name_tok)
            self.eat("IDENT")

            dims = None
            if self.peek()[0] == "LPAREN":
                self.eat("LPAREN")
                if self.peek()[0] == "RPAREN":
                    # `Dim a()` declares a dynamic array.
                    # 中文：`Dim a()` 声明动态数组。
                    dims = []
                else:
                    dims = [self.expr()]
                    while self.peek()[0] == "COMMA":
                        self.eat("COMMA")
                        dims.append(self.expr())
                self.eat("RPAREN")

            type_name = self.parse_type_annotation()
            items.append((name_tok[1], dims, type_name))

            if self.peek()[0] == "COMMA":
                self.eat("COMMA")
                continue
            break

        return DimDecl(items).at(tok[2], tok[3])

    def redim_statement(self):
        tok = self.eat("IDENT")   # ReDim

        preserve = False
        if self.is_word(self.peek(), "preserve"):
            self.eat("IDENT")
            preserve = True

        name_tok = self.peek()
        if name_tok[0] != "IDENT":
            self.error(f"Expected an array name after 'ReDim', "
                       f"got {self.describe(name_tok)}", name_tok)
        self.eat("IDENT")

        before = self.peek()
        if before[0] != "LPAREN":
            self.error(f"Expected '(' after 'ReDim {name_tok[1]}', "
                       f"got {self.describe(before)}", before)
        self.eat("LPAREN")
        dims = [self.expr()]
        while self.peek()[0] == "COMMA":
            self.eat("COMMA")
            dims.append(self.expr())
        self.eat("RPAREN")

        return ReDimDecl(name_tok[1], dims, preserve).at(tok[2], tok[3])

    def const_statement(self):
        tok = self.eat("IDENT")   # Const
        name_tok = self.peek()
        if name_tok[0] != "IDENT":
            self.error(f"Expected a constant name after 'Const', "
                       f"got {self.describe(name_tok)}", name_tok)
        self.eat("IDENT")

        if self.peek()[0] != "ASSIGN":
            self.error(f"Expected '=' after 'Const {name_tok[1]}', "
                       f"got {self.describe(self.peek())}", self.peek())
        self.eat("ASSIGN")
        expr = self.expr()
        return ConstDecl(name_tok[1], expr).at(tok[2], tok[3])

    # ------------------------------------------------------------ procedure definitions / 中文：过程定义
    def procedure_statement(self, kind):
        tok = self.eat("IDENT")   # Sub / Function

        if self.proc_stack:
            self.error(f"'{kind.title()}' cannot be nested inside another procedure "
                       f"(OpenVBS++ only supports script-level procedures)", tok)

        name_tok = self.peek()
        if name_tok[0] != "IDENT":
            self.error(f"Expected a {kind.title()} name, "
                       f"got {self.describe(name_tok)}", name_tok)
        self.eat("IDENT")

        params = self.parse_param_list()
        return_type = self.parse_type_annotation()

        header = f"{kind.title()} ... End {kind.title()}"
        self.proc_stack.append(kind)
        try:
            body = self.parse_block(("end",), header, f"End {kind.title()}")
        finally:
            self.proc_stack.pop()

        self.end_block(kind)

        node_cls = SubDef if kind == "sub" else FunctionDef
        return node_cls(name_tok[1], params, body, return_type).at(tok[2], tok[3])

    def parse_param_list(self):
        params = []
        if self.peek()[0] != "LPAREN":
            return params

        self.eat("LPAREN")
        if self.peek()[0] == "RPAREN":
            self.eat("RPAREN")
            return params

        while True:
            # ByRef is the VBScript default.
            # 中文：VBS 默认 ByRef。
            byref = True
            optional = False
            paramarray = False
            head = self.peek()
            if self.is_word(head, "optional"):
                self.eat("IDENT")
                optional = True
            elif self.is_word(head, "paramarray"):
                self.eat("IDENT")
                paramarray = True

            head = self.peek()
            if self.is_word(head, "byval"):
                self.eat("IDENT")
                byref = False
            elif self.is_word(head, "byref"):
                self.eat("IDENT")
                byref = True

            name_tok = self.peek()
            if name_tok[0] != "IDENT":
                self.error(f"Expected a parameter name, "
                           f"got {self.describe(name_tok)}", name_tok)
            self.eat("IDENT")

            is_array = False
            if self.peek()[0] == "LPAREN":
                self.eat("LPAREN")
                close = self.peek()
                if close[0] != "RPAREN":
                    self.error("Array parameters must be written as 'name()'", close)
                self.eat("RPAREN")
                is_array = True

            if paramarray and not is_array:
                self.error("ParamArray must be declared as 'name()'", name_tok)

            type_name = self.parse_type_annotation()

            default = None
            if self.peek()[1] == "=":
                if not optional:
                    self.error("Only 'Optional' parameters can have a default value",
                               self.peek())
                if self.peek()[0] == "ASSIGN":
                    self.eat("ASSIGN")
                else:
                    # The lexer reports `=` as the EQ operator.
                    # 中文：词法器给的是 "=" 的 EQ 形式。
                    self.pos += 1
                default = self.expr()

            params.append(Param(name_tok[1], byref, is_array, type_name,
                                optional, default, paramarray)
                          .at(name_tok[2], name_tok[3]))

            if self.peek()[0] == "COMMA":
                self.eat("COMMA")
                continue
            break

        for index, param in enumerate(params):
            if param.paramarray and index != len(params) - 1:
                self.error("ParamArray must be the last parameter", head)
            if param.optional and index + 1 < len(params):
                for later in params[index + 1:]:
                    if not later.optional and not later.paramarray:
                        self.error(
                            "Optional parameters must come after all required ones",
                            head)
        tail = self.peek()
        if tail[0] != "RPAREN":
            self.error(f"Expected ')' or ',' in the parameter list, "
                       f"got {self.describe(tail)}", tail)
        self.eat("RPAREN")
        return params

    # ------------------------------------------------------------ classes and properties / 中文：类与属性
    def class_statement(self):
        tok = self.eat("IDENT")   # Class

        if self.proc_stack:
            self.error("'Class' cannot be declared inside a procedure", tok)
        if self.class_stack:
            self.error("Nested 'Class' declarations are not supported", tok)

        name_tok = self.peek()
        if name_tok[0] != "IDENT":
            self.error(f"Expected a class name, got {self.describe(name_tok)}", name_tok)
        self.eat("IDENT")

        interfaces = []
        base = None
        if self.is_word(self.peek(), "inherits"):
            self.eat("IDENT")
            base_tok = self.peek()
            if base_tok[0] != "IDENT":
                self.error(f"Expected a base class name after 'Inherits', "
                           f"got {self.describe(base_tok)}", base_tok)
            self.eat("IDENT")
            base = base_tok[1]

        if self.is_word(self.peek(), "implements"):
            self.eat("IDENT")
            while True:
                iface_tok = self.peek()
                if iface_tok[0] != "IDENT":
                    self.error("Expected an interface name after 'Implements'",
                               iface_tok)
                self.eat("IDENT")
                interfaces.append(iface_tok[1])
                if self.peek()[0] == "COMMA":
                    self.eat("COMMA")
                    continue
                break

        self.class_stack.append(name_tok[1])
        members = []
        try:
            while True:
                self.skip_separators()
                nxt = self.peek()
                if nxt[0] == "EOF":
                    self.error("Missing 'End Class': unexpected end of file", nxt)
                if self.is_word(nxt, "end"):
                    break
                members.extend(self.class_member())
        finally:
            self.class_stack.pop()

        self.end_block("class")
        node = ClassDef(name_tok[1], members, base).at(tok[2], tok[3])
        node.interfaces = interfaces
        return node

    def class_member(self):
        tok = self.peek()
        visibility = "public"
        if self.is_word(tok, "public") or self.is_word(tok, "private"):
            visibility = tok[1].lower()
            self.eat("IDENT")
            tok = self.peek()

        shared = False
        if self.is_word(tok, "shared"):
            self.eat("IDENT")
            shared = True
            tok = self.peek()

        if self.is_word(tok, "sub"):
            node = self.procedure_statement("sub")
            node.visibility = visibility
            node.shared = shared
            return [node]
        if self.is_word(tok, "function"):
            node = self.procedure_statement("function")
            node.visibility = visibility
            node.shared = shared
            return [node]
        if self.is_word(tok, "property"):
            node = self.property_statement()
            node.visibility = visibility
            return [node]
        if self.is_word(tok, "attribute"):
            return [self.attribute_statement()]
        if self.is_word(tok, "dim"):
            self.eat("IDENT")

        fields = self.field_decl(visibility)
        for field in fields:
            field.shared = shared
        return fields

    def field_decl(self, visibility):
        nodes = []
        while True:
            name_tok = self.peek()
            if name_tok[0] != "IDENT":
                self.error(f"Expected a field name, got {self.describe(name_tok)}", name_tok)
            self.eat("IDENT")

            dims = None
            if self.peek()[0] == "LPAREN":
                self.eat("LPAREN")
                if self.peek()[0] == "RPAREN":
                    dims = []
                else:
                    dims = [self.expr()]
                    while self.peek()[0] == "COMMA":
                        self.eat("COMMA")
                        dims.append(self.expr())
                self.eat("RPAREN")

            type_name = self.parse_type_annotation()
            nodes.append(FieldDecl(name_tok[1], type_name, visibility, dims)
                         .at(name_tok[2], name_tok[3]))

            if self.peek()[0] == "COMMA":
                self.eat("COMMA")
                continue
            break

        return nodes

    def property_statement(self):
        tok = self.eat("IDENT")   # Property

        kind_tok = self.peek()
        if kind_tok[0] != "IDENT" or kind_tok[1].lower() not in ("get", "let", "set"):
            self.error(f"Expected 'Get', 'Let' or 'Set' after 'Property', "
                       f"got {self.describe(kind_tok)}", kind_tok)
        kind = kind_tok[1].lower()
        self.eat("IDENT")

        name_tok = self.peek()
        if name_tok[0] != "IDENT":
            self.error(f"Expected a property name, got {self.describe(name_tok)}", name_tok)
        self.eat("IDENT")

        params = self.parse_param_list()
        type_name = self.parse_type_annotation() if kind == "get" else None

        header = f"Property {kind.title()} ... End Property"
        self.proc_stack.append("property")
        try:
            body = self.parse_block(("end",), header, "End Property")
        finally:
            self.proc_stack.pop()

        self.end_block("property")
        return PropertyDef(kind, name_tok[1], params, body, type_name).at(tok[2], tok[3])

    def attribute_statement(self):
        tok = self.eat("IDENT")   # Attribute

        member_tok = self.peek()
        if member_tok[0] != "IDENT":
            self.error(f"Expected a member name after 'Attribute', "
                       f"got {self.describe(member_tok)}", member_tok)
        self.eat("IDENT")

        if self.peek()[0] != "DOT":
            self.error(f"Expected '.' after 'Attribute {member_tok[1]}', "
                       f"got {self.describe(self.peek())}", self.peek())
        self.eat("DOT")

        key_tok = self.peek()
        if key_tok[0] != "IDENT":
            self.error(f"Expected an attribute name after '.', "
                       f"got {self.describe(key_tok)}", key_tok)
        self.eat("IDENT")

        if self.peek()[0] != "ASSIGN":
            self.error(f"Expected '=' in 'Attribute' statement, "
                       f"got {self.describe(self.peek())}", self.peek())
        self.eat("ASSIGN")
        value = self.expr()

        return AttributeDecl(member_tok[1], key_tok[1], value).at(tok[2], tok[3])

    # ------------------------------------------------------------ calls and assignments / 中文：调用 / 赋值
    def call_statement(self):
        tok = self.eat("IDENT")   # Call
        if self.peek()[0] == "DOT":
            callee = self.with_member_target()
        else:
            name_tok = self.peek()
            if name_tok[0] != "IDENT":
                self.error(f"Expected a procedure name after 'Call', "
                           f"got {self.describe(name_tok)}", name_tok)
            self.eat("IDENT")
            callee = Var(name_tok[1]).at(name_tok[2], name_tok[3])
            callee = self.member_chain(callee)

        while self.peek()[0] == "LPAREN":
            args = self.parse_args()
            if self.peek()[0] not in ("DOT", "QDOT"):
                return CallStmt(callee, args).at(tok[2], tok[3])
            callee = Call(callee, args).at(tok[2], tok[3])
            callee = self.member_chain(callee)

        args = self.parse_bare_args()
        return CallStmt(callee, args).at(tok[2], tok[3])

    def with_member_statement(self):
        tok = self.peek()
        callee = self.with_member_target()
        return self.finish_call_or_assign(callee, tok)

    def with_member_target(self):
        dot = self.eat("DOT")
        nm = self.peek()
        if nm[0] != "IDENT":
            self.error(f"Expected a member name after '.', "
                       f"got {self.describe(nm)}", nm)
        self.eat("IDENT")
        node = Member(None, nm[1]).at(dot[2], dot[3])
        return self.member_chain(node)

    _COMPOUND_OPS = {"+=": "PLUS", "-=": "MINUS", "*=": "MUL",
                     "/=": "DIV", "&=": "AMP"}

    def identifier_statement(self):
        name_tok = self.peek()
        if name_tok[0] != "IDENT":
            self.error(f"Expected a variable or procedure name, "
                       f"got {self.describe(name_tok)}", name_tok)
        self.eat("IDENT")
        node = Var(name_tok[1]).at(name_tok[2], name_tok[3])
        node = self.member_chain(node)
        if self.peek()[0] == "COMPOUND":
            op = self.eat("COMPOUND")[1]
            value = self.expr()
            binary = BinOp(node, self._COMPOUND_OPS[op], value)
            target = Assign(name_tok[1], binary) if isinstance(node, Var) \
                else MemberAssign(node.base, node.name, binary)
            return target.at(name_tok[2], name_tok[3])
        return self.finish_call_or_assign(node, name_tok)

    def _interpolate(self, tok):
        """The interpolated string $"a{x}b" becomes "a" & x & "b", expanded at compile time.

        中文：$"a{x}b" -> "a" & x & "b"（编译期展开成拼接）
        """
        raw = tok[1]
        text = raw[2:-1].replace('""', '"')
        parts = []
        buffer = ""
        index = 0
        while index < len(text):
            char = text[index]
            if char == "{":
                depth = 1
                end = index + 1
                while end < len(text) and depth:
                    if text[end] == "{":
                        depth += 1
                    elif text[end] == "}":
                        depth -= 1
                    end += 1
                if depth == 0:
                    if buffer:
                        parts.append(String(buffer))
                        buffer = ""
                    inner = text[index + 1:end - 1]
                    sub = Parser(list(lex(inner)), filename=self.filename,
                                 source=inner)
                    parts.append(sub.expr())
                    index = end
                    continue
            buffer += char
            index += 1
        if buffer:
            parts.append(String(buffer))
        if not parts:
            return String("").at(tok[2], tok[3])
        node = parts[0]
        for part in parts[1:]:
            node = BinOp(node, "AMP", part).at(tok[2], tok[3])
        return node

    def member_chain(self, node):
        while self.peek()[0] in ("DOT", "QDOT"):
            dot = self.eat(self.peek()[0])
            nm = self.peek()
            if nm[0] != "IDENT":
                self.error(f"Expected a member name after '.', "
                           f"got {self.describe(nm)}", nm)
            self.eat("IDENT")
            member = Member(node, nm[1]).at(dot[2], dot[3])
            member.safe = dot[0] == "QDOT"
            node = member
        return node

    def finish_call_or_assign(self, target, tok):
        nxt = self.peek()

        if nxt[0] == "ASSIGN":
            self.eat("ASSIGN")
            expr = self.expr()
            return self.make_assignment(target, expr, tok)

        if nxt[0] == "LPAREN":
            args = self.parse_args()

            if self.peek()[0] in ("DOT", "QDOT"):
                target = Call(target, args).at(tok[2], tok[3])
                target = self.member_chain(target)
                return self.finish_call_or_assign(target, tok)

            if self.peek()[0] == "ASSIGN":
                self.eat("ASSIGN")
                expr = self.expr()
                if isinstance(target, Var):
                    return AssignIndex(target.name, args, expr).at(tok[2], tok[3])
                if isinstance(target, Member):
                    return MemberAssign(target.base, target.name, expr,
                                        args).at(tok[2], tok[3])
                self.error("Invalid assignment target", tok)
            return CallStmt(target, args).at(tok[2], tok[3])

        if self.at_statement_end(nxt):
            return CallStmt(target, []).at(tok[2], tok[3])

        args = []
        if self._argument_present():
            args.append(self.expr())
        else:
            args.append(EmptyValue().at(nxt[2], nxt[3]))
        while self.peek()[0] == "COMMA":
            comma = self.eat("COMMA")
            if self._argument_present():
                args.append(self.expr())
            else:
                args.append(EmptyValue().at(comma[2], comma[3]))
        return CallStmt(target, args).at(tok[2], tok[3])

    def make_assignment(self, target, expr, tok):
        if isinstance(target, Var):
            return Assign(target.name, expr).at(tok[2], tok[3])
        if isinstance(target, Member):
            return MemberAssign(target.base, target.name, expr).at(tok[2], tok[3])
        self.error("Invalid assignment target", tok)

    def _argument_present(self):
        """Whether the next token starts an argument instead of an omitted slot.

        `GetObject(, "WScript.Shell")` leaves the first argument out; VBScript writes
        that as a bare comma, so an empty slot becomes `EmptyValue`.
        中文：下一个 token 是实参的开头，还是被省略的实参槽位。
        `GetObject(, "WScript.Shell")` 省略了第一个实参，VBS 里写作一个逗号，
        因此空槽位用 `EmptyValue` 表示。
        """
        return self.peek()[0] not in ("COMMA", "RPAREN", "NEWLINE", "EOF")

    def parse_args(self):
        """Parse a parenthesised argument list.

        中文：解析括号内的实参列表。
        """
        self.eat("LPAREN")
        args = []
        if self.peek()[0] == "RPAREN":
            self.eat("RPAREN")
            return args

        self.skip_newlines()
        if self._argument_present():
            args.append(self.expr())
        else:
            tok = self.peek()
            args.append(EmptyValue().at(tok[2], tok[3]))
        while self.peek()[0] == "COMMA":
            comma = self.eat("COMMA")
            self.skip_newlines()
            if self._argument_present():
                args.append(self.expr())
            else:
                args.append(EmptyValue().at(comma[2], comma[3]))
        tail = self.peek()
        if tail[0] != "RPAREN":
            self.error(f"Expected ')' or ',' in the argument list, "
                       f"got {self.describe(tail)}", tail)
        self.eat("RPAREN")
        return args

    def parse_bare_args(self):
        """Parse an unparenthesised argument list: Foo 1, 2.

        An omitted first argument is a leading comma: `GetObject , "WScript.Shell"`.
        中文：解析无括号的实参列表：Foo 1, 2。省略第一个实参时写作前导逗号。
        """
        if self.at_statement_end():
            return []
        args = []
        if self._argument_present():
            args.append(self.expr())
        else:
            tok = self.peek()
            args.append(EmptyValue().at(tok[2], tok[3]))
        while self.peek()[0] == "COMMA":
            comma = self.eat("COMMA")
            if self._argument_present():
                args.append(self.expr())
            else:
                args.append(EmptyValue().at(comma[2], comma[3]))
        return args

    # ------------------------------------------------------------ If
    def _is_label(self):
        """A name alone on its line is a label, not a colon-separated statement.

        中文：名字: 独占一行 = 行标签（不是 : 语句分隔符）。
        """
        index = self.pos
        if index + 2 >= len(self.tokens):
            return False
        return (self.tokens[index][0] == "IDENT"
                and self.tokens[index + 1][0] == "COLON"
                and self.tokens[index + 2][0] in ("NEWLINE", "EOF"))

    def label_statement(self):
        tok = self.eat("IDENT")
        self.eat("COLON")
        return LabelStmt(tok[1]).at(tok[2], tok[3])

    def _next_is_newline(self):
        """Try is a block keyword only when it stands alone; `Try = 1` is still a variable.

        中文：Try 只在独占一行时才是块语句，Try = 1 仍然可以是变量名。
        """
        index = self.pos + 1
        if index >= len(self.tokens):
            return False
        return self.tokens[index][0] in ("NEWLINE", "COLON", "EOF")

    def try_statement(self):
        tok = self.eat("IDENT")                  # Try
        body = self.parse_block(("catch", "finally", "end"),
                                "Try ... End Try", "End Try")

        handlers = []
        finally_body = None

        while True:
            nxt = self.peek()
            if self.is_word(nxt, "catch"):
                self.eat("IDENT")
                var = None
                when = None
                head = self.peek()
                if head[0] == "IDENT" and head[1].lower() != "when":
                    self.eat("IDENT")
                    var = head[1]
                if self.is_word(self.peek(), "when"):
                    self.eat("IDENT")
                    when = self.expr()
                if when is None and handlers and handlers[-1][1] is None:
                    self.error("Only the last 'Catch' can omit 'When'", nxt)
                block = self.parse_block(("catch", "finally", "end"),
                                         "Catch ... End Try", "End Try")
                handlers.append((var, when, block))
                continue
            if self.is_word(nxt, "finally"):
                self.eat("IDENT")
                finally_body = self.parse_block(("end",), "Finally ... End Try",
                                                "End Try")
                continue
            break

        self.end_block("try")
        return TryStmt(body, handlers, finally_body).at(tok[2], tok[3])

    def if_statement(self):
        tok = self.eat("IDENT")   # If
        cond = self.expr()
        self.expect_keyword("then")

        nxt = self.peek()
        if nxt[0] in ("NEWLINE", "COLON", "EOF"):
            return self.block_if(tok, cond)
        return self.inline_if(tok, cond)

    def block_if(self, tok, cond):
        branches = [(cond, self.parse_block(("elseif", "else", "end"),
                                            "If ... End If", "End If"))]
        else_body = None

        while True:
            nxt = self.peek()
            if self.is_word(nxt, "elseif"):
                self.eat("IDENT")
                new_cond = self.expr()
                self.expect_keyword("then")
                body = self.parse_block(("elseif", "else", "end"),
                                        "If ... End If", "End If")
                branches.append((new_cond, body))
                continue
            if self.is_word(nxt, "else"):
                self.eat("IDENT")
                else_body = self.parse_block(("end",), "If ... End If", "End If")
                continue
            break

        self.end_block("if")

        node = If(branches[-1][0], branches[-1][1], else_body)
        for cond_item, body in reversed(branches[:-1]):
            node = If(cond_item, body, [node])
        return node.at(tok[2], tok[3])

    def inline_if(self, tok, cond):
        """Single-line If: If cond Then stmt [: stmt...] [Else stmt...]

        中文：单行 If：If cond Then stmt [: stmt...] [Else stmt...]
        """
        then_body = self.inline_statements(("else",))
        else_body = None

        if self.is_word(self.peek(), "else"):
            self.eat("IDENT")
            else_body = self.inline_statements(())

        return If(cond, then_body, else_body).at(tok[2], tok[3])

    def inline_statements(self, stop_words):
        body = []
        while True:
            nxt = self.peek()
            if nxt[0] in ("NEWLINE", "EOF"):
                break
            if nxt[0] == "IDENT" and nxt[1].lower() in stop_words:
                break
            body.append(self.statement())
            if self.peek()[0] == "COLON":
                self.eat("COLON")
                continue
            break
        return body

    # ------------------------------------------------------------ While / Wend
    def while_statement(self):
        tok = self.eat("IDENT")   # While
        cond = self.expr()
        body = self.parse_block(("wend",), "While ... Wend", "Wend")
        self.expect_keyword("wend")
        return While(cond, body).at(tok[2], tok[3])

    # ------------------------------------------------------------ Do / Loop
    def do_loop(self):
        tok = self.eat("IDENT")   # Do

        pre_cond = None
        pre_cond_type = None
        nxt = self.peek()
        if self.is_word(nxt, "while") or self.is_word(nxt, "until"):
            self.eat("IDENT")
            pre_cond_type = "WHILE" if nxt[1].lower() == "while" else "UNTIL"
            pre_cond = self.expr()

        body = self.parse_block(("loop",), "Do ... Loop", "Loop")
        self.expect_keyword("loop")

        post_cond = None
        post_cond_type = None
        nxt = self.peek()
        if self.is_word(nxt, "while") or self.is_word(nxt, "until"):
            self.eat("IDENT")
            post_cond_type = "WHILE" if nxt[1].lower() == "while" else "UNTIL"
            post_cond = self.expr()

        return DoLoop(body, pre_cond, pre_cond_type,
                      post_cond, post_cond_type).at(tok[2], tok[3])

    # ------------------------------------------------------------ With
    def with_statement(self):
        tok = self.eat("IDENT")   # With
        obj = self.expr()
        body = self.parse_block(("end",), "With ... End With", "End With")
        self.end_block("with")
        return With(obj, body).at(tok[2], tok[3])

    # ------------------------------------------------------------ For / Next
    def for_statement(self):
        tok = self.eat("IDENT")   # For

        if self.is_word(self.peek(), "each"):
            self.eat("IDENT")
            var_tok = self.peek()
            if var_tok[0] != "IDENT":
                self.error(f"Expected a loop variable after 'For Each', "
                           f"got {self.describe(var_tok)}", var_tok)
            self.eat("IDENT")
            self.expect_keyword("in")
            iterable = self.expr()

            body = self.parse_block(("next",), "For Each ... Next", "Next")
            self.expect_keyword("next")
            self.eat_optional_next_variable()
            return ForEach(var_tok[1], iterable, body).at(tok[2], tok[3])

        var_tok = self.peek()
        if var_tok[0] != "IDENT":
            self.error(f"Expected a loop variable after 'For', "
                       f"got {self.describe(var_tok)}", var_tok)
        var_name = self.eat("IDENT")[1]

        self.eat("ASSIGN")
        start_expr = self.expr()

        self.expect_keyword("to")
        end_expr = self.expr()

        step_expr = None
        if self.is_word(self.peek(), "step"):
            self.eat("IDENT")
            step_expr = self.expr()

        body = self.parse_block(("next",), "For ... Next", "Next")
        self.expect_keyword("next")
        self.eat_optional_next_variable()

        return For(var_name, start_expr, end_expr, step_expr, body).at(tok[2], tok[3])

    def eat_optional_next_variable(self):
        """Accept the `Next i` form.

        中文：允许 'Next i' 这种写法。
        """
        nxt = self.peek()
        if nxt[0] == "IDENT" and nxt[1].lower() not in KEYWORD_STARTERS:
            self.eat("IDENT")

    # ------------------------------------------------------------ Select Case
    def select_case_statement(self):
        tok = self.eat("IDENT")        # Select
        self.expect_keyword("case")

        test_expr = self.expr()

        cases = []
        else_body = None

        while True:
            self.skip_separators()
            nxt = self.peek()

            if nxt[0] == "EOF":
                self.error("Missing 'End Select': unexpected end of file", nxt)

            if self.is_word(nxt, "case"):
                self.eat("IDENT")

                if self.is_word(self.peek(), "else"):
                    self.eat("IDENT")
                    else_body = self.parse_block(("end",), "Select Case", "End Select")
                    continue

                expr_list = [self.parse_case_condition()]
                while self.peek()[0] == "COMMA":
                    self.eat("COMMA")
                    expr_list.append(self.parse_case_condition())

                body = self.parse_block(("case", "end"), "Select Case", "End Select")
                cases.append((expr_list, body))
                continue

            if self.is_word(nxt, "end"):
                break

            self.error(f"Expected 'Case' or 'End Select', got {self.describe(nxt)}", nxt)

        self.end_block("select")
        return SelectCase(test_expr, cases, else_body).at(tok[2], tok[3])

    def parse_case_condition(self):
        tok = self.peek()

        if self.is_word(tok, "is"):
            self.eat("IDENT")
            op_tok = self.peek()
            if op_tok[0] == "ASSIGN":
                # `Case Is = 5` is legal VBScript.
                # 中文：`Case Is = 5` 是合法 VBS。
                op = "EQ"
            elif op_tok[0] in CMP_OPS:
                op = op_tok[0]
            else:
                self.error("Expected a comparison operator after 'Is', "
                           f"got {self.describe(op_tok)}", op_tok)
            self.eat(op_tok[0])
            expr = self.expr()
            return CaseCond("IS", (op, expr)).at(tok[2], tok[3])

        left = self.expr()

        if self.is_word(self.peek(), "to"):
            self.eat("IDENT")
            right = self.expr()
            return CaseCond("TO", (left, right)).at(tok[2], tok[3])

        return CaseCond("EXPR", left).at(tok[2], tok[3])

    # ------------------------------------------------------------ expressions / 中文：表达式
    # Precedence, low to high:
    #   Imp < Eqv < Xor < Or < And < Not < comparison < & < +- < Mod < \ < */ < unary
    #   < ^ < atom
    #
    # 中文：优先级（低 → 高）—— Imp < Eqv < Xor < Or < And < Not < 比较 < & < +- <
    # Mod < \ < */ < unary < ^ < atom.
    # 中文：Mod < \ < */ < 一元 < ^ < 原子。
    def expr(self):
        self.skip_newlines()
        left = self.imp_expr()
        while self.peek()[0] == "QQ":
            self.eat("QQ")
            right = self.imp_expr()
            left = BinOp(left, "COALESCE", right)
        return left

    def xword_expr(self, word, op_name, next_level):
        node = next_level()
        while self.is_word(self.peek(), word):
            op_tok = self.eat("IDENT")
            node = BinOp(node, op_name, next_level()).at(op_tok[2], op_tok[3])
        return node

    def imp_expr(self):
        return self.xword_expr("imp", "IMP", self.eqv_expr)

    def eqv_expr(self):
        return self.xword_expr("eqv", "EQV", self.xor_expr)

    def xor_expr(self):
        return self.xword_expr("xor", "XOR", self.or_expr)

    def or_expr(self):
        node = self.and_expr()
        while self.peek()[0] == "OR":
            op_tok = self.eat("OR")
            right = self.and_expr()
            node = BinOp(node, "OR", right).at(op_tok[2], op_tok[3])
        return node

    def and_expr(self):
        node = self.not_expr()
        while self.peek()[0] == "AND":
            op_tok = self.eat("AND")
            right = self.not_expr()
            node = BinOp(node, "AND", right).at(op_tok[2], op_tok[3])
        return node

    def not_expr(self):
        if self.peek()[0] == "NOT":
            tok = self.eat("NOT")
            expr = self.not_expr()
            return UnaryOp("NOT", expr).at(tok[2], tok[3])
        return self.compare_expr()

    def compare_expr(self):
        node = self.concat_expr()
        while True:
            tok = self.peek()
            if tok[0] in CMP_OPS:
                op = tok[0]
            elif tok[0] == "ASSIGN":
                # In VBScript `=` is also the equality comparison.
                # 中文：VBS 里 = 同时是相等比较。
                op = "EQ"
            elif self.is_word(tok, "is"):
                # Object identity: `x Is Nothing`.
                # 中文：对象身份比较 —— `x Is Nothing`。
                op = "IS"
            else:
                break

            self.eat(tok[0])
            right = self.concat_expr()
            node = Compare(node, op, right).at(tok[2], tok[3])
        return node

    def concat_expr(self):
        node = self.add_expr()
        while self.peek()[0] == "AMP":
            op_tok = self.eat("AMP")
            right = self.add_expr()
            node = BinOp(node, "AMP", right).at(op_tok[2], op_tok[3])
        return node

    def add_expr(self):
        node = self.mod_expr()
        while self.peek()[0] in ("PLUS", "MINUS"):
            op_tok = self.peek()
            self.eat(op_tok[0])
            right = self.mod_expr()
            node = BinOp(node, op_tok[0], right).at(op_tok[2], op_tok[3])
        return node

    def mod_expr(self):
        node = self.intdiv_expr()
        while self.is_word(self.peek(), "mod"):
            op_tok = self.eat("IDENT")
            node = BinOp(node, "MOD", self.intdiv_expr()).at(op_tok[2], op_tok[3])
        return node

    def intdiv_expr(self):
        node = self.mul_expr()
        while self.peek()[0] == "BACKSLASH":
            op_tok = self.eat("BACKSLASH")
            node = BinOp(node, "INTDIV", self.mul_expr()).at(op_tok[2], op_tok[3])
        return node

    def mul_expr(self):
        node = self.unary_expr()
        while self.peek()[0] in ("MUL", "DIV"):
            op_tok = self.peek()
            self.eat(op_tok[0])
            right = self.unary_expr()
            node = BinOp(node, op_tok[0], right).at(op_tok[2], op_tok[3])
        return node

    def unary_expr(self):
        tok = self.peek()

        if tok[0] == "PLUS":
            self.eat("PLUS")
            return UnaryOp("UPLUS", self.unary_expr()).at(tok[2], tok[3])

        if tok[0] == "MINUS":
            self.eat("MINUS")
            return UnaryOp("UMINUS", self.unary_expr()).at(tok[2], tok[3])

        return self.power_expr()

    def power_expr(self):
        # VBScript's ^ is left-associative: equal precedence binds left to right.
        # 中文：VBS 的 ^ 是左结合（同级运算符从左到右）。
        node = self.atom()
        while self.peek()[0] == "CARET":
            op_tok = self.eat("CARET")
            node = BinOp(node, "POW", self.power_operand()).at(op_tok[2], op_tok[3])
        return node

    def power_operand(self):
        """The right side of ^ may carry a unary sign, but no further ^.

        中文：^ 右侧可以带一元正负号，但不再继续吃 ^。
        """
        tok = self.peek()
        if tok[0] == "MINUS":
            self.eat("MINUS")
            return UnaryOp("UMINUS", self.power_operand()).at(tok[2], tok[3])
        if tok[0] == "PLUS":
            self.eat("PLUS")
            return UnaryOp("UPLUS", self.power_operand()).at(tok[2], tok[3])
        return self.atom()

    def atom(self):
        tok = self.peek()

        if tok[0] == "LPAREN":
            self.eat("LPAREN")
            node = self.expr()
            close = self.peek()
            if close[0] != "RPAREN":
                self.error(f"Expected ')', got {self.describe(close)}", close)
            self.eat("RPAREN")
            return self.postfix(node)

        if tok[0] == "NUMBER":
            self.eat("NUMBER")
            return self.postfix(Number(tok[1]).at(tok[2], tok[3]))

        if tok[0] == "FLOAT":
            self.eat("FLOAT")
            return self.postfix(Float(tok[1]).at(tok[2], tok[3]))

        if tok[0] == "DATE_LIT":
            self.eat("DATE_LIT")
            return self.postfix(DateLit(tok[1]).at(tok[2], tok[3]))

        if tok[0] in ("HEXNUM", "OCTNUM"):
            base = 16 if tok[0] == "HEXNUM" else 8
            self.eat(tok[0])
            return self.postfix(Number(int(tok[1][2:], base)).at(tok[2], tok[3]))

        if tok[0] == "INTERP_STRING":
            self.eat("INTERP_STRING")
            return self._interpolate(tok)
        if tok[0] == "STRING":
            self.eat("STRING")
            return self.postfix(String(tok[1]).at(tok[2], tok[3]))

        if tok[0] == "DOT":
            return self.postfix(self.with_member_target())

        if tok[0] == "IDENT":
            word = tok[1].lower()
            self.eat("IDENT")

            if word == "true":
                return self.postfix(Boolean(True).at(tok[2], tok[3]))
            if word == "false":
                return self.postfix(Boolean(False).at(tok[2], tok[3]))
            if word == "empty":
                return self.postfix(EmptyValue().at(tok[2], tok[3]))
            if word == "null":
                return self.postfix(NullValue().at(tok[2], tok[3]))
            if word == "nothing":
                return self.postfix(NothingValue().at(tok[2], tok[3]))
            if word == "addressof":
                target = self.peek()
                if target[0] != "IDENT":
                    self.error("Expected a procedure name after 'AddressOf'",
                               target)
                self.eat("IDENT")
                return self.postfix(AddressOf(target[1]).at(tok[2], tok[3]))

            if word == "new":
                cls_tok = self.peek()
                if cls_tok[0] != "IDENT":
                    self.error(f"Expected a class name after 'New', "
                               f"got {self.describe(cls_tok)}", cls_tok)
                self.eat("IDENT")
                new_args = []
                if self.peek()[0] == "LPAREN":
                    new_args = self.parse_args()
                return self.postfix(
                    NewExpr(cls_tok[1], new_args).at(tok[2], tok[3]))

            return self.postfix(Var(tok[1]).at(tok[2], tok[3]))

        if tok[0] == "EOF":
            self.error("Unexpected end of file: expression is incomplete", tok)

        self.error(f"Bad expression token: {self.describe(tok)}", tok)

    def postfix(self, node):
        """Handle the .member / ?.member and (args) suffixes.

        中文：处理 .member / ?.member 与 (args) 两种后缀。
        """
        while True:
            if self.peek()[0] in ("DOT", "QDOT"):
                node = self.member_chain(node)
                continue
            if self.peek()[0] == "LPAREN" and isinstance(node, (Var, Member)):
                args = self.parse_args()
                line = getattr(node, "line", None)
                col = getattr(node, "col", None)
                node = Call(node, args)
                if line:
                    node.at(line, col)
                continue
            break
        return node
