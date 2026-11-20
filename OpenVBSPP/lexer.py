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


import re

# The token grammar. Order is priority: the regex alternation tries them in sequence.
# Notes:
#   * COMMENT_REM / COMMENT_QUOTE must come before IDENT and the operators, otherwise
#     "rem" would lex as an identifier and a "'" comment as an unknown character.
#   * STRING honours the VBS "" escape for a literal double quote.
#   * NEWLINE covers \n, \r\n and a lone \r (see the normalisation inside lex()).
#
# 中文：词法规则表，顺序即优先级（正则交替按先后匹配）。COMMENT_REM /
# COMMENT_QUOTE has to come before IDENT and the operators, or "rem" becomes an identifier
# and "' " an unknown character.
# 中文：COMMENT_QUOTE 必须排在 IDENT 与运算符之前，否则 "rem" 会被当成标识符、"' " 会被
# STRING supports VBScript's doubled-quote escape; NEWLINE covers \n, \r\n and
# 中文：当成未知字符；STRING 支持 VBS 的 "" 双引号转义；NEWLINE 同时覆盖 \n、\r\n 与
# a lone \r, which lex() normalises.
# 中文：单独的 \r（见 lex() 里的归一化处理）。
TOK_SPEC = [
    ("COMMENT_REM",   r"(?i:\brem\b)[^\n]*"),
    ("COMMENT_QUOTE", r"'[^\n]*"),

    ("INTERP_STRING", r'\$"(?:[^"\n]|"")*"'),
    ("QDOT",    r"\?\."),
    ("QQ",      r"\?\?"),
    ("COMPOUND", r"(\+=|-=|\*=|/=|&=)"),

    ("LE",      r"<="),
    ("GE",      r">="),
    ("NE",      r"<>"),
    ("ARROW",   r"->"),

    ("EQ",      r"=="),

    ("ASSIGN",  r"="),

    ("LT",      r"<"),
    ("GT",      r">"),

    ("PLUS",    r"\+"),
    ("MINUS",   r"-"),
    ("MUL",     r"\*"),
    ("DIV",     r"/"),
    ("BACKSLASH", r"\\"),
    ("CARET",   r"\^"),

    ("FLOAT",   r"\d+\.\d+"),
    ("NUMBER",  r"\d+"),
    ("STRING",  r'"(?:[^"\n]|"")*"'),
    ("COMMA",   r","),
    ("AND",     r"(?i:\band\b)"),
    ("OR",      r"(?i:\bor\b)"),
    ("NOT",     r"(?i:\bnot\b)"),

    ("AMP",     r"&"),

    ("LPAREN",  r"\("),
    ("RPAREN",  r"\)"),
    ("DOT",     r"\."),

    ("IDENT",   r"[A-Za-z_][A-Za-z0-9_]*"),

    # Line continuation: a trailing " _" (space plus underscore) joins the next line.
    # It has to come before SKIP / NEWLINE so it swallows the whole run first.
    # 中文：行续接 —— 行尾的 " _"（空格 + 下划线）表示下一条语句继续，必须排在
    # the whole stretch can only be eaten before SKIP / NEWLINE.
    # 中文：SKIP / NEWLINE 之前才能先吃掉整段。
    ("CONTINUATION", r"[ \t]+_[ \t]*\n"),
    ("NEWLINE", r"\n"),
    ("COLON",   r":"),
    ("SKIP",    r"[ \t]+"),
    ("MISMATCH", r"."),
]
TOKEN_RE = re.compile("|".join(f"(?P<{name}>{pat})" for name, pat in TOK_SPEC), re.IGNORECASE)

COMMENT_KINDS = ("COMMENT_REM", "COMMENT_QUOTE")


def normalize_newlines(code: str) -> str:
    """Normalise CRLF and CR to LF, so a .vbs saved on Windows compiles as-is.

    中文：把 CRLF / CR 统一成 LF，这样 Windows 上保存的 .vbs 也能直接编译。
    """
    if "\r" not in code:
        return code
    return code.replace("\r\n", "\n").replace("\r", "\n")


def lex(code):
    """Split the source into (kind, value, line, col) tuples.

    line and col are 1-based and drive syntax and runtime error positions.
    Comments and whitespace produce no tokens.

    中文：把源码切成 (kind, value, line, col) 四元组。line / col 从 1 开始，用于语法
    error and run-time error positions; comments and whitespace produce no token.
    中文：错误与运行时错误的定位；注释与空白不会产生 token。
    """
    code = normalize_newlines(code)

    line = 1
    line_start = 0

    for m in TOKEN_RE.finditer(code):
        kind = m.lastgroup
        value = m.group()

        if kind in COMMENT_KINDS or kind == "SKIP":
            continue

        if kind == "CONTINUATION":
            # A line continuation yields no token, yet the newline still counts as a line.
            # 中文：行续接不产生 token，但换行本身仍要计入行号。
            line += 1
            line_start = m.end()
            continue

        col = m.start() - line_start + 1
        yield (kind, value, line, col)

        if kind == "NEWLINE":
            line += 1
            line_start = m.end()
