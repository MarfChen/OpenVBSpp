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

Multi-file declarations: two XML-style directives, a superset extension.

  <script src="utils.vbs">              merge: the target file's procedures become usable
  <namespace src="utils.vbs" alias="u"> namespace: u.Double(21)

The attributes are flexible: `alias` and `nickyname` both name the alias, `src` points
at the target file.

Directives are extracted before lexing and replaced by a placeholder statement, so line
numbers never shift; the real loading happens in the compiler, at compile time.

中文：多文件声明（超集扩展）—— `<script src>` 合并（目标文件的过程直接可用）与
and `<namespace src alias>` for a namespace (`u.Double(21)`). `alias` and `nickyname` both name the alias,
中文：`<namespace src alias>` 命名空间（`u.Double(21)`）。`alias` / `nickyname` 都表示别名，
`src` names the target file. Directives are pulled out before lexing and replaced in place by a placeholder statement, so line numbers hold; the actual load happens
中文：`src` 表示目标文件。指令在词法之前抽出、原地留占位语句（行号不变），真正的文件加载
in the compiler at compile time.
中文：由 compiler 在编译期完成。
'''

import re


# Either a self-closing <script src="..." /> or an opening <script src="...">.
# 中文：自闭合的 `<script src="..." />` 或开标签 `<script src="...">`。
SCRIPT_RE = re.compile(
    r"^[ \t]*<script\s+src\s*=\s*(['\"])(?P<src>[^'\"]+)\1\s*/?>[ \t]*$",
    re.IGNORECASE | re.MULTILINE,
)

# <namespace src="..." alias="..." />
NAMESPACE_RE = re.compile(
    r"^[ \t]*<namespace\s+(?P<attrs>[^>]*?)\s*/?>[ \t]*$",
    re.IGNORECASE | re.MULTILINE,
)

ATTR_RE = re.compile(
    r"(?P<key>[A-Za-z_][\w-]*)\s*=\s*(?P<quote>['\"])(?P<value>[^'\"]*)(?P=quote)"
)

PLACEHOLDER = "__ovpp_module_{index}__"


class Directive:
    """One multi-file directive.

    中文：一条多文件声明。
    """

    __slots__ = ("kind", "src", "alias", "line")

    def __init__(self, kind, src, alias, line):
        self.kind = kind          # "script" | "namespace"
        # Target file, relative to the file that declares it.
        # 中文：目标文件（相对当前文件）。
        self.src = src
        # Alias of a namespace directive.
        # 中文：namespace 的别名。
        self.alias = alias
        self.line = line

    def __repr__(self):
        return f"Directive({self.kind}, {self.src!r}, alias={self.alias!r})"


def _line_of(text, offset):
    return text.count("\n", 0, offset) + 1


def _parse_attrs(text):
    return {m.group("key").lower(): m.group("value") for m in ATTR_RE.finditer(text)}


def extract(code):
    """Extract the directives and return (rewritten source, [Directive]).

    中文：抽出多文件指令，返回 (替换后的源码, [Directive])。
    """
    directives = []

    def script_sub(match):
        index = len(directives)
        directives.append(Directive("script", match.group("src"), None,
                                    _line_of(code, match.start())))
        return PLACEHOLDER.format(index=index) + "\n" * match.group(0).count("\n")

    def namespace_sub(match):
        attrs = _parse_attrs(match.group("attrs"))
        source = attrs.get("src", "")
        alias = attrs.get("alias") or attrs.get("nickyname") or attrs.get("tar") or ""
        index = len(directives)
        line = _line_of(code, match.start())
        directives.append(Directive("namespace", source, alias, line))
        return PLACEHOLDER.format(index=index) + "\n" * match.group(0).count("\n")

    new_code = SCRIPT_RE.sub(script_sub, code)
    new_code = NAMESPACE_RE.sub(namespace_sub, new_code)
    return new_code, directives


def is_placeholder(name):
    return name.startswith("__ovpp_module_") and name.endswith("__")


def placeholder_index(name):
    return int(name[len("__ovpp_module_"):-2])
