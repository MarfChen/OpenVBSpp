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

The `VBScript.RegExp` object: VBScript's native regular expressions.

  Set re = New RegExp
  re.Pattern    = "\\d{3}-\\d{4}"
  re.Global     = True
  re.IgnoreCase = True
  If re.Test("call 555-1234") Then ...
  echo re.Replace("a1b2", "#")
  For Each m In re.Execute("a1 b22")
      echo m.Value & " @" & m.FirstIndex
  Next

Member names are stored lower-case because VBScript identifiers are case
insensitive: `re.Pattern` and `re.pattern` are the same thing. The `re` module is
imported only when a pattern is actually compiled.

中文：`VBScript.RegExp` 对象——VBS 原生的正则能力。成员名一律小写存储，因为 VBS 的
Identifiers are case-insensitive (`re.Pattern` and `re.pattern` are the same thing). `re` is imported only when
中文：标识符不区分大小写（`re.Pattern` 与 `re.pattern` 是同一个东西）。`re` 只在真正用到
a regular expression is actually used.
中文：正则时才 import。
'''


class Match:
    """One match: Value / FirstIndex / Length / SubMatches.

    中文：一次匹配——Value / FirstIndex / Length / SubMatches。
    """

    def __init__(self, match):
        self.value = match.group(0)
        self.firstindex = match.start()
        self.length = len(match.group(0))
        self.submatches = list(match.groups())

    Value = property(lambda self: self.value)
    FirstIndex = property(lambda self: self.firstindex)
    Length = property(lambda self: self.length)
    SubMatches = property(lambda self: self.submatches)

    def __setattr__(self, name, value):
        object.__setattr__(self, name.lower(), value)

    def __repr__(self):
        return f"Match({self.value!r})"


class MatchCollection(list):
    """A collection of matches: `For Each` iterates it, and it has Count / Item.

    中文：匹配集合——可以 `For Each` 遍历，也有 Count / Item。
    """

    @property
    def count(self):
        # The name must stay lower-case and override list.count with a property.
        # 中文：名字必须是小写，并且用 property 覆盖 list.count。
        return len(self)

    def item(self, index):
        return self[int(index)]

    Count = property(lambda self: len(self))
    Item = item

    def __repr__(self):
        return f"MatchCollection({len(self)})"


class RegExp:
    """A regular expression.

    中文：正则对象。
    """

    def __init__(self):
        self.pattern = ""
        self.global_ = False
        self.ignorecase = False

    # VBScript property names are case insensitive, so everything is stored lower-case.
    # 中文：VBS 的属性名不区分大小写，统一小写存储。
    def __setattr__(self, name, value):
        key = name.lower()
        if key == "global":
            key = "global_"
        object.__setattr__(self, key, value)

    Pattern = property(lambda self: self.pattern)
    Global = property(lambda self: self.global_)
    IgnoreCase = property(lambda self: self.ignorecase)

    # Methods / 中文：方法
    def _compile(self):
        # Imported lazily: `re` costs about 3 ms and most scripts never touch a RegExp.
        # 中文：延迟 import——`re` 约 3 ms，而绝大多数脚本根本不用 RegExp。
        import re

        flags = re.IGNORECASE if self.ignorecase else 0
        try:
            return re.compile(self.pattern, flags)
        except re.error as exc:
            from OpenVBSPP.exceptions import OVPPRuntimeError
            raise OVPPRuntimeError(f"Invalid regular expression: {exc}") from exc

    def Test(self, text):
        return self._compile().search(str(text)) is not None

    def Replace(self, text, replacement):
        count = 0 if self.global_ else 1
        return self._compile().sub(str(replacement), str(text), count=count)

    def Execute(self, text):
        regex = self._compile()
        source = str(text)
        collection = MatchCollection()
        if self.global_:
            for match in regex.finditer(source):
                collection.append(Match(match))
        else:
            found = regex.search(source)
            if found is not None:
                collection.append(Match(found))
        return collection

    def __repr__(self):
        return f"RegExp({self.pattern!r})"
