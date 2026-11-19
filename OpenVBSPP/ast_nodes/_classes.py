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

类与对象相关的 AST 节点（VBS 对象模型）。
'''


from OpenVBSPP.ast_nodes._base import AST


class FieldDecl(AST):
    """类里的字段：Public x / Private y / Dim z（可带 -> Type）"""
    def __init__(self, name, type_name=None, visibility="public", dims=None):
        self.name = name
        self.type_name = type_name
        self.visibility = visibility
        self.dims = dims
    def __repr__(self):
        return f"FieldDecl({self.visibility} {self.name}: {self.type_name})"


class PropertyDef(AST):
    """Property Get/Let/Set ... End Property"""
    def __init__(self, kind, name, params, body, type_name=None):
        self.kind = kind          # "get" / "let" / "set"
        self.name = name
        self.params = params      # list[Param]
        self.body = body
        self.type_name = type_name
    def __repr__(self):
        return f"PropertyDef({self.kind} {self.name}, {self.params})"


class ClassDef(AST):
    """Class Name [Inherits Base] ... End Class"""
    def __init__(self, name, body, base=None):
        self.name = name
        self.body = body          # list[AST]
        self.base = base          # 基类名（超集扩展）
    def __repr__(self):
        return f"ClassDef({self.name}, {len(self.body)} members)"


class NewExpr(AST):
    """Set x = New Foo [ (args) ]"""
    def __init__(self, name, args=None):
        self.name = name
        self.args = args or []
    def __repr__(self):
        return f"NewExpr({self.name})"


class AttributeDecl(AST):
    """Attribute 语句：目前只认 Value.VB_UserMemId = 0（默认成员）"""
    def __init__(self, member, key, value):
        self.member = member
        self.key = key
        self.value = value
    def __repr__(self):
        return f"AttributeDecl({self.member}.{self.key} = {self.value})"
