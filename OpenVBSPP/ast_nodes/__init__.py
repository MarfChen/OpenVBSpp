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
from ._base import *
from ._expressions import *
from ._control import *
from ._functions import *
from ._classes import *
from ._statements import *


__version__ = "0.9.20"
__all__ = [
    "AST",
    "Program",
    # statements
    # 中文：语句
    "Assign",
    "MemberAssign",
    "AssignIndex",
    "Print",
    "OptionExplicit",
    "OptionTyping",
    "OptionCompare",
    "DimDecl",
    "ReDimDecl",
    "ConstDecl",
    "EraseStmt",
    "OnErrorResumeNext",
    "OnErrorGoToZero",
    "OnErrorGoToLabel",
    "ResumeStmt",
    # expressions
    # 中文：表达式
    "Number",
    "Float",
    "String",
    "Boolean",
    "EmptyValue",
    "NullValue",
    "NothingValue",
    "Var",
    "Member",
    "BinOp",
    "UnaryOp",
    "Compare",
    "Call",
    "AddressOf",
    "NewExpr",
    # control flow
    # 中文：控制流
    "If",
    "While",
    "DoLoop",
    "Break",
    "ExitDo",
    "ExitFor",
    "ExitSub",
    "ExitFunction",
    "With",
    "TryStmt",
    "LabelStmt",
    "GoToStmt",
    "ForEach",
    "SelectCase",
    "CaseCond",
    "For",
    # procedures and classes
    # 中文：过程与类
    "Param",
    "SubDef",
    "FunctionDef",
    "CallStmt",
    "FieldDecl",
    "PropertyDef",
    "ClassDef",
    "AttributeDecl",
]
