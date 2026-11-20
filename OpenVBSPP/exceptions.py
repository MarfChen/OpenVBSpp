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


class OVPPError(Exception):
    number: int = 0 
    name: str = "OVPPError"

    def __init__(self, message=None, description=None,
                 line_number=None, column=None, source_line=None,
                 help_context=None, help_file=None):

        self.message = message or self.name
        self.description = description
        self.line_number = line_number
        self.column = column
        self.source_line = source_line
        self.help_context = help_context
        self.help_file = help_file

        super().__init__(self.message)

    def __str__(self):
        parts = [f"{self.name} ({self.number}): {self.message}"]
        if self.description:
            parts.append(f"\nDescription: {self.description}")
        if self.line_number is not None:
            loc = f"Line {self.line_number}"
            if self.column is not None:
                loc += f", Column {self.column}"
            parts.append(f"\nLocation: {loc}")
        if self.source_line:
            parts.append(f"\nSource: {self.source_line}")
        return "".join(parts)

# VBScript-compatible error classes. VBScript has more than 460 error numbers and only
# the common ones are modelled here. This is a VBScript compiler, not a full interpreter,
# so the set of needed error types may differ from VBScript proper.
# 中文：VBScript 兼容的错误类型。VBS 错误类型有 460 多种，这里只实现了一部分常用的。
# This is a VBS compiler, not the standard VBS interpreter, so the error types it needs
# may not match the standard ones exactly.
# 中文：这是 VBS 编译器项目，不是标准 VBS 解释器，需要的错误类型可能和标准 VBS 不完全一样。

class OVPPSyntaxError(OVPPError):
    number = 1002
    name = "Syntax error"


class OVPPRuntimeError(OVPPError):
    number = 5 
    name = "Runtime error"


class OVPPTypeError(OVPPError):
    number = 13 
    name = "Type mismatch"


class OVPPVariableNotDefinedError(OVPPError):
    number = 1001 
    name = "Variable is undefined"


class OVPPProcedureNotDefinedError(OVPPError):
    number = 35
    name = "Sub or Function not defined"


class OVPPDivisionByZeroError(OVPPError):
    number = 11
    name = "Division by zero"


class OVPPOverflowError(OVPPError):
    number = 6
    name = "Overflow"


class OVPPArgumentError(OVPPError):
    number = 450
    name = "Wrong number of arguments"


class OVPPObjectNotSupportedError(OVPPError):
    number = 438
    name = "Object doesn't support this property or method"


class OVPPPermissionDeniedError(OVPPError):
    number = 70
    name = "Permission denied"


class OVPPFileNotFoundError(OVPPError):
    number = 53
    name = "File not found"


class OVPPOutOfRangeError(OVPPError):
    number = 9
    name = "Subscript out of range"


class OVPPInvalidProcedureCallError(OVPPError):
    number = 5
    name = "Invalid procedure call"


class OVPPAutomationError(OVPPError):
    number = -2147467259  
    name = "Automation error"
