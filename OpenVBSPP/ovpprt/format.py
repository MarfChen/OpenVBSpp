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

把值转成 VBS 风格的文本。

VBS 在 Echo 与 & 拼接时都不会输出 "5.0"：整数形状的浮点会写成 "5"，
Empty 写成空串，Null 写成 "Null"，布尔写成 True / False。
'''

from OpenVBSPP.ovpprt.empty import EMPTY, EmptyType
from OpenVBSPP.ovpprt.null import NULL, NullType


def to_text(value):
    if value is EMPTY or isinstance(value, EmptyType):
        return ""
    if value is NULL or isinstance(value, NullType):
        return "Null"
    if isinstance(value, bytes):
        return value.decode("utf-8", "replace")
    if isinstance(value, float) and value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return str(value)
