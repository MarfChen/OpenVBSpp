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

JSON encoding and decoding (a superset extension).

  Set data = Json.Parse("{""name"": ""ann"", ""tags"": [1, 2]}")
  echo data("name")            ' ann（Dictionary）
  echo Json.Stringify(data)    ' {"name": "ann", "tags": [1, 2]}

Mapping rules, reusing existing structures instead of adding new types:

  JSON object  ->  Dictionary
  JSON array   ->  array (list)
  JSON string  ->  String
  JSON number  ->  number (integers stay integers)
  true/false   ->  Boolean
  null         ->  Null

`json` and `datetime` are imported only when they are used (see Parse, Stringify
and `_to_python`).

中文：JSON 编解码（超集扩展）。映射规则不引入新类型，全部复用已有结构：object →
to a Dictionary, an array to a list, a string to String and a number to a number (integers stay integers);
中文：Dictionary，array → 数组（list），string → String，number → 数字（整数仍是整数），
true/false become Boolean and null becomes Null. `json` and `datetime` are imported only when used.
中文：true/false → Boolean，null → Null。`json` / `datetime` 只在真正用到时才 import。
'''

from OpenVBSPP.exceptions import OVPPRuntimeError
from OpenVBSPP.ovpprt.boolean import VBool
from OpenVBSPP.ovpprt.collections import Dictionary
from OpenVBSPP.ovpprt.empty import EMPTY, EmptyType
from OpenVBSPP.ovpprt.null import NULL, NullType


def _to_vb(value):
    """Convert a Python structure into a VBS structure.

    中文：Python 结构 → VBS 结构。
    """
    if isinstance(value, dict):
        dictionary = Dictionary()
        for key, item in value.items():
            dictionary.add(str(key), _to_vb(item))
        return dictionary
    if isinstance(value, (list, tuple)):
        return [_to_vb(item) for item in value]
    if isinstance(value, bool):
        return VBool(value)
    if value is None:
        return NULL
    return value


def _to_python(value):
    """Convert a VBS structure into a Python structure.

    中文：VBS 结构 → Python 结构。
    """
    import datetime

    if isinstance(value, Dictionary):
        return {str(key): _to_python(value.item(key)) for key in value.keys()}
    if isinstance(value, (list, tuple)):
        return [_to_python(item) for item in value]
    if isinstance(value, VBool):
        return bool(value)
    if value is NULL or isinstance(value, (NullType, EmptyType)):
        return None
    if isinstance(value, datetime.datetime):
        return value.isoformat(sep=" ")
    return value


class JsonObject:
    """`Json.Parse` / `Json.Stringify`.

    中文：`Json.Parse` / `Json.Stringify`。
    """

    def Parse(self, text):
        import json as _json

        try:
            data = _json.loads(str(text))
        except _json.JSONDecodeError as exc:
            raise OVPPRuntimeError(f"Invalid JSON: {exc}") from exc
        return _to_vb(data)

    def Stringify(self, value, indent=None):
        import json as _json

        data = _to_python(value)
        if indent is None:
            return _json.dumps(data, ensure_ascii=False)
        return _json.dumps(data, ensure_ascii=False, indent=int(indent))

    def __repr__(self):
        return "Json()"


JSON = JsonObject()
