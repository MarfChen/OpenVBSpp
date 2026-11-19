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

JSON 编解码（超集扩展）。

  Set data = Json.Parse("{""name"": ""ann"", ""tags"": [1, 2]}")
  echo data("name")            ' ann（Dictionary）
  echo Json.Stringify(data)    ' {"name": "ann", "tags": [1, 2]}

映射规则（不引入新类型，全部复用已有结构）：
  JSON object  ->  Dictionary
  JSON array   ->  数组（list）
  JSON string  ->  String
  JSON number  ->  数字（整数就是整数）
  true/false   ->  Boolean
  null         ->  Null

性能说明：`json` / `datetime` 只在真正用到时才 import（见 Parse / Stringify /
_to_python），冷启动省下约 4 ms。
'''

from OpenVBSPP.exceptions import OVPPRuntimeError
from OpenVBSPP.ovpprt.boolean import VBool
from OpenVBSPP.ovpprt.collections import Dictionary
from OpenVBSPP.ovpprt.empty import EMPTY, EmptyType
from OpenVBSPP.ovpprt.null import NULL, NullType


def _to_vb(value):
    """Python 结构 -> VBS 结构。"""
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
    """VBS 结构 -> Python 结构。"""
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
    """`Json.Parse` / `Json.Stringify`。"""

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
