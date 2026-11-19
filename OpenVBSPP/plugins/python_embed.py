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

`#python ... #endpython` 代码块提取。

在词法分析之前把 Python 代码块抽出来，原地留下一个占位语句
（同时补齐换行，保证行号不变），这样 VBS 部分的行号、错误信息都不受影响。
'''

import re


BLOCK_RE = re.compile(
    r"^[ \t]*#python[ \t]*\n(.*?)^[ \t]*#endpython[ \t]*$",
    re.DOTALL | re.MULTILINE,
)

PLACEHOLDER = "__ovpp_python_block_{index}__"


def extract(code):
    """返回 (替换后的源码, {index: python 源码})。"""
    blocks = {}

    def replace(match):
        index = len(blocks)
        body = match.group(1)
        blocks[index] = body
        placeholder = PLACEHOLDER.format(index=index)
        # 保持行数一致：占位行 + 原本 block 占用的换行
        consumed = match.group(0).count("\n")
        return placeholder + "\n" * consumed

    new_code = BLOCK_RE.sub(replace, code)
    return new_code, blocks


def is_placeholder(name):
    return name.startswith("__ovpp_python_block_") and name.endswith("__")


def placeholder_index(name):
    return int(name[len("__ovpp_python_block_"):-2])
