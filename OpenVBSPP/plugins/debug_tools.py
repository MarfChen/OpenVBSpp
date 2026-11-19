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


import ast

try:                      # astpretty 只是调试辅助，缺失时不应影响编译
    import astpretty
except ImportError:       # pragma: no cover
    astpretty = None


def ast2str(node,
            print_ast:bool=False,
            print_code:bool=True,
            save_code:bool=False,
            file_name:str="output.py") -> str:
    """把 Python AST 反编译成源码文本，并可打印 / 保存。

    Args:
        node (ast.AST): 要被处理的 AST 节点。
        print_ast (bool): 是否打印 AST 结构（需要 astpretty，缺失时忽略）。
        print_code (bool): 是否打印生成的代码。
        save_code (bool): 是否保存生成的代码。
        file_name: 保存路径。

    Returns:
        生成的 Python 源码字符串。
    """
    if print_ast:
        if astpretty is None:
            print("[ast2str] astpretty is not installed; skip AST dump "
                  "(pip install astpretty)")
        else:
            astpretty.pprint(node)

    code = ast.unparse(node)
    if print_code:
        print(code)
    if save_code:
        with open(file_name, "w", encoding="utf-8") as f:
            f.write(code)

    return code
