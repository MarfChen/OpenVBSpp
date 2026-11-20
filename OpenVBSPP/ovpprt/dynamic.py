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

Dynamic execution: `Execute` and `Eval`.

Both reuse the compiler itself: the string is compiled as an OpenVBS++ source and
run in the caller's global scope, matching VBScript, where both always run in the
global scope. `compiler` is imported lazily to avoid the import cycle
ovpprt → compiler → transformer → ovpprt.

中文：动态执行 `Execute` / `Eval`。两者都复用编译器：把字符串当成一段 OpenVBS++
The source is compiled and run in the caller's global scope, matching VBScript, which always runs these globally.
中文：源码编译后在调用方的全局作用域里执行（与 VBS 一致——它们总是在全局作用域运行）。
`compiler` is imported lazily to avoid the ovpprt -> compiler -> transformer -> ovpprt import cycle.
中文：`compiler` 用延迟导入，避免 ovpprt → compiler → transformer → ovpprt 的循环导入。
'''

from OpenVBSPP.ovpprt.empty import EMPTY


EVAL_RESULT = "__ovpp_eval_result"


def _compile(source, mode_name):
    from OpenVBSPP.compiler import vbs_compile
    return vbs_compile(str(source), filename=mode_name)


def execute(source, scope):
    """`Execute "statements"`: run the code in the given global scope.

    中文：`Execute "语句..."`——在给定的全局作用域里执行。
    """
    exec(_compile(source, "<execute>"), scope)


def evaluate(source, scope):
    """`Eval("expression")`: evaluate the expression and return its value.

    中文：`Eval("表达式")`——求值并返回结果。
    """
    wrapped = f"{EVAL_RESULT} = ({source})"
    exec(_compile(wrapped, "<eval>"), scope)
    return scope.pop(EVAL_RESULT, EMPTY)
