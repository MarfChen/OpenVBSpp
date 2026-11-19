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

动态执行：Execute / Eval。

两者都复用编译器本身：把字符串当成一段 OpenVBS++ 源码编译后在
**调用方的全局作用域**里执行（与 VBS 的 Execute/Eval 语义一致 ——
它们总是在全局作用域运行）。

这里延迟导入 compiler，避免 ovpprt → compiler → transformer → ovpprt 的循环导入。
'''

from OpenVBSPP.ovpprt.empty import EMPTY


EVAL_RESULT = "__ovpp_eval_result"


def _compile(source, mode_name):
    from OpenVBSPP.compiler import vbs_compile
    return vbs_compile(str(source), filename=mode_name)


def execute(source, scope):
    """Execute "语句..."：在给定的全局作用域里执行。"""
    exec(_compile(source, "<execute>"), scope)


def evaluate(source, scope):
    """Eval("表达式")：求值并返回结果。"""
    wrapped = f"{EVAL_RESULT} = ({source})"
    exec(_compile(wrapped, "<eval>"), scope)
    return scope.pop(EVAL_RESULT, EMPTY)
