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


from abc import ABC


class AST(ABC):
    """Base class of every OpenVBS++ AST node.

    中文：所有 OpenVBS++ AST 节点的基类。

    line and col record the node's position in the source, 1-based.
    中文：line / col 记录该节点在源码中的位置（从 1 开始），
    The transformer carries them onto the generated Python AST,
    中文：Transformer 会把它带到生成的 Python AST 上，
    so a run-time traceback points straight at the matching .vbs line.
    中文：这样运行期 traceback 能直接指向 .vbs 的对应行。
    """
    line = None
    col = None

    def at(self, line=None, col=None):
        self.line = line
        self.col = col
        return self


class Program(AST):
    def __init__(self, body): self.body = body
    def __repr__(self): return f"Program({self.body})"
