"""
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
OpenVBS++ command line entry point, a convenience script.

中文：OpenVBS++ 命令行入口（便捷脚本）。

Usage:

中文：用法：
    python cli.py                            # run the built-in demo
                                             # 运行内置示例
    python cli.py examples/control1.vbs      # run a given script
                                             # 运行指定脚本
    python cli.py examples/for_loop.vbs -v   # print tokens / AST / generated Python
                                             # 打印 token、AST 与生成的 Python
    python cli.py examples/control1.vbs --save   # also write .pyc and .ast.txt
                                                # 额外写出 .pyc 与 .ast.txt
"""

import sys

from OpenVBSPP.cli import main


if __name__ == "__main__":
    sys.exit(main())
