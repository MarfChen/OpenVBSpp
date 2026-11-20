"""
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

PyInstaller entry point.

中文：PyInstaller 打包入口。

It lives in its own script so the `sys.exit` in __main__.py stays out of the exe analysis pass.

中文：单独放一个脚本，避免把 __main__.py 里的 `sys.exit` 带进 exe 的分析流程。
"""

import sys

from OpenVBSPP.cli import main


if __name__ == "__main__":
    sys.exit(main())
