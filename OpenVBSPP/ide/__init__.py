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

内置网页 IDE（可选组件，零第三方依赖）。

  openvbs ide                 # 启动并打开浏览器
  python -m OpenVBSPP.ide     # 同上
'''

from OpenVBSPP.ide.service import (Workspace, capabilities, check_source,
                                   examples, run_source)
from OpenVBSPP.ide.server import create_server, serve

__all__ = ["Workspace", "capabilities", "check_source", "examples",
           "run_source", "create_server", "serve"]
