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

网页 IDE 的运行时桥：让 VBS 脚本也能启动 IDE。

  Ide.Start 7860          ' 启动并阻塞（Ctrl+C 结束）
  echo Ide.Url(7860)      ' 打印地址
  Ide.Open 7860           ' 用默认浏览器打开
'''

import threading
import webbrowser


class IdeBridge:
    """`Ide` 对象：把 `openvbs ide` 的能力暴露给 VBS 脚本。"""

    def __init__(self):
        self.port = 7860
        self.host = "127.0.0.1"
        self._thread = None

    def Url(self, port=None):
        return f"http://{self.host}:{int(port or self.port)}/"

    def Open(self, port=None):
        url = self.Url(port)
        threading.Thread(target=webbrowser.open, args=(url,), daemon=True).start()
        return url

    def Start(self, port=7860, open_browser=True):
        """启动 IDE（阻塞直到 Ctrl+C）。"""
        from OpenVBSPP.ide.server import serve

        self.port = int(port)
        print(f"OpenVBS++ IDE: {self.Url()}")
        serve(host=self.host, port=self.port, open_browser=bool(open_browser))
        return self.Url()

    def __repr__(self):
        return f"Ide({self.host}:{self.port})"


IDE = IdeBridge()
