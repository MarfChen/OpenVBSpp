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

A minimal HTTP server for the built-in web IDE, standard library only.

  GET  /                     the frontend page
  GET  /api/capabilities     capability declaration (including the reserved package info)
  GET  /api/examples         the built-in examples
  GET  /api/scripts          script names in the workspace
  POST /api/run              {source, filename} -> output / errors / elapsed time
  POST /api/check            {source} -> diagnostics
  POST /api/save             {name, source}
  POST /api/load             {name}
  POST /api/packages/search  -> 501 (reserved)

中文：极简 HTTP 服务 —— 只为内置网页 IDE 服务，只用标准库。接口同上：页面、
capabilities, built-in examples, script list, run, check, save and load; `/api/packages/search` is reserved and answers
中文：能力声明、内置示例、脚本列表、运行、检查、保存、加载；`/api/packages/search` 预留返回
501。
'''

import json
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from OpenVBSPP.ide import dialogs, service
from OpenVBSPP.ide.packages import get_provider


STATIC_DIR = Path(__file__).resolve().parent / "static"
# 2 MB, so a huge file is not accidentally posted.
# 中文：2 MB，防止误传大文件。
MAX_BODY = 2 * 1024 * 1024


class IdeHandler(BaseHTTPRequestHandler):
    workspace = service.Workspace()
    server_version = "OpenVBSPP-IDE"

    # --------------------------------------------------------------- helpers
    # 中文：工具函数
    # Stay quiet: no console spam.
    # 中文：安静一点，别刷屏。
    def log_message(self, fmt, *args):
        pass

    def _send(self, status, payload, content_type="application/json; charset=utf-8"):
        body = payload if isinstance(payload, bytes) else json.dumps(
            payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0 or length > MAX_BODY:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return {}

    # ---------------------------------------------------------------- GET
    def do_GET(self):
        try:
            return self._do_get()
        except Exception as exc:
            return self._send(500, {"ok": False, "message": f"{type(exc).__name__}: {exc}"})

    def _do_get(self):
        path = urlparse(self.path).path

        if path in ("/", "/index.html"):
            index = STATIC_DIR / "index.html"
            if not index.is_file():
                return self._send(500, {"message": "UI files are missing"})
            return self._send(200, index.read_bytes(), "text/html; charset=utf-8")

        if path == "/api/capabilities":
            return self._send(200, service.capabilities())

        if path == "/api/examples":
            return self._send(200, {"examples": service.examples()})

        if path == "/api/dialogs":
            return self._send(200, dialogs.platform_hint())

        if path == "/api/scripts":
            return self._send(200, {"scripts": self.workspace.list()})

        return self._send(404, {"message": f"no route for {path}"})

    # ---------------------------------------------------------------- POST
    def do_POST(self):
        try:
            return self._do_post()
        # Last resort, so the connection is not dropped outright.
        # 中文：兜底，避免连接被直接断开。
        except Exception as exc:
            return self._send(500, {"ok": False, "message": f"{type(exc).__name__}: {exc}"})

    def _do_post(self):
        path = urlparse(self.path).path
        data = self._read_json()
        source = data.get("source") or ""
        filename = data.get("filename") or "<ide>"

        if path == "/api/run":
            return self._send(200, service.run_source(source, filename))

        if path == "/api/check":
            return self._send(200, service.check_source(source, filename))

        if path == "/api/save":
            name = (data.get("name") or "untitled").strip()
            try:
                return self._send(200, self.workspace.save(name, source))
            except (ValueError, OSError) as exc:
                return self._send(400, {"ok": False, "message": str(exc)})

        if path == "/api/load":
            name = (data.get("name") or "").strip()
            try:
                return self._send(200, self.workspace.load(name))
            except (ValueError, OSError) as exc:
                return self._send(400, {"ok": False, "message": str(exc)})

        # ---- Native dialogs (Windows / Linux / macOS) / 中文：原生对话框 ----
        if path == "/api/dialog/open":
            chosen = dialogs.ask_open(initial_dir=data.get("initial_dir"))
            if not chosen:
                return self._send(200, {"ok": False, "cancelled": True})
            try:
                text = dialogs.read_script(chosen)
            except (OSError, UnicodeDecodeError) as exc:
                return self._send(400, {"ok": False, "message": str(exc)})
            return self._send(200, {"ok": True, "path": chosen, "source": text})

        if path == "/api/dialog/save":
            default_name = data.get("name") or "untitled.vbs"
            chosen = dialogs.ask_save(default_name,
                                      initial_dir=data.get("initial_dir"))
            if not chosen:
                return self._send(200, {"ok": False, "cancelled": True})
            try:
                saved = dialogs.write_script(chosen, source)
            except OSError as exc:
                return self._send(400, {"ok": False, "message": str(exc)})
            return self._send(200, {"ok": True, "path": saved})

        # ---- Read and write by absolute path: saving again or reloading after a
        # dialog pick. / 中文：按绝对路径读写（对话框选完之后的再次保存 / 重新加载）----
        if path == "/api/file/open":
            target = data.get("path") or ""
            if not target:
                return self._send(400, {"ok": False, "message": "path is required"})
            try:
                text = dialogs.read_script(target)
            except (OSError, UnicodeDecodeError) as exc:
                return self._send(400, {"ok": False, "message": str(exc)})
            return self._send(200, {"ok": True, "path": target, "source": text})

        if path == "/api/file/save":
            target = data.get("path") or ""
            if not target:
                return self._send(400, {"ok": False, "message": "path is required"})
            try:
                saved = dialogs.write_script(target, source)
            except OSError as exc:
                return self._send(400, {"ok": False, "message": str(exc)})
            return self._send(200, {"ok": True, "path": saved})

        # Reserved: the package manager. / 中文：预留 —— 包管理器。
        if path == "/api/packages/search":
            return self._send(501, {
                "ok": False,
                "reserved": True,
                "message": "package manager is reserved for a future release",
                "provider": type(get_provider()).__name__,
            })

        return self._send(404, {"message": f"no route for {path}"})


def create_server(host="127.0.0.1", port=7860, workspace=None):
    if workspace is not None and not isinstance(workspace, service.Workspace):
        # A plain directory string is accepted.
        # 中文：允许直接传目录字符串。
        workspace = service.Workspace(workspace)
    handler = IdeHandler
    if workspace is not None:
        handler = type("BoundIdeHandler", (IdeHandler,), {"workspace": workspace})
    httpd = ThreadingHTTPServer((host, port), handler)
    return httpd


def serve(host="127.0.0.1", port=7860, workspace=None, open_browser=True):
    httpd = create_server(host, port, workspace)
    url = f"http://{host}:{httpd.server_address[1]}/"
    print(f"OpenVBS++ IDE started at {url}")
    print("Press Ctrl+C to quit.")
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        httpd.server_close()
