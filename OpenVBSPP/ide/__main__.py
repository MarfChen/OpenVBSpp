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

`python -m OpenVBSPP.ide [--port 7860]`
'''

import argparse
import sys

from OpenVBSPP.ide.server import serve


def main(argv=None):
    parser = argparse.ArgumentParser(prog="openvbs ide",
                                     description="Start the OpenVBS++ web IDE / 启动 OpenVBS++ 网页 IDE")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=7860)
    parser.add_argument("--workspace", default=None,
                        help="Script directory, default ~/.openvbspp/scripts / 脚本存放目录")
    parser.add_argument("--no-browser", action="store_true",
                        help="Do not open a browser automatically / 不自动打开浏览器")
    args = parser.parse_args(argv)
    serve(host=args.host, port=args.port, workspace=args.workspace,
          open_browser=not args.no_browser)
    return 0


if __name__ == "__main__":
    sys.exit(main())
