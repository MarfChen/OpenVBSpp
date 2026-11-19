"""最小示例：编译并运行一段 OpenVBS++ 代码。

运行：
    python cli_demo.py
"""

import sys

from OpenVBSPP.cli import DEMO
from OpenVBSPP.runtime import vbs_run


if __name__ == "__main__":
    vbs_run(DEMO, filename="<demo>")
