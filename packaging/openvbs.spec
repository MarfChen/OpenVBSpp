# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：生成单文件 `openvbs.exe`。

用法（在仓库根目录）：

    pyinstaller packaging/openvbs.spec --noconfirm --distpath dist --workpath build/pyinstaller

或直接跑 `tools/build.ps1 -Target exe`。
"""

import os

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))

# 生成的 VBS 代码在运行期会 `from OpenVBSPP.ovpprt import xxx`，
# 这些 import 是编译产物，静态分析看不到，必须显式声明。
HIDDEN_IMPORTS = [
    "OpenVBSPP",
    "OpenVBSPP.cli",
    "OpenVBSPP.compiler",
    "OpenVBSPP.exceptions",
    "OpenVBSPP.lexer",
    "OpenVBSPP.parser",
    "OpenVBSPP.runtime",
    "OpenVBSPP.transformer",
    "OpenVBSPP.typecheck",
    "OpenVBSPP.ast_nodes",
    "OpenVBSPP.plugins.python_embed",
    "OpenVBSPP.ovpprt.array",
    "OpenVBSPP.ovpprt.boolean",
    "OpenVBSPP.ovpprt.builtins",
    "OpenVBSPP.ovpprt.collections",
    "OpenVBSPP.ovpprt.com",
    "OpenVBSPP.ovpprt.cop",
    "OpenVBSPP.ovpprt.dynamic",
    "OpenVBSPP.ovpprt.empty",
    "OpenVBSPP.ovpprt.err",
    "OpenVBSPP.ovpprt.filesystem",
    "OpenVBSPP.ovpprt.format",
    "OpenVBSPP.ovpprt.ide_bridge",
    "OpenVBSPP.ovpprt.jsonobj",
    "OpenVBSPP.ovpprt.nothing",
    "OpenVBSPP.ovpprt.null",
    "OpenVBSPP.ovpprt.object",
    "OpenVBSPP.ovpprt.ref",
    "OpenVBSPP.ovpprt.regexp",
    "OpenVBSPP.ovpprt.string",
    # over 包管理器，cli 里按需 import；也被 compiler._resolve 用于按包名解析
    "OpenVBSPP.over",
]

# `com-support` 是**数据文件**而不是 Python 包（运行期用 spec_from_file_location
# 按路径加载），静态分析看不见它，必须显式作为 datas 打进去，否则 exe 里
# CreateObject("WScript.Shell") 会退化成「COM 不可用」。
BUNDLED_COM = os.path.join(ROOT, "OpenVBSPP", "vendor", "com-support")
BUNDLED_EMBED = os.path.join(ROOT, "OpenVBSPP", "vendor", "python-embed")

# `com_support.py` 是数据文件，PyInstaller 不会分析它里面的 `import win32com`，
# 所以 pywin32 的模块与 DLL 只能手动列。装没装 pywin32 都能打包：没装就打出一个
# 「COM 不可用但其它功能正常」的 exe。
HAVE_PYWIN32 = False
EXTRA_HIDDEN = []
EXTRA_BINARIES = []
try:  # pragma: no cover - 只在打包时执行
    import win32com  # noqa: F401
    HAVE_PYWIN32 = True
except Exception:
    HAVE_PYWIN32 = False

if HAVE_PYWIN32:
    EXTRA_HIDDEN += [
        "pythoncom",
        "pywintypes",
        "win32api",
        "win32con",
        "winerror",
        "win32com",
        "win32com.client",
        "win32com.client.dynamic",
        "win32com.client.gencache",
        "win32com.client.util",
    ]
    import glob
    import sys as _sys
    _dll_dir = os.path.join(_sys.prefix, "Lib", "site-packages", "pywin32_system32")
    EXTRA_BINARIES += [(dll, ".") for dll in glob.glob(os.path.join(_dll_dir, "*.dll"))]

a = Analysis(
    [os.path.join(SPECPATH, "entry.py")],
    pathex=[ROOT],
    binaries=EXTRA_BINARIES,
    datas=[(BUNDLED_COM, os.path.join("OpenVBSPP", "vendor", "com-support")),
           (BUNDLED_EMBED, os.path.join("OpenVBSPP", "vendor", "python-embed"))],
    hiddenimports=HIDDEN_IMPORTS + EXTRA_HIDDEN,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "astpretty", "unittest", "pydoc"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="openvbs",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
