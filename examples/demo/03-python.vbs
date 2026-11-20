' 3) Python interop: import it explicitly, then use #python blocks and COP.
' 中文：Python 互操作——显式导入后使用 #python 代码块与 COP 公共对象池。

<script src="python-embed" />

#python
import platform, sys
print("host python:", platform.python_version(), "on", sys.platform)
#endpython

COP("greeting") = "hello from VBScript"
echo "COP -> " & COP("greeting")

' Anything Python can do is one block away:
#python
import math
print("pi ~", round(math.pi, 4))
#endpython
