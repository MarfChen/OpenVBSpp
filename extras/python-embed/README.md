# python-embed —— OpenVBS++ 的 Python 互操作能力包

这个包本身没有可执行代码：它向语言**声明一项能力**。脚本里用

```vbs
<script src="python-embed" />
```

导入之后才能使用：

- `#python ... #endpython` 代码块（直接写 Python，行号保持不变）
- `COP` 公共对象池（VBS 与 Python 共享同一个字典）

## 为什么需要显式导入

0.9.19 及以前，`#python` 无条件可用、`COP` 靠 `Option COP` 打开。
0.9.20 起两者统一由这个包提供：

- 没有导入就用 `#python` → 编译期报错，提示补上 `<script src="python-embed" />`
- 仍然写 `Option COP` → 报错说明该语句已移除

这样脚本一眼能看出「这段代码依赖 Python 互操作」，也方便给不同项目
钉不同版本的实现。

## 三种用法

1. **什么都不做（推荐）**：包随 OpenVBS++ 分发（`OpenVBSPP/vendor/python-embed/`），
   `<script src="python-embed" />` 直接生效，不需要安装。
2. **装进项目**：`openvbs over add python-embed` —— 复制到 `packages/python-embed/`，
   之后用 `openvbs over list` 能看到、`openvbs over remove python-embed` 能卸掉。
3. **换实现**：自己写一个带 `"capabilities": ["python-embed"]` 的包放进
   `packages/python-embed/`，项目内的副本优先于随包分发的副本。

## 示例

```vbs
<script src="python-embed" />

#python
import math
answer = math.factorial(5)
#endpython

echo answer          ' 120

COP("from-vbs") = "hello"
#python
from OpenVBSPP.ovpprt.cop import COP as PYCOP
PYCOP["from-python"] = PYCOP["from-vbs"].upper()
#endpython
echo COP("from-python")   ' HELLO
```

## 已知限制

- `#python` 块共享脚本的全局命名空间；写 `import` 时注意别和 VBS 变量重名。
- 语法错误仍由 Python 自己报，行号是 VBS 文件里的真实行号。
