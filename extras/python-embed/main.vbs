' python-embed —— OpenVBS++ 的 Python 互操作能力包。
'
' 这个文件本身不做什么，它是"能力声明"：只要脚本里写了
'
'     <script src="python-embed" />
'
' 编译器就会打开下面两项能力（0.9.20 起必须显式导入，旧写法 Option COP 已移除）：
'
'     #python ... #endpython   —— 内嵌 Python 代码块（行号保持不变）
'     COP                      —— 与 Python 共享的全局字典
'
' 用法见 README / doc/ 里的「Python 互操作」一节。
