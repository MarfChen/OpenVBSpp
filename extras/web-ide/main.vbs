rem web-ide：启动 OpenVBS++ 的简易网页 IDE
rem
rem 用法：openvbs main.vbs
rem 端口改下面这一行即可（默认 7860）。

Dim port
port = 7860

echo "OpenVBS++ IDE -> " & Ide.Url(port)
Ide.Start port
