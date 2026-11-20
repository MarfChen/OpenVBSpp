' 5) Multi-file: a namespace keeps the included module's globals separate.
' 中文：多文件——用命名空间隔离被包含模块的全局变量。

<namespace src="lib/utils.vbs" alias="u">

echo u.Double(21)
echo u.Shout("openvbs")

' Same file, merged this time: everything shares one global scope.
<script src="lib/utils.vbs">
echo "merged: " & Double(5)
