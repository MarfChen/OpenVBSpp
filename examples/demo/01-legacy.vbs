' 1) Plain VBScript written for cscript.exe — runs unchanged.
' 中文：为 cscript.exe 写的普通 VBScript，原样运行。

Dim i, total
For i = 1 To 5
    total = total + i * i
Next
echo "sum of squares 1..5 = " & total

Function Greet(name)
    Greet = "Hello, " & name & "!"
End Function

echo Greet("VBScript")
