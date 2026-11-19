rem On Error Resume Next / Err 对象 / With 块
Option Explicit

Dim value

On Error Resume Next

value = 10 / 0
echo "err=" & Err.Number
Err.Clear

value = 10 / 2
echo "ok=" & value

With Err
    .Raise 1001, "demo", "手动抛出的错误"
End With
echo "raised=" & Err.Number

Sub TryIt()
    On Error Resume Next
    Dim x
    x = "abc" + 1
    echo "inner=" & Err.Number
End Sub

Call TryIt
echo "after sub err=" & Err.Number
