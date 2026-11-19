rem Try / Catch / Finally：结构化异常（超集扩展）
rem 与 On Error Resume Next 可以共存：Try 块内优先交给 Catch。

Function Divide(a, b)
    Try
        If b = 0 Then Err.Raise 11, "Demo", "除数不能为 0"
        Divide = a / b
    Catch e When e.Number = 11
        echo "  [业务错误] " & e.Description
        Divide = 0
    Catch e
        echo "  [意外错误] " & e.Number & " @ 行 " & e.Line
        Divide = -1
    Finally
        echo "  [finally] 清理"
    End Try
End Function

echo Divide(10, 2)
echo Divide(1, 0)

rem 嵌套 Try 与重抛
Sub Inner()
    Try
        Err.Raise 53
    Catch e
        echo "  内层捕获 " & e.Number
        Err.Raise 99
    End Try
End Sub

Try
    Inner
Catch e
    echo "  外层捕获 " & e.Number
End Try
