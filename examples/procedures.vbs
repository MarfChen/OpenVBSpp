rem Sub / Function / ByRef / ByVal / 递归
Option Explicit

Sub PrintLine(text)
    echo "-- " & text
End Sub

Function Fact(n)
    If n <= 1 Then
        Fact = 1
    Else
        Fact = n * Fact(n - 1)
    End If
End Function

Sub Swap(ByRef a, ByRef b)
    Dim t
    t = a
    a = b
    b = t
End Sub

Sub Bump(ByVal n)
    n = n + 100
    echo "inside=" & n
End Sub

Dim x, y

Call PrintLine("start")
PrintLine "factorial"
echo Fact(5)

x = 1
y = 2
Swap x, y
echo x & "," & y

Bump x
echo "outside=" & x
