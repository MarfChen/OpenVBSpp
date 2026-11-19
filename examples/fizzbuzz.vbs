rem 函数 + 循环 + 条件：经典的 FizzBuzz
Option Explicit

Function IsDivisible(n, d)
    IsDivisible = (n / d == Int(n / d))
End Function

Dim i

For i = 1 To 15
    If IsDivisible(i, 15) Then
        echo "FizzBuzz"
    ElseIf IsDivisible(i, 3) Then
        echo "Fizz"
    ElseIf IsDivisible(i, 5) Then
        echo "Buzz"
    Else
        echo i
    End If
Next
