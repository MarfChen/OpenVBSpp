rem 语法糖：冒号分隔、单行 If、行续接
Option Explicit

Dim a, b, total, i

a = 1 : b = 2
total = a + _
        b
echo total

If a = 1 Then echo "a is one"
If b = 1 Then echo "b is one" Else echo "b is not one"

For i = 1 To 3 : echo i : Next
