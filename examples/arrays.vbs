rem Dim / ReDim / 数组 / 内置函数
Option Explicit

Dim scores(4)
Dim i
Dim total

For i = 0 To UBound(scores)
    scores(i) = (i + 1) * 10
Next

total = 0
For i = 0 To UBound(scores)
    total = total + scores(i)
Next
echo "total=" & total

ReDim Preserve scores(5)
scores(5) = 60
echo UBound(scores) & " -> " & scores(5)

Dim names
names = Split("ann,bob,cid", ",")
echo Join(names, " / ")
echo UCase(names(0)) & Len(names(2))

Dim grid(1, 1)
grid(1, 1) = 42
echo grid(1, 1)
