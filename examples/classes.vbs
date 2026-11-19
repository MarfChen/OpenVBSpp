rem 类、属性、方法、Dictionary
Option Explicit

Class Counter
    Private n
    Public Property Get Value()
        Value = n
    End Property
    Public Sub Bump(ByVal step)
        n = n + step
    End Sub
    Private Sub Class_Initialize()
        n = 0
    End Sub
End Class

Dim c
Set c = New Counter
c.Bump 3
c.Bump 4
echo "counter=" & c.Value

Dim d, k
Set d = CreateObject("Scripting.Dictionary")
d.Add "one", 1
d("two") = 2
echo "count=" & d.Count
For Each k In d
    echo k & "=" & d(k)
Next
