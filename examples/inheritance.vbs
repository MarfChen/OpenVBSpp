rem 继承 + 方法 ByRef + 属性
Option Explicit

Class Counter
    Private n
    Public Property Get Value()
        Value = n
    End Property
    Public Sub Bump(ByRef step)
        n = n + step
        step = step * 2
    End Sub
    Private Sub Class_Initialize()
        n = 0
    End Sub
End Class

Class NamedCounter Inherits Counter
    Public label
    Public Function Describe()
        Describe = label & "=" & Value
    End Function
End Class

Dim c, step
Set c = New NamedCounter
c.label = "hits"
step = 1
c.Bump step
c.Bump step
echo c.Describe()
echo "step=" & step
