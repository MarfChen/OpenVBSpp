' 2) Modern extensions: classes, inheritance, Try/Catch, typed parameters.
' 中文：超集扩展——类、继承、Try/Catch、参数类型标注。

Option Typing   ' compile-time type checking (optional)

Class Counter
    Private n
    Public Property Get Value() -> Integer
        Value = n
    End Property
    Public Sub Bump(ByVal step -> Integer)
        n = n + step
    End Sub
    Private Sub Class_Initialize()
        n = 0
    End Sub
End Class

Class NamedCounter Inherits Counter
    Public label -> String
    Public Function Describe() -> String
        Describe = label & " = " & Value
    End Function
End Class

Dim c, step
Set c = New NamedCounter
c.label = "hits"

Try
    c.Bump 3
    c.Bump 4
    Err.Raise 11, "demo", "boom"
Catch e When e.Number = 11
    echo "caught #" & e.Number & ": " & e.Description
Finally
    echo "finally: " & c.Describe()
End Try
