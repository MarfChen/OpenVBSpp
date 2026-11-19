rem 正则 + JSON + 过程值（v0.9.2）

Set re = New RegExp
re.Pattern = "\d{3}-\d{4}"
re.Global = True
echo re.Test("call 555-1234")

Set all = re.Execute("a 111-2222 b 333-4444")
echo "匹配数 " & all.Count
For Each m In all
    echo "  " & m.Value & " @" & m.FirstIndex
Next

text = Json.Stringify(Array(1, "two", True))
echo text
arr = Json.Parse(text)
echo UBound(arr) & ":" & arr(1)

Function Square(n)
    Square = n * n
End Function
f = AddressOf Square
echo f(9)
