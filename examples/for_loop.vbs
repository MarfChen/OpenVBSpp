rem For ... Next：求和、负步长、Exit For
total = 0
For i = 1 To 5
    total = total + i
Next
echo total

For i = 10 To 1 Step -3
    echo i
Next

For i = 1 To 100
    If i == 4 Then
        echo "stop at 4"
        Exit For
    End If
Next
