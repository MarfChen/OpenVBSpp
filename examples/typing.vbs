rem Option Typing：开启编译期静态类型检查
Option Typing

Dim count -> Integer
Dim ratio -> Double
Dim name -> String

count = 10
ratio = 0.5
name = "openvbs"

Function Twice(n -> Integer) -> Integer
    Twice = n * 2
End Function

echo Twice(count)
echo ratio
echo name
