rem Option Strict：算术语义收紧 —— 文本操作数不再被转换，而是报类型不匹配
rem Option Strict: arithmetic rejects text operands instead of converting them
Option Strict

echo 1 + 2
echo 10 / 4
echo "a" & 1
echo IsNull(Null + 1)

On Error Resume Next
x = "1" + "2"
echo Err.Number
Err.Clear
x = "2" * 3
echo Err.Number
