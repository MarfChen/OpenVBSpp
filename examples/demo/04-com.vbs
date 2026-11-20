' 4) Real COM on Windows: the com-support package ships with the runtime,
'    so CreateObject works without installing pywin32 or importing anything.
' 中文：Windows 上的真实 COM——com-support 随包分发，无需装 pywin32、无需 import。

On Error Resume Next

Dim shell, osName
Set shell = CreateObject("WScript.Shell")
If Err.Number <> 0 Then
    echo "COM unavailable here: " & Err.Description
Else
    osName = shell.ExpandEnvironmentStrings("%OS%")
    echo "WScript.Shell OK -> %OS% = " & osName
    Set shell = Nothing
End If

Dim wmi, rows, row
Set wmi = CreateObject("WbemScripting.SWbemLocator")
If Err.Number = 0 Then
    Set rows = wmi.ConnectServer(".", "root\cimv2").ExecQuery( _
        "SELECT Caption FROM Win32_OperatingSystem")
    For Each row In rows
        echo "WMI -> " & row.Caption
    Next
End If
