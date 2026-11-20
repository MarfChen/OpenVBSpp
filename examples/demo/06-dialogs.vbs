' Optional seventh take: real MsgBox / InputBox dialogs.
' 可选的第七段：真实的 MsgBox / InputBox 对话框。
'
' This script blocks on purpose: press a button in the dialog, then type an answer.
' Set OPENVBS_UI=gui first (dialogs otherwise fall back to console stand-ins).
' 这个脚本会特意暂停：先在对话框里点按钮，再输入答案。
' 先设置 OPENVBS_UI=gui（否则退回控制台替身）。

Dim answer, typed

answer = MsgBox("Save the report before closing?", vbYesNoCancel + vbQuestion, "OpenVBS++")

If answer = vbYes Then
  echo "user picked: Yes (" & answer & ")"
ElseIf answer = vbNo Then
  echo "user picked: No (" & answer & ")"
Else
  echo "user picked: Cancel (" & answer & ")"
End If

typed = InputBox("Report name:", "OpenVBS++", "quarterly-report")
echo "typed: [" & typed & "]"
