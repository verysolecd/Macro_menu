Attribute VB_Name = "CNC_Prog_List2"
Option Explicit
Sub CATMain()
On Error Resume Next
Shell oCATVBA_Folder.Path & "\" & "CNCtools.exe", vbNormalFocus
If Err.Number <> 0 Then
MsgBox "无法调用CNCtools！" & vbCrLf & "请确认CNCtools.exe存在于" & oCATVBA_Folder.Path, vbQuestion, "臭豆腐工具箱CATIA版"
End If
End Sub
