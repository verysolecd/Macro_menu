Attribute VB_Name = "AutoDrw_2D_3D"
Option Explicit
Sub CATMain()
On Error Resume Next
Shell oCATVBA_Folder.Path & "\" & "AutoDrw.exe", vbNormalFocus
If Err.Number <> 0 Then
MsgBox "无法调用AutoDrw！" & vbCrLf & "请确认AutoDrw.exe存在于" & oCATVBA_Folder.Path, vbQuestion, "臭豆腐工具箱CATIA版"
End If
End Sub

