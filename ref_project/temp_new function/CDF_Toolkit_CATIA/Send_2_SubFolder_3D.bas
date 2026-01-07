Attribute VB_Name = "Send_2_SubFolder_3D"
Option Explicit
Sub CATMain()
On Error Resume Next
Shell oCATVBA_Folder.Path & "\" & "SendToSubFolders.exe", vbNormalFocus
If Err.Number <> 0 Then
MsgBox "无法调用[创建文件夹保存分总成]！" & vbCrLf & "请确认SendToSubFolders.exe存在于" & oCATVBA_Folder.Path, vbQuestion, "臭豆腐工具箱CATIA版"
End If
End Sub
