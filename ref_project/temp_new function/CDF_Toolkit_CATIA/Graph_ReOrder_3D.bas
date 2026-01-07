Attribute VB_Name = "Graph_ReOrder_3D"
Option Explicit
Sub CATMain()

On Error Resume Next
'MsgBox oCATVBA_Folder.Path & "\" & "GraphTreeReOrder.exe"
Shell oCATVBA_Folder.Path & "\" & "GraphTreeReOrder.exe", vbNormalFocus
If Err.Number <> 0 Then
MsgBox "无法调用图形树重新排序！" & vbCrLf & "请确认GraphTreeReOrder.exe存在于" & oCATVBA_Folder.Path, vbQuestion, "臭豆腐工具箱CATIA版"
End If
End Sub
