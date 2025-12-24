Attribute VB_Name = "global_var"
Public gPrd As Object
Public rootDoc
Public rootPrd  As Object  '全局产品obj
Public startrow, lastrow  '全局excel行定义
Public xlAPP As Object  '全局excelcom组件
Public gwb As Object
Public gws  As Object
Public pdm As Object
Public xlm As Object
Public allPN As Object
Public counter As Integer
Public Const gfn As Integer = 400
Public Cls_PrdOB As New Cls_PrdOB

Public gPic_Path


Sub clearall()

End Sub

'Dim btn, bTitle, bResult
'imsg = "将备份到" & bckpath "您确认吗"

'btn = vbYesNo + vbExclamation
'bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)

'Select Case bResult
'Case 7: Exit Sub '===选择“否”====
'Case 6  '===选择“是”====
'Case 2  '===选择“取消”====



