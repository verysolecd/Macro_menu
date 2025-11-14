VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReplaceNum 
   Caption         =   "替换部分字符"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2310
   OleObjectBlob   =   "ReplaceNum.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReplaceNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddPreSuf_Click()
'On Error Resume Next
'*****************运行环境检查*********************
If txtAddPre.Text = "" And txtAddSuf.Text = "" Then
MsgBox "增加的前缀和后缀同时为空!"
Exit Sub
End If

'*****************运行环境检查*********************

Dim op As ManufacturingSetup
Dim MPs As MfgActivities
Set op = MProg.Sel("ManufacturingSetup")
If Err.Number <> 0 Then
Exit Sub
End If
Set MPs = op.Programs

Dim i
For i = 1 To MPs.Count
If MPs.GetElement(i).Active Then
MPs.GetElement(i).Name = txtAddPre.Text & MPs.GetElement(i).Name & txtAddSuf.Text
End If
Next

MsgBox "操作完成！请 Ctrl + S 保存修改"
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
End Sub

Private Sub cmdReName2_Click()
On Error Resume Next
'*****************运行环境检查*********************
If txtFind = "" Then
MsgBox lblFind.Caption & "不能为空"
Exit Sub
End If

If IsNumeric(txtStart) = False Then
MsgBox "填写错误"
Exit Sub
End If
If CInt(txtStart) < 1 Then
MsgBox "至少要从第1个字符开始"
Exit Sub
End If

If IsNumeric(txtCount) = False Or CInt(txtCount) = 0 Then
MsgBox "填写错误"
Exit Sub
End If
'*****************运行环境检查*********************

Dim casesens
If chkCASE.Value = True Then
casesens = 1
Else
casesens = 0
End If

Dim op As ManufacturingSetup
Dim MPs As MfgActivities
Set op = MProg.Sel("ManufacturingSetup")
If Err.Number <> 0 Then
Exit Sub
End If
Set MPs = op.Programs

Dim i
For i = 1 To MPs.Count
If MPs.GetElement(i).Active Then
MPs.GetElement(i).Name = Replace(MPs.GetElement(i).Name, RTrim(txtFind), RTrim(txtReplaceWith), CInt(txtStart), CInt(txtCount), casesens)
End If
Next

MsgBox "操作完成！请 Ctrl + S 保存修改"

End Sub


Private Sub txtCount_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入正整数！"
            
    End Select
End Sub

Private Sub txtStart_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入正整数！"
            
    End Select
End Sub


Private Sub UserForm_Terminate()
MProg.Show vbModeless
MProg.Left = MProg.Left * 2
End Sub
