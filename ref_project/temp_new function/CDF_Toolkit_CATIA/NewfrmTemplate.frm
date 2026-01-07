VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewfrmTemplate 
   Caption         =   "从模板新建 | 臭豆腐工具箱CATIA版"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   OleObjectBlob   =   "NewfrmTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewfrmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lstTemplate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim sTempfile, i
sTempfile = vbNullString
For i = 0 To lstTemplate.ListCount - 1
    If lstTemplate.Selected(i) = True Then
        sTempfile = lstTemplate.List(i)
        Exit For
    End If
Next

If sTempfile <> vbNullString Then
GB_Frame.AddDoc2Tree oCATVBA_Folder("Template\Start_Part") & "\" & sTempfile
End If
End Sub

Private Sub UserForm_Initialize()
'1.扫描所有..\Template\Start_Part文件夹下的CATPart,CATDrawing和CATProduct文件，并将文件名放入数组
'2.数组内容添加执行按钮
Dim f, fso, fname
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In oCATVBA_Folder("Template\Start_Part").Files
        fname = f.Name
        If Right(f.Name, 8) = ".CATPart" Then
           lstTemplate.AddItem f.Name
        ElseIf Right(f.Name, 11) = ".CATProduct" Then
           lstTemplate.AddItem f.Name
        ElseIf Right(f.Name, 11) = ".CATDrawing" Then
           lstTemplate.AddItem f.Name
        End If
    Next
End Sub

