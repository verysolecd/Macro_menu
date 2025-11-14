Attribute VB_Name = "Body_ShowReverse_3D"
Option Explicit
Sub CATMain()
       
IntCATIA
'Set alreadysh = CreateObject("Scripting.Dictionary")
On Error Resume Next
If TypeName(oActDoc) = "ProductDocument" Then
Body_ShowAll_3D.ShowallBodies oActDoc.Product, True
ElseIf TypeName(oActDoc) = "PartDocument" Then
Body_ShowAll_3D.ShowallBodies2 oActDoc.Part, True
Else
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
End If
'alreadysh.RemoveAll
End Sub
