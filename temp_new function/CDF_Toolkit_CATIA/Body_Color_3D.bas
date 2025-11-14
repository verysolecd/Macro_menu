Attribute VB_Name = "Body_Color_3D"
Sub CATMain()
IntCATIA
On Error Resume Next
If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If
BodyColor.Show vbModeless
BodyColor.Left = BodyColor.Left * 2 - 12

End Sub
