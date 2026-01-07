Attribute VB_Name = "Body_Hide_3D"
Sub CATMain()
IntCATIA
If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If
BodyColor.HideBodies
End Sub
