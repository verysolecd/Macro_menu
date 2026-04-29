Attribute VB_Name = "Batch_Rename_3D"
Option Explicit
Sub CATMain()
        
IntCATIA
If TypeName(oActDoc) <> "ProductDocument" Then
        MsgBox "当前文件不是CATIA产品 !" & vbCrLf & "此命令只能在产品文件中运行!"
        Exit Sub

End If

Batch_ReName.Show vbModeless
Batch_ReName.Left = Batch_ReName.Left * 2
End Sub


