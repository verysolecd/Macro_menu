Attribute VB_Name = "MP_ReName_Export_3D"
Option Explicit
Sub CATMain()
       
        
 IntCATIA
        
'MsgBox TypeName(oActDoc)
    If TypeName(oActDoc) <> "ProcessDocument" Then
    MsgBox "此命令只能在加工模块中运行！"
    Exit Sub
    End If

        
MProg.Show vbModeless
MProg.Left = MProg.Left * 2
MProg.Repaint
CATIA.RefreshDisplay = True
End Sub

