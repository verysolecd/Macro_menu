Attribute VB_Name = "CAT_Filepath"
'{GP:7}
'{Ep:gotoDocpath}
'{Caption:当前文件夹}
'{ControlTipText:打开当前活动产品所在的文件夹}
'{BackColor: }
Private Const mdlname As String = "CAT_Filepath"
Sub gotoDocpath()
   Dim oDoc: Set oDoc = CATIA.ActiveDocument
   Dim opath: opath = IIf(oDoc.path = "", "", oDoc.FullName)
   Call KCL.SmartOPenPath(opath)
End Sub

