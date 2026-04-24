Attribute VB_Name = "CAT_Filepath"
'{GP:7}
'{Ep:openfilepath}
'{Caption:当前文件夹}
'{ControlTipText:打开当前活动产品所在的文件夹}
'{BackColor: }
Private Quick
Private Const mdlname As String = "CAT_Filepath"

Sub openfilepath()
On Error Resume Next
   Dim oDoc: Set oDoc = CATIA.ActiveDocument
   Dim opath: opath = IIf(oDoc.path = "", "", oDoc.FullName)
   KCL.SmartOPenPath opath
 On Error GoTo 0
End Sub

