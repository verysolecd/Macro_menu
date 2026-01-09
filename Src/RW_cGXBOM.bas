Attribute VB_Name = "RW_cGXBOM"
'{GP:1}
'{Ep:cgxBom}
'{Caption:GXBOM}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub cgxBom()
If Not KCL.CanExecute("ProductDocument") Then Exit Sub
If pdm Is Nothing Then Set pdm = New Cls_PDM
If pdm.CurrentProduct Is Nothing Then Set pdm.CurrentProduct = pdm.getiPrd()
g_counter = 1
Set iprd = pdm.CurrentProduct
If Not iprd Is Nothing Then
      If gws Is Nothing Then Set xlm = New Cls_XLM
      xlm.inject_gxbom pdm.gxBom(iprd, 1)
End If
Set iprd = Nothing
xlm.xlshow
   xlm.freesheet
     
End Sub



