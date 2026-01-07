Attribute VB_Name = "RW_cGXBOM"
'{GP:1}
'{Ep:cgxBom}
'{Caption:GXBOM}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub cgxBom()
If Not KCL.CanExecute("ProductDocument") Then Exit Sub
     If pdm Is Nothing Then
          Set pdm = New Cls_PDM
     End If
     If gPrd Is Nothing Then
    
     Set gPrd = pdm.getiPrd()
    Set Cls_PrdOB.CurrentProduct = gPrd ' 这会自动触发事件
      End If
      
    If gws Is Nothing Then
     Set xlm = New Cls_XLM
    End If
      Set iprd = gPrd
            counter = 1
          If Not iprd Is Nothing Then
          xlm.inject_gxbom pdm.gxBom(iprd, 1)
     End If
     Set iprd = Nothing
     xlm.freesheet
     
End Sub



