Attribute VB_Name = "m52_Cbom"
'{GP:5}
'{Ep:cBom}
'{Caption:生成BOM}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub cBom()
     If pdm Is Nothing Then
   
          Set pdm = New class_PDM
          End If
   If gws Is Nothing Then
     Set xlm = New Class_XLM
     End If
     
     If gPrd Is Nothing Then
          pdm.defgprd
     End If
     
    
          Set iprd = gPrd
     counter = 1
          
          
          If Not iprd Is Nothing Then
          xlm.inject_bom pdm.recurPrd(iprd, 1)
     End If
     
     Set iprd = Nothing
     xlm.freesheet
End Sub
 


