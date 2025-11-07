Attribute VB_Name = "RW_Cbom"
'{GP:1}
'{Ep:cBom}
'{Caption:生成BOM}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub cBom()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
     If pdm Is Nothing Then
          Set pdm = New class_PDM
     End If
     If gPrd Is Nothing Then
        Set gPrd = pdm.defgprd()
        Set ProductObserver.CurrentProduct = gPrd ' 这会自动触发事件
      End If
      
    If gws Is Nothing Then
     Set xlm = New Class_XLM
    End If
        Call Cal_Mass
    
      Set iprd = gPrd
      
            counter = 1
          If Not iprd Is Nothing Then
          counter = 1
'           xlm.inject_bom pdm.collsAttPrd(iprd, 1)
''              idx = Array(0, 5, 4, 3, 2, 8, 56, 23, 56) ' 需提取的属性索引（0-based）
''                 idcols = Array(0, 1, 3, 5, 7) ' 目标列号, 0号元素不占位置
''
''          Dim mapping
''                mapping = Array(0, 0, 0, 1, 2, 3, 4, 9, 7, 0, 5, 8, 0, 5, 6, 0, 0, 0)
''
''     Call xlm.inject_ary(pdm.collsAttPrd(iprd, 1), idx, idy)
          
          xlm.inject_bom pdm.recurPrd(iprd, 1)
     End If
     Set iprd = Nothing
     xlm.freesheet
     
End Sub


