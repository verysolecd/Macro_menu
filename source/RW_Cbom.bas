Attribute VB_Name = "RW_Cbom"
'{GP:1}
'{Ep:cBom}
'{Caption:生成BOM}
'{ControlTipText:一键生成带有截图的BOM}
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
      Set iprd = gPrd
     If iprd Is Nothing Then Exit Sub
     Call Cal_Mass2
     counter = 1
     LV = 1
     
     Dim idx, idcol
  
        Dim tmpData():  tmpData = pdm.recurInfoPrd(iprd, LV)
               
        ReDim resultAry(1 To UBound(tmpData, 1), 1 To UBound(tmpData, 2) + 2)
      
        For i = 1 To UBound(tmpData, 1)
             For j = 1 To UBound(resultAry, 2)
               Select Case j
                    Case 1: resultAry(i, j) = i
                    Case Else: resultAry(i, j) = tmpData(i, (j - 2))
               End Select
             Next j
        Next i
        
        If gws Is Nothing Then
           Set xlm = New Class_XLM
        End If
        
        
      idcol = Array(0, 1, 3, 5, 7, 9, 13) ' 目标列号, 0号元素不占位置
      idx = Array(0, 1, 2, 3, 4, 5, 6, 9, 7, 10, 7)  ' 需提取属性索引（0-based)
      xlm.inject_bom resultAry, idcol, idx
      
      
          Call Capme
            col_pn = 3
            col_pic = 6
          Call xlm.inject_pic(gPic_Path, col_pn, col_pic)
          KCL.ClearDir (gPic_Path)
          
        Call xlm.xlshow
      
     
  
   Set iprd = Nothing
   gPic_Path = ""
   xlm.freesheet
End Sub
