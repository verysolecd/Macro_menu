Attribute VB_Name = "RW_Cbom"
'{GP:1}
'{Ep:cBom}
'{Caption:生成BOM}
'{ControlTipText:一键生成带有截图的BOM}
'{BackColor:16744703}

Sub cBom()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
     If pdm Is Nothing Then
          Set pdm = New Cls_PDM
     End If
     
     If gPrd Is Nothing Then
        Set gPrd = pdm.getiPrd()
        Set Cls_PrdOB.CurrentProduct = gPrd ' 这会自动触发事件
     End If
      Set iprd = gPrd
     If iprd Is Nothing Then Exit Sub
     Call Cal_Mass2
     counter = 1
     LV = 1
      If gws Is Nothing Then
           Set xlm = New Cls_XLM
        End If
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
   Dim idx, idcol
      idcol = Array(0, 1, 2, 3, 4, 5, 7, 8, 10, 11, 13) ' 目标列号, 0号元素不占位置
        idx = Array(0, 1, 2, 3, 4, 5, 11, 9, 7, 10, 7)  ' 需提取属性索引（0-based)
      xlm.inject_bom resultAry, idcol, idx
    Dim btn, bTitle, bResult
        CATIA.StartCommand ("* iso")
        bTitle = ""
         imsg = "如要截图到BOM截图，请等待ISO视角调整完毕后点击确认"
        btn = vbYesNo + vbExclamation
        bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
        Select Case bResult
            Case 7: GoTo CleanUp '===选择“否”====
            Case 2: Exit Sub '===选择“取消”====
            Case 6  '===选择“是”====
                Call Capme
                Dim Colpn, colPic
                Colpn = 3: colPic = 6
                Call xlm.inject_pic(gPic_Path, Colpn, colPic)
                GoTo CleanUp
            End Select
CleanUp::
   Call xlm.xlshow
   Set iprd = Nothing
       KCL.ClearDir (gPic_Path)
   gPic_Path = ""
   xlm.freesheet
End Sub
