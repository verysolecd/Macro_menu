Attribute VB_Name = "RW_4readPrd"
'Attribute VB_Name = "ReadPrd"
'{gp:1}
'{Ep:readPrd}
'{Caption:读取产品属性}
'{ControlTipText:读取待操作产品}
'{BackColor: }

Sub readPrd()
    If pdm Is Nothing Then
     Set pdm = New class_PDM
    End If
 '---------获取待修改产品 '---------遍历修改产品及子产品
    If gPrd Is Nothing Then
         MsgBox "请先选择产品，程序将退出"
         Exit Sub
    Else
         If gws Is Nothing Then
           Set xlm = New Class_XLM
           End If
    End If
        Dim currRow: currRow = 2
         counter = 1
        Dim Prd2Read
        Set Prd2Read = gPrd
        If Not Prd2Read Is Nothing Then
            Prd2Read.ApplyWorkMode (3)
        idx = Array(0, 1, 2, 3, 4, 5, 6, 7, 8) ' 需提取的属性索引（0-based）
        idcol = Array(0, 1, 3, 5, 7, 9, 11, 13, 14) ' 目标列号, 0号元素不占位置
        
        Dim idata()
        idata = pdm.attLv2Prd(Prd2Read)
        
            ReDim resultAry(1 To UBound(idata, 1), 1 To UBound(idx))
        For i = 1 To UBound(idata, 1)
             For j = 1 To UBound(resultAry, 2)
               resultAry(i, j) = idata(i, (idx(j)))
             Next j
        Next i
        xlm.inject_ary resultAry, idcol
        End If
          xlAPP.Visible = True
        Set Prd2Read = Nothing
End Sub
