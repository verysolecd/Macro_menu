Attribute VB_Name = "RW_4readPrd"
'Attribute VB_Name = "ReadPrd"
'{gp:1}
'{Ep:readPrd}
'{Caption:读取产品属性}
'{ControlTipText:读取待操作产品}
'{BackColor: }

Sub readPrd()
 If Not CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then
     Set pdm = New Cls_PDM
    End If
 '---------获取待修改产品 '---------遍历修改产品及子产品
    If gPrd Is Nothing Then
         MsgBox "请先选择产品，程序将退出"
         Exit Sub
    Else
         If gws Is Nothing Then
           Set xlm = New Cls_XLM
           End If
           Dim Prd2Read: Set Prd2Read = gPrd
    End If
        If Not Prd2Read Is Nothing Then
            Dim currRow: currRow = 2
            counter = 1
            Prd2Read.ApplyWorkMode (3)
            idcol = Array(0, 1, 3, 5, 7, 9, 11, 13, 14) ' 目标列号, 0号元素不占位置
            idx = Array(0, 1, 2, 3, 4, 5, 6, 7, 8) ' 对应的属性索引（0-based）
            Dim tmpData(): tmpData = pdm.attLv2Prd(Prd2Read)
            xlm.inject_ary tmpData, idcol, idx
            xlm.setxlHead ("rv")
            xlm.xlshow
                xlAPP.Visible = True
            End If

        Set Prd2Read = Nothing
End Sub
