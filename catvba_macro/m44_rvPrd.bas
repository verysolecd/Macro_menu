'{GP:4}
'{Ep:rvme}
'{Caption:修改产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub rvme()

     If Not gPrd Is Nothing Then
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
        
'---------遍历修改产品及子产品   Set data =
        Dim Prd2rv
        Set Prd2rv = gPrd
        
        Dim odata As Variant
        odata = xlm.extract_data(currRow)

        
        Call pdm.modatt(Prd2rv, odata)
        Dim children
        Set children = Prd2rv.Products
        For i = 1 To children.Count
         currRow = currRow + 1
            odata = xlm.extract_data(currRow)
           Call pdm.modatt(children.item(i), odata)
        Next
        Set Prd2rv = Nothing
        MsgBox "已经修改产品"
    Else
        MsgBox "请先选择产品，程序将退出"
        Exit Sub
     End If
    On Error GoTo 0


End Sub
