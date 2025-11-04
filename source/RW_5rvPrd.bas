Attribute VB_Name = "RW_5rvPrd"
'{GP:1}
'{Ep:rvme}
'{Caption:修改产品属性}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor: }

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
            
        If children.count > 0 Then
            For i = 1 To children.count
                currRow = currRow + 1
                Dim sdata As Variant
                sdata = xlm.extract_data(currRow)
                Call pdm.modatt(children.item(i), sdata)
            Next
        End If
            Set Prd2rv = Nothing
        MsgBox "已经修改产品"
    Else
        MsgBox "请先选择产品，程序将退出"
        Exit Sub
     End If
    On Error GoTo 0
End Sub





