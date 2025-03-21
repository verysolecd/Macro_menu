Attribute VB_Name = "m44_rvPrd"
'{GP:4}
'{Ep:rvme}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub rvme()
     If Not gPrd Is Nothing Then
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
'---------遍历修改产品及子产品
        Dim oPrd: Set oPrd = gPrd
        xlm.extract_data currRow, pdm.infoPrd(Prd2Read)
        Dim children
        Set children = Prd2Read.Products
        For i = 1 To children.Count
         currRow = i + 2
         xlm.inject_data currRow, pdm.infoPrd(children.item(i))
        Next
        Set Prd2Read = Nothing
    Else
        MsgBox "请先选择产品，程序将退出"
        Exit Sub
     End If
On Error GoTo 0
End Sub


