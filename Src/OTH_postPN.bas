Attribute VB_Name = "OTH_postPN"
'Attribute VB_Name = "m30_postPN"
'{GP:3}
'{Ep:new_pn_suffix}
'{Caption:零件号后缀}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}
Private oSuffix
Private Const mdlname As String = "OTH_postPN"
Sub new_pn_suffix()
If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    Set oPrd = KCL.SelectItem("请选择产品", "Product")
    If oPrd Is Nothing Then
        MsgBox "没有选择产品"
    Else
        Dim imsg
              imsg = "请输入后缀"
            oSuffix = KCL.GetInput(imsg)
            If oSuffix = "" Then
                MsgBox imsg: Exit Sub
            End If
        Call postPn(oPrd)
    End If
End Sub

Sub postPn(oPrd)
    pn = oPrd.partNumber
    oPrd.partNumber = pn & "_" & oSuffix
    For Each Product In oPrd.Products
        Call postPn(Product)
        Next
End Sub

