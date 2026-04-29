Attribute VB_Name = "OTH_postPN"
'Attribute VB_Name = "m30_postPN"
'{GP:6}
'{Ep:CATMain}
'{Caption:零件号后缀}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}
Private oSuffix
Sub CATMain()
If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    Set oprd = KCL.SelectItem("请选择产品", "Product")
    If oprd Is Nothing Then
        MsgBox "没有选择产品"
    Else
        Dim imsg
              imsg = "请输入后缀"
            oSuffix = KCL.GetInput(imsg)
            If oSuffix = "" Then
                MsgBox imsg: Exit Sub
            End If
        Call postPn(oprd)
    End If
End Sub

Sub postPn(oprd)
    pn = oprd.PartNumber
    oprd.PartNumber = pn & "_" & oSuffix
    For Each Product In oprd.Products
        Call postPn(Product)
        Next
End Sub

