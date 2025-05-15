Attribute VB_Name = "M31_NewPN2"
'Attribute VB_Name = "m30_NewPn2"
'{GP:3}
'{Ep:CATMain}
'{Caption:件号后缀}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}
Private odate
Sub CATMain()
    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
    Set oprd = pdm.selPrd()
    If oprd Is Nothing Then
        MsgBox "没有选择产品"
    Else
        Dim imsg
              imsg = "请输入后缀"
            odate = KCL.GetInput(imsg)
            If odate = "" Then
                Exit Sub
            End If
        Call rePn(oprd)
    End If
End Sub


Sub rePn(oprd)
    pn = oprd.PartNumber
    oprd.PartNumber = pn & "_" & odate
    For Each Product In oprd.Products
        Call rePn(Product)
        Next
End Sub
