Attribute VB_Name = "M31_NewPN"
'Attribute VB_Name = "m30_NewPn"
'{GP:3}
'{Ep:CATMain}
'{Caption:项目前缀}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}

Private prj
Sub CATMain()

If Not KCL.CanExecute("ProductDocument") Then Exit Sub
Set rootprd = CATIA.ActiveDocument.product
If Not rootprd Is Nothing Then
 Dim imsg
          imsg = "请输入你的项目名称"
        prj = KCL.GetInput(imsg)
        If prj = "" Then
            Exit Sub
        End If
    Call rePn(rootprd)
Else
 Exit Sub
End If
End Sub

Sub rePn(oprd)
    pn = oprd.PartNumber
    purePN = KCL.straf1st(pn, "_")
    oprd.PartNumber = prj & "_" & purePN
    For Each product In oprd.Products
        Call rePn(product)
        Next
End Sub
