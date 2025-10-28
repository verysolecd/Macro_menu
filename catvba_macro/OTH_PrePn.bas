Attribute VB_Name = "OTH_PrePn"
'Attribute VB_Name = "m30_PrePn"
'{GP:6}
'{Ep:CATMain}
'{Caption:项目前缀}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}

Private prj
Sub CATMain()

If Not KCL.CanExecute("ProductDocument") Then Exit Sub
Set rootprd = CATIA.ActiveDocument.Product
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
    For Each Product In oprd.Products
        Call rePn(Product)
        Next
End Sub
