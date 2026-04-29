Attribute VB_Name = "M31_NewPN"
'Attribute VB_Name = "m30_NewPn"
'{GP:3}
'{Ep:CATMain}
'{Caption:项目前缀}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}

Private prj
Sub CATMain()
 If Not CanExecute("ProductDocument") Then Exit Sub
Set rootPrd = CATIA.ActiveDocument.Product
If rootPrd.PartNumber = "_Prj_Housing_Asm" Then
 Dim imsg
          imsg = "请输入你的项目名称"
        prj = KCL.GetInput(imsg)
        If prj = "" Then
            Exit Sub
        End If
    Call rePn(rootPrd)
Else
 Exit Sub
End If
End Sub

Sub rePn(oprd)
    pn = oprd.PartNumber
    oprd.PartNumber = prj & "_" & pn
    For Each Product In oprd.Products
        Call rePn(Product)
        Next
End Sub
