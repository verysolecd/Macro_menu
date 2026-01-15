Attribute VB_Name = "BOM_cMassLV2"
'{GP:2}
'{Ep:L2Mass}
'{Caption:迭代重量L2}
'{ControlTipText:只计算第二级产品重量}
'{BackColor:}

Sub L2Mass()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    If pdm.CurrentProduct Is Nothing Then Call setgprd
   Set oprd = pdm.CurrentProduct
   Call cal2mass(oprd, 1)
    Set oprd = Nothing
End Sub
Function cal2mass(oprd, lv)
        If lv <= 3 Then
                Set children = oprd.Products
                If children.count > 0 Then
                    For i = 1 To children.count
                        Call cal2mass(children.item(i), lv + 1)
                        total = total + children.item(i).ReferenceProduct.UserRefProperties.item("Mass").value
                    Next
                        oprd.ReferenceProduct.UserRefProperties.item("Mass").value = total
                Else
                        total = oprd.ReferenceProduct.UserRefProperties.item("Mass").value
                End If
        End If
End Function
