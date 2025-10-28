Attribute VB_Name = "BOM_LV2_Mass"
'{GP:5}
'{Ep:L2Mass}
'{Caption:迭代重量L2}
'{ControlTipText:只计算第二级产品重量}
'{BackColor:}

Sub L2Mass()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
    If gPrd Is Nothing Then
        Call setgprd
    End If
   Set oprd = gPrd
   Call cal2mass(oprd, 1)
    Set oprd = Nothing
End Sub
Function cal2mass(oprd, LV)
If LV <= 3 Then
            Set children = oprd.Products
            If children.Count > 0 Then
                For i = 1 To children.Count
                    Call cal2mass(children.item(i), LV + 1)
                    total = total + children.item(i).ReferenceProduct.UserRefProperties.item("Mass").Value
                Next
                    oprd.ReferenceProduct.UserRefProperties.item("Mass").Value = total
            Else
                    total = oprd.ReferenceProduct.UserRefProperties.item("Mass").Value
            End If
    End If
End Function
