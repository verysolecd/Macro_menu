Attribute VB_Name = "RW_LV2_Mass"
'{GP:}
'{Ep:L2Mass}
'{Caption:迭代重量L2}
'{ControlTipText:只计算第二级产品重量}
'{BackColor:}

Sub L2Mass()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then
        Set pdm = New Cls_PDM
    End If
    If gPrd Is Nothing Then
        Call setgprd
    End If
   Set oPrd = gPrd
   Call cal2mass(oPrd, 1)
    Set oPrd = Nothing
End Sub
Function cal2mass(oPrd, LV)
If LV <= 3 Then
            Set children = oPrd.Products
            If children.count > 0 Then
                For i = 1 To children.count
                    Call cal2mass(children.item(i), LV + 1)
                    total = total + children.item(i).ReferenceProduct.UserRefProperties.item("Mass").value
                Next
                    oPrd.ReferenceProduct.UserRefProperties.item("Mass").value = total
            Else
                    total = oPrd.ReferenceProduct.UserRefProperties.item("Mass").value
            End If
    End If
End Function
