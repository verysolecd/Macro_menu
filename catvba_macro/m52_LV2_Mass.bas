Attribute VB_Name = "m52_LV2_Mass"
'{GP:5}
'{Ep:Cal_Mass}
'{Caption:第二级产品重量}
'{ControlTipText:只计算第二级产品重量}
'{BackColor:16744703}

Sub Cal_Mass()
    
    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
    
    If gPrd Is Nothing Then
        Call setgprd
    End If
   Set oprd = gPrd
   Call MassLV2(oprd, 1)
  
    Set oprd = Nothing
End Sub


Function MassLV2(oprd, LV)

If LV <= 3 Then
            Set children = oprd.Products
            If children.Count > 0 Then
                For i = 1 To children.Count
                    Call MassLV2(children.item(i), LV + 1)
                    total = total + children.item(i).ReferenceProduct.UserRefProperties.item("Mass").Value
                Next
                    oprd.ReferenceProduct.UserRefProperties.item("Mass").Value = total
            Else
                    total = oprd.ReferenceProduct.UserRefProperties.item("Mass").Value
            End If
    End If

End Function

