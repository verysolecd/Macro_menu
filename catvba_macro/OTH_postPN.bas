Attribute VB_Name = "OTH_postPN"
'Attribute VB_Name = "m30_postPN"
'{GP:6}
'{Ep:CATMain}
'{Caption:���ź�׺}
'{ControlTipText:Ϊ���������������Ŀǰ׺}
'{BackColor:}
Private oSuffix
Sub CATMain()
If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
    Set oprd = pdm.selPrd()
    If oprd Is Nothing Then
        MsgBox "û��ѡ���Ʒ"
    Else
        Dim imsg
              imsg = "�������׺"
            oSuffix = KCL.GetInput(imsg)
            If oSuffix = "" Then
                Exit Sub
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

