Attribute VB_Name = "M31_NewPN2"
'Attribute VB_Name = "m30_NewPn2"
'{GP:3}
'{Ep:CATMain}
'{Caption:���ź�׺}
'{ControlTipText:Ϊ���������������Ŀǰ׺}
'{BackColor:}
Private odate
Sub CATMain()
    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
    Set oprd = pdm.selPrd()
    If oprd Is Nothing Then
        MsgBox "û��ѡ���Ʒ"
    Else
        Dim imsg
              imsg = "�������׺"
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
