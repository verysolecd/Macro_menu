Attribute VB_Name = "m40_setgprd"
'{GP:4}
'{Ep:setgprd}
'{Caption:ѡ���Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}


Sub setgprd()
'            Dim oSel, status, iType(0)
'            CATIA.ActiveWindow.WindowState = 0
'            MsgBox "��ѡ��Ҫ��ȡ�Ĳ�Ʒ"
'
'            Set oSel = CATIA.Activedocument.Selection
'            oSel.Clear
'            iType(0) = "Product"
'            If oSel.Count2 = 0 Then
'                status = oSel.SelectElement2(iType, "��ѡ��Ҫ��ȡ�Ĳ�Ʒ", False)
'                status = oSel.SelectElement3(iType, "��ѡ��Ҫ��ȡ�Ĳ�Ʒ", True, 2, False)
'            End If
'            If status = "Cancel" Then
'                Exit Function
'            End If
'            If status = "Normal" And oSel.Count2 = 1 Then
'                    Set selPrd = oSel.Item(1).LeafProduct.ReferenceProduct
'                    oSel.Clear
'            Else
'                MsgBox "��ѡ���ҽ�ѡ��һ����Ʒ"
'                Exit Function
'                oSel.Clear
'            End If

    Dim xlm, pdm
    Set pdm = New class_PDM
    On Error Resume Next
     Set gPrd = pdm.defgprd()
        If gPrd Is Nothing Then
        MsgBox "���˳������򽫽���"
        Exit Sub
        Else
        MsgBox "��������Ʒ�Ѿ�ѡ��" & gPrd.PartNumber
        End If
End Sub
