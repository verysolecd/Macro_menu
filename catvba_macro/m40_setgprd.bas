Attribute VB_Name = "m40_setgprd"
'{GP:4}
'{Ep:setgprd}
'{Caption:ѡ���Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}

Sub setgprd()
    If Not CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If

    Set gPrd = pdm.defgprd()
    Set ProductObserver.CurrentProduct = gPrd ' ����Զ������¼�
        If Not gPrd Is Nothing Then
           imsg = "��ѡ��Ĳ�Ʒ��" & gPrd.PartNumber
            MsgBox imsg
        Else
             MsgBox "���˳������򽫽���"
        End If
End Sub
