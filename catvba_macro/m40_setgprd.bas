Attribute VB_Name = "m40_setgprd"
'{GP:4}
'{Ep:setgprd}
'{Caption:ѡ���Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}

Sub setgprd()

    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
    
    Dim oprd
    
    imsg = MsgBox("���ǡ�ѡ���Ʒ�����񡱶�ȡ����Ʒ����ȡ���˳���", vbYesNoCancel + vbDefaultButton2, "��ѡ���Ʒ")

        Select Case imsg
            Case 7 '===ѡ�񡰷�====
                Set oprd = CATIA.ActiveDocument.Product
            Case 2:
                Exit Sub '===ѡ��ȡ����====
            Case 6  '===ѡ���ǡ�,���в�Ʒѡ��====
                On Error Resume Next
                    Set oprd = pdm.selPrd()
                    If Err.Number <> 0 Then
                        Err.Clear
                        Exit Sub
                    End If
                On Error GoTo 0
        End Select
        
         Set gPrd = oprd
         Set oprd = Nothing
         
        If Not gPrd Is Nothing Then
           imsg = "��ѡ��Ĳ�Ʒ��" & gPrd.PartNumber
            MsgBox imsg
        Else
             MsgBox "���˳������򽫽���"
        End If
End Sub
