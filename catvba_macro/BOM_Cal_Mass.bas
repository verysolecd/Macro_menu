Attribute VB_Name = "BOM_Cal_Mass"
'{GP:5}
'{Ep:Cal_Mass}
'{Caption:��������}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:}

Sub Cal_Mass()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub

    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
   On Error Resume Next
            If Not gPrd Is Nothing Then
                Call pdm.Assmass(gPrd)
            Else
                Call setgprd
                Err.Clear
                Call pdm.Assmass(gPrd)
            End If
            
   If Err.Number > 0 Then
        MsgBox "�������,��ȷ�����ģ���Ƿ�Ӧ�ã�" & Err.Description, vbCritical
   Else
            MsgBox "�����Ѽ���"
    End If


End Sub
