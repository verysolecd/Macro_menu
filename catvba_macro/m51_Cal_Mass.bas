Attribute VB_Name = "m51_Cal_Mass"
'{GP:5}
'{Ep:Cal_Mass}
'{Caption:��������}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}
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
                Call pdm.Assmass(gPrd)
            End If
            
   If Error.Number <> 0 Then
        MsgBox "�������,��ȷ�����ģ���Ƿ�Ӧ�ã�" & Err.Description, vbCritical
    Else
            MsgBox "�����Ѽ���"
    End If


End Sub
