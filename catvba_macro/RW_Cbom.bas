Attribute VB_Name = "RW_Cbom"
'{GP:1}
'{Ep:cBom}
'{Caption:����BOM}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}

Sub cBom()
If Not KCL.CanExecute("ProductDocument") Then Exit Sub
     If pdm Is Nothing Then
          Set pdm = New class_PDM
     End If
     If gPrd Is Nothing Then
    
     Set gPrd = pdm.defgprd()
    Set ProductObserver.CurrentProduct = gPrd ' ����Զ������¼�
      End If
      
    If gws Is Nothing Then
     Set xlm = New Class_XLM
    End If
      Set iprd = gPrd
            counter = 1
          If Not iprd Is Nothing Then
          xlm.inject_bom pdm.recurPrd(iprd, 1)
     End If
     Set iprd = Nothing
     xlm.freesheet
     
End Sub


