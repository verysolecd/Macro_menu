Attribute VB_Name = "m52_Cbom"
'Attribute VB_Name = "m5_Cbom"
'{GP:5}
'{Ep:cBom}
'{Caption:ѡ���Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}
Sub cBom()
     Dim xlm As New Class_XLM
     Dim pdm As New class_PDM
     Dim bPrd
     If gprd Is Nothing Then
          pdm.catchgPrd
     Else
          Set iPrd = gprd
          xlm.inject_bom pdm.recurPrd(iPrd, 0)
     End If
     Set iPrd = Nothing
End Sub
 


