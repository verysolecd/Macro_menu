Attribute VB_Name = "m44_rvPrd"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:rvme}
'{Caption:ѡ���Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}
Sub rvme()   
     If Not gprd Is Nothing Then
        gprd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
'---------�����޸Ĳ�Ʒ���Ӳ�Ʒ
        Dim oprd: Set oprd = gprd
        xlm.extract_data currRow, pdm.infoPrd(Prd2Read)
        Dim children
        Set children = Prd2Read.Products
        For i = 1 To children.Count
         currRow = i + 2
         xlm.inject_data currRow, pdm.infoPrd(children.Item(i))
        Next
        Set Prd2Read = Nothing
    Else
        MsgBox "����ѡ���Ʒ�������˳�"
        Exit Sub
     End If
On Error GoTo 0
End Sub


