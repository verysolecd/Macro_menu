Attribute VB_Name = "m43_readPrd"
'Attribute VB_Name = "ReadPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:��ȡ����}
'{ControlTipText:��ȡ��������Ʒ}
'{BackColor:16744703}

Sub readPrd()
    If pdm Is Nothing Then
     Set pdm = New class_PDM
    End If

    If gws Is Nothing Then
         Set xlm = New Class_XLM
    End If

 '---------��ȡ���޸Ĳ�Ʒ
    If Not gPrd Is Nothing Then
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
        
        
'---------�����޸Ĳ�Ʒ���Ӳ�Ʒ

        Dim Prd2Read: Set Prd2Read = gPrd
        xlm.inject_data currRow, pdm.infoPrd(Prd2Read)        
        Dim children
        Set children = Prd2Read.Products
        For i = 1 To children.Count
        currRow = i + 2
        xlm.inject_data currRow, pdm.infoPrd(children.item(i))
        
        Next
        Set Prd2Read = Nothing
    Else
        MsgBox "����ѡ���Ʒ�������˳�"
        Exit Sub
    
    End If
End Sub
