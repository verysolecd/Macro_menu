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
 

 '---------��ȡ���޸Ĳ�Ʒ '---------�����޸Ĳ�Ʒ���Ӳ�Ʒ
    If gPrd Is Nothing Then
         MsgBox "����ѡ���Ʒ�������˳�"
         Exit Sub
    Else
         If gws Is Nothing Then
           Set xlm = New Class_XLM
         End If
        
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
         
        Dim Prd2Read
        Set Prd2Read = gPrd
        
        If Not Prd2Read Is Nothing Then
            xlm.inject_data currRow, pdm.infoPrd(Prd2Read)
            
            Dim children
            Set children = Prd2Read.Products
            For i = 1 To children.Count
                currRow = i + 2
                xlm.inject_data currRow, pdm.infoPrd(children.item(i))
            Next
        End If
      End If
        Set Prd2Read = Nothing
End Sub
