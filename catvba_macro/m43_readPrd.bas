Attribute VB_Name = "m43_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:��ȡ����}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}

Sub readPrd()
    Dim xlm, pdm, ws
    Set xlm = New Class_XLM
    Set pdm = New class_PDM
    Set ws = gws
    
'---------��ȡ���޸Ĳ�Ʒ
    On Error resume Next
        Set g_Prd2wt = pdm.catchPrd()
        If g_Prd2wt Is Nothing Then
            MsgBox "δѡ���Ʒ"
        End If
        if Err.Number <> 0 Then
            msg box "δѡ���Ʒ"
    On Error GoTo 0
    On Error resume Next
    Dim currRow: currRow = 2
 '---------�����޸Ĳ�Ʒ���Ӳ�Ʒ
    Dim Prd2Read: Set Prd2Read = g_Prd2wt
        xlm.inject_data currRow, pdm.infoPrd(Prd2Read), "rv"
        
    Dim children
    Set children = Prd2Read.Products
        For i = 1 To children.Count
         currRow = i + 2
         xlm.inject_data currRow, pdm.infoPrd(children.Item(i)), "rv"
        Next
    Set Prd2Read = Nothing
     On Error GoTo 0   
    
ErrHandler:
    Select Case Err.Number
    Case 429 ' CATIAδ����
    MsgBox "�޷�����CATIA����ȷ������������", vbCritical
    Case Else
    MsgBox "���� " & Err.Number & ": " & Err.Description, vbCritical
    End Select
    
End Sub
