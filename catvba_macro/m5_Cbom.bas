Attribute VB_Name = "m5_Cbom"
'Attribute VB_Name = "m5_Cbom"
'{GP:5}
'{Ep:recurPrd}
'{Caption:?????}
'{ControlTipText:?????????????????}
'{BackColor:16744703}


Sub CATMain()
    Dim xlm As New Class_XLM
    Dim pdm As New class_PDM
    Dim oPrd As Object    
    ' ��ȡҪ����Ĳ�Ʒ
    Set oPrd = pdm.catchPrd()
    If oPrd Is Nothing Then
        MsgBox "δѡ���Ʒ"
        Exit Sub
    End If
    
    ' ��ȡBOM����
    Dim bomArray As Variant
    bomArray = pdm.recurPrd(oPrd, 0)
    
    ' һ����д��Excel
    Dim lastRow As Long
    lastRow = UBound(bomArray, 1)
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow + 1, UBound(bomArray, 2))).Value = bomArray
    
    ' ����Excel�ɼ�
    xlm.xlApp.Visible = True
End Sub


