Attribute VB_Name = "m42_rvPrd"
'{GP:4}
'{Ep:rvme}
'{Caption:�޸Ĳ�Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor: }

Sub rvme()

     If Not gPrd Is Nothing Then
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
        
'---------�����޸Ĳ�Ʒ���Ӳ�Ʒ   Set data =
        Dim Prd2rv
        Set Prd2rv = gPrd
        
        Dim odata As Variant
        odata = xlm.extract_data(2)

        
        Call pdm.modatt(Prd2rv, odata)
        Dim children
        Set children = Prd2rv.Products
        If children.Count > 0 Then
            For i = 1 To children.Count
            currRow = currRow + 1
            Dim sdata As Variant
            sdata = xlm.extract_data(currRow)
           Call pdm.modatt(children.item(i), sdata)
        Next
        End If
        Set Prd2rv = Nothing
        MsgBox "�Ѿ��޸Ĳ�Ʒ"
    Else
        MsgBox "����ѡ���Ʒ�������˳�"
        Exit Sub
     End If
    On Error GoTo 0


End Sub


Public Function extract_data(indRow)

    Dim iCols
    iCols = Array(0, 2, 4, 6, 8, 10, 12)

    
     Set ws = xlApp.ActiveSheet
     
    
    Dim temparr As Variant
    
    temparr = ws.Rows(indRow).Resize(1, 14).Value
    

    Dim outputArr As Variant
    Dim j As Long
    ReDim outputArr(1 To UBound(iCols))
    For j = 1 To UBound(iCols)
         outputArr(j) = ""
         If IsEmpty(temparr(1, iCols(j))) = False Then
         outputArr(j) = temparr(1, iCols(j))
         End If
    Next
    extract_data = outputArr ' ������ȡ������
End Function




