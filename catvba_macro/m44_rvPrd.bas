'{GP:4}
'{Ep:rvme}
'{Caption:�޸Ĳ�Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}

Sub rvme()

     If Not gPrd Is Nothing Then
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
        
'---------�����޸Ĳ�Ʒ���Ӳ�Ʒ   Set data =
        Dim Prd2rv
        Set Prd2rv = gPrd
        
        Dim odata As Variant
        odata = xlm.extract_data(currRow)

        
        Call pdm.modatt(Prd2rv, odata)
        Dim children
        Set children = Prd2rv.Products
        For i = 1 To children.Count
         currRow = currRow + 1
            odata = xlm.extract_data(currRow)
           Call pdm.modatt(children.item(i), odata)
        Next
        Set Prd2rv = Nothing
        MsgBox "�Ѿ��޸Ĳ�Ʒ"
    Else
        MsgBox "����ѡ���Ʒ�������˳�"
        Exit Sub
     End If
    On Error GoTo 0


End Sub
