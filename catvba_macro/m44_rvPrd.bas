Attribute VB_Name = "m44_rvPrd"
'{GP:4}
'{Ep:rvme}
'{Caption:�޸Ĳ�Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}
Private Prd2rv

Sub rvme()
     If Not gPrd Is Nothing Then
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
'---------�����޸Ĳ�Ʒ���Ӳ�Ʒ   Set data =
        Dim Prd2rv
        Set Prd2rv = gPrd

        Dim data As Variant
       xlm.extract_data (currRow)
        Call pdm.modatt(Prd2rv, data)
        Dim children
        Set children = Prd2rv.Products
        For i = 1 To children.Count
         currRow = currRow + 1
            data = xlm.extract_data(currRow)
           Call pdm.modatt(children.item(i), data)
        Next
        Set Prd2rv = Nothing
        MsgBox "�Ѿ��޸Ĳ�Ʒ"
    Else
        MsgBox "����ѡ���Ʒ�������˳�"
        Exit Sub
     End If
On Error GoTo 0
End Sub



