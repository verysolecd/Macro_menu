Attribute VB_Name = "m43_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:��ȡ����}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}


Sub readPrd()
'excel�����catia�������ʼ��
Dim xlm, pdm
Set xlm = New Class_XLM
Set pdm = New class_PDM

'---------��ȡ���޸Ĳ�Ʒ
On Error Resume Next
     pdm.catchgPrd
     If Not gprd Is Nothing Then
     gprd.ApplyWorkMode (3)
     Dim currRow: currRow = 2
'---------�����޸Ĳ�Ʒ���Ӳ�Ʒ
     Dim Prd2Read: Set Prd2Read = gprd
     xlm.inject_data currRow, pdm.infoPrd(Prd2Read), "rv"
     Dim children
     Set children = Prd2Read.Products
     For i = 1 To children.Count
      currRow = i + 2
      xlm.inject_data currRow, pdm.infoPrd(children.Item(i)), "rv"
     Next
     Set Prd2Read = Nothing
     xlm.xlApp.Visible = True
 Else
     MsgBox "δѡ���Ʒ���˳�"
     Exit Sub
     End If
On Error GoTo 0
End Sub
