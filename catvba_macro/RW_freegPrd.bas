Attribute VB_Name = "RW_freegPrd"
'Attribute VB_Name = "selPrd"
'{GP:1}
'{Ep:freegprd}
'{Caption:�ͷŲ�Ʒ}
'{ControlTipText:����������Ʒ���}
'{BackColor:16744703}


Sub freegprd()
    Set gPrd = Nothing

    Set ProductObserver.CurrentProduct = gPrd ' ����Զ������¼�
    MsgBox "����մ�������Ʒ"
    Call clearall
End Sub



