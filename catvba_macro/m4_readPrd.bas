Attribute VB_Name = "m4_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:��ȡ��Ʒ����}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}


Sub readPrd()
    
    Dim xlm, pdm, ws
    Set xlm = New Class_XLM
    xlm.init
    Set pdm = New class_PDM
    pdm.init
    Set ws = xlm.ws
    
    Dim data
    data = pdm.infoPrd(pdm.rootPrd)
    xlm.inject_data 1, data, "rv"
        



Call selPrd

End Sub

Function selPrd()

MsgBox "ѡ��ɹ�"


End Function
