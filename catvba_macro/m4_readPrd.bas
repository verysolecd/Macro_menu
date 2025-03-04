Attribute VB_Name = "m4_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:读取产品属性}
'{ControlTipText:选择要被读取或修改的产品}
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

MsgBox "选择成功"


End Function
