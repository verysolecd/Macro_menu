Attribute VB_Name = "m43_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:读取属性}
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
    data = pdm.infoPrd(rootPrd)
    xlm.inject_data 1, data, "rv"

End Sub

