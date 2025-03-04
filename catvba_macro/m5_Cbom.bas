Attribute VB_Name = "m5_Cbom"
'Attribute VB_Name = "m5_Cbom"
'{GP:5}
'{Ep:recurPrd}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}


Sub CATMain()

    Dim xlm, pdm
    xlm = New Class_XLM: xlm.init
    pdm = New class_PDM: pdm.init
    ws = xlm.ws

    Dim oPrd
    Set oPrd = pdm.rootPrd
    
    Dim oRowNb: oRowNb = 2

    Call recurPrd(oPrd, ws, oRowNb, 1)


End Sub

Sub recurPrd(oPrd, xlsht, oRowNb, LV)
    Dim xlm, pdm
    xlm = New Class_XLM: xlm.init
    pdm = New class_PDM: pdm.init
    
        Dim bdict, i
        xlm.inject_data Info(oPrd, LV), xlsht, oRowNb
        
        
        If oPrd.Products.Count > 0 Then ' I
            Set bdict = CreateObject("Scripting.Dictionary")
            For i = 1 To oPrd.Products.Count
                If Not bdict.Exists(oPrd.Products.Item(i).PartNumber) Then
                    bdict(oPrd.Products.Item(i).PartNumber) = 1
                    oRowNb = oRowNb + 1
                    recurPrd oPrd.Products.Item(i), xlsht, oRowNb, LV + 1
                End If
            Next
        End If
    End Sub


