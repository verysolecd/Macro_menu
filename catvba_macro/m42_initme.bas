Attribute VB_Name = "m42_initme"
'Attribute VB_Name = "initme"
'{GP:4}
'{Ep:initme}
'{Caption:初始化模板}
'{ControlTipText:将选择的产品和子产品文档按模板格式化}
'{BackColor:16744703}

Sub initme()

Set pdm = New class_PDM
    Set allPN = KCL.InitDic(vbTextCompare)
    allPN.RemoveAll
    
            Dim oPrd
        
'    If Not gPrd Is Nothing Then

        Set oPrd = rootPrd
        If allPN.Exists(oPrd.PartNumber) = False Then
            allPN(oPrd.PartNumber) = 1
            Call pdm.initPrd(oPrd)
        End If
            For Each product In oPrd.Products
                Call pdm.initPrd(product)
        Next
            allPN.RemoveAll
'    Else
'            MsgBox "请先选择要初始化的产品"
'    End If
End Sub
