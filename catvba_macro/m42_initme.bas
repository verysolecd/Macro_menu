Attribute VB_Name = "m42_initme"
'Attribute VB_Name = "initme"
'{GP:4}
'{Ep:initme}
'{Caption:初始化模板}
'{ControlTipText:将选择的产品和子产品文档按模板格式化}
'{BackColor:16744703}

Sub initme()

     If pdm Is Nothing Then
          Set pdm = New class_PDM
     End If
     
     Set allPN = KCL.InitDic(vbTextCompare)
     allPN.RemoveAll
     
    Dim oPrd
'    If Not gPrd Is Nothing Then
'        Set oPrd = gPrd
        Set oPrd = CATIA.ActiveDocument.Product
'
'    End If
    
     Call ini_oPrd(oPrd)
     
     allPN.RemoveAll
     MsgBox "零件模板已经应用"
'
'    Else
'        MsgBox "请先选择要初始化的产品"
'    End If
    
End Sub

Sub ini_oPrd(oPrd)

        If allPN.Exists(oPrd.PartNumber) = False Then
            allPN(oPrd.PartNumber) = 1
            Call pdm.initPrd(oPrd)
        End If
            For Each Product In oPrd.Products
                Call ini_oPrd(Product)
          Next
End Sub

