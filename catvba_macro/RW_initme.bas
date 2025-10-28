Attribute VB_Name = "RW_initme"
'Attribute VB_Name = "initme"
'{GP:1}
'{Ep:initme}
'{Caption:初始化模板}
'{ControlTipText:将选择的产品和子产品文档按模板格式化}
'{BackColor: }

Sub initme()



 
 If KCL.CanExecute("ProductDocument") Then
     If pdm Is Nothing Then
       Set pdm = New class_PDM
     End If
     Set allPN = KCL.InitDic(vbTextCompare)
     allPN.RemoveAll
    Dim iprd: Set iprd = pdm.defgprd()
    If Not iprd Is Nothing Then
        On Error Resume Next
           Call ini_oPrd(iprd)
             allPN.RemoveAll
            
                 If Error.Number = 0 Then
                        MsgBox "零件模板已经应用"
                     Else
                     MsgBox "至少一个参数创建错误，请检查lisence"
                 End If
        On Error GoTo 0
    End If
    
    
Else
 On Error Resume Next
    If KCL.CanExecute("PartDocument") Then
      Set oprd = CATIA.ActiveDocument.Product
      If pdm Is Nothing Then
                Set pdm = New class_PDM
           End If
      Call pdm.initPrd(oprd)
    If Error.Number = 0 Then
            MsgBox "零件模板已经应用"
            Else
            MsgBox "至少一个参数创建错误，请检查lisence"
          End If
      On Error GoTo 0
End If
End If

End Sub
Sub ini_oPrd(oprd)
        If allPN.Exists(oprd.PartNumber) = False Then
            allPN(oprd.PartNumber) = 1
            Call pdm.initPrd(oprd)
        End If
            For Each Product In oprd.Products
                Call ini_oPrd(Product)
        Next
End Sub


