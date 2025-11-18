Attribute VB_Name = "RW_3initme"
'Attribute VB_Name = "initme"
'{GP:1}
'{Ep:initme}
'{Caption:初始化模板}
'{ControlTipText:将选择的产品和子产品文档按模板格式化}
'{BackColor:1229803}

Sub initme()

 If Not KCL.CanExecute("ProductDocument,PartDocument") Then Exit Sub
 If pdm Is Nothing Then Set pdm = New class_PDM
 Set allPN = KCL.InitDic(vbTextCompare): allPN.RemoveAll  'allPn 是全局变量，不需要传递
  Dim iprd
 If KCL.checkDocType("PartDocument") Then
 Set iprd = CATIA.ActiveDocument.Product
 
 Call pdm.initPrd(rootprd)
 Exit Sub
 End If

 Set iprd = pdm.defgprd()
 If Not iprd Is Nothing Then
     On Error Resume Next
            Call recurInitPrd(iprd)
            allPN.RemoveAll
       If Error.Number = 0 Then
                MsgBox "零件模板已经应用"
          Else
                MsgBox "至少一个参数创建错误，请检查lisence"
                 End If
      On Error GoTo 0
   Else
    MsgBox "没有要初始化的产品或零件，将退出"
 End If
 
End Sub
Sub recurInitPrd(oprd)
        If allPN.Exists(oprd.PartNumber) = False Then
            allPN(oprd.PartNumber) = 1
            Call pdm.initPrd(oprd)
        End If
    If oprd.Products.count > 0 Then
            For Each Product In oprd.Products
                Call recurInitPrd(Product)
             Next
    End If
End Sub


