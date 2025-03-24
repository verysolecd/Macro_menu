Attribute VB_Name = "m32_weldSel"
'Attribute VB_Name = "weldSel"
'{GP:3}
'{Ep:CATMain}
'{Caption:产品焊缝}
'{ControlTipText:选择被连接的产品}
'{BackColor:16744703}

Sub CATmain()


MsgBox "代码还没写"
'Set Doc = CATIA.Activedocument
'Set rootPrd = Doc.product
'Set sPrd = rootPrd.Products
'Set iPrd = sPrd.item("点焊信息")
'Set oSel = Doc.Selection
'Dim oPn
'Dim iType(0)
'oSel.Clear
'iType(0) = "Product"
'status = oSel.SelectElement3(iType, "选择被连接产品", True, 2, False)
'If status = "Normal" And oSel.Count2 <= 3 Then
'oName = ""
'For i = 1 To oSel.Count
'     oPn = oPn & "_" & oSel.item(i).LeafProduct.PartNumber
'Next
' iPn = "SotWeld_" & oPn
'     MsgBox iPn
'End If
'Set oPrd = iPrd.Products.AddNewComponent("Part", "")
'oPrd.PartNumber = iPn
'oSel.Clear
End Sub
 

