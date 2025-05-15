Attribute VB_Name = "m32_weldSel"
'Attribute VB_Name = "weldSel"
'{GP:3}
'{Ep:CATMain}
'{Caption:产品焊缝}
'{ControlTipText:选择被连接的产品}
'{BackColor:}

Sub CATMain()
Set Doc = CATIA.ActiveDocument
Set rootPrd = Doc.Product
Set sPrd = rootPrd.Products
Set iPrd = sPrd.item("点焊信息")
Set osel = Doc.Selection
Dim oPn
Dim iType(0)
osel.Clear
iType(0) = "Product"
status = osel.SelectElement3(iType, "选择被连接产品", True, 2, False)
If status = "Normal" And osel.Count2 <= 3 Then
oName = ""
For i = 1 To osel.Count
     oPn = oPn & "_" & osel.item(i).LeafProduct.PartNumber
Next
 iPn = "SotWeld_" & oPn
     MsgBox iPn
End If
Set oprd = iPrd.Products.AddNewComponent("Part", "")
oprd.PartNumber = iPn
osel.Clear
End Sub
 

