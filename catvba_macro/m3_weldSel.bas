Attribute VB_Name = "m3_weldSel"
'Attribute VB_Name = "weldSel"
'{GP:3}
'{Ep:CATMain}
'{Caption:��Ʒ����}
'{ControlTipText:ѡ�����ӵĲ�Ʒ}
'{BackColor:16744703}

Sub CATMain()
Set Doc = CATIA.ActiveDocument
Set rootPrd = Doc.Product
Set sPrd = rootPrd.Products
Set iPrd = sPrd.Item("�㺸��Ϣ")
Set oSel = Doc.Selection
Dim oPn
Dim iType(0)
oSel.Clear
iType(0) = "Product"
status = oSel.SelectElement3(iType, "ѡ�����Ӳ�Ʒ", True, 2, False)
If status = "Normal" And oSel.Count2 <= 3 Then
oName = ""
For i = 1 To oSel.Count
     oPn = oPn & "_" & oSel.Item(i).LeafProduct.PartNumber
Next
 iPn = "SotWeld_" & oPn
     MsgBox iPn
End If
Set oPrd = iPrd.Products.AddNewComponent("Part", "")
oPrd.PartNumber = iPn
oSel.Clear
End Sub
 

