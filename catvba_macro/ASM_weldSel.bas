Attribute VB_Name = "ASM_weldSel"
'Attribute VB_Name = "weldSel"
'{GP:3}
'{Ep:CATMain}
'{Caption:��Ʒ����}
'{ControlTipText:ѡ�����ӵĲ�Ʒ}
'{BackColor:}

Sub CATMain()
If Not KCL.CanExecute("ProductDocument") Then Exit Sub
MsgBox "��û����"
'
'Set Doc = CATIA.ActiveDocument
'Set rootPrd = Doc.Product
'Set sPrd = rootPrd.Products
'Set iprd = sPrd.item("�㺸��Ϣ")
'Set osel = Doc.Selection
'Dim oPn
'Dim iType(0)
'osel.Clear
'iType(0) = "Product"
'status = osel.SelectElement3(iType, "ѡ�����Ӳ�Ʒ", True, 2, False)
'If status = "Normal" And osel.Count2 <= 3 Then
'oName = ""
'For i = 1 To osel.Count
'     oPn = oPn & "_" & osel.item(i).LeafProduct.PartNumber
'Next
' iPn = "SotWeld_" & oPn
'     MsgBox iPn
'End If
'Set oprd = iprd.Products.AddNewComponent("Part", "")
'oprd.PartNumber = iPn
'osel.Clear
End Sub
 

