








Sub CATMain() 
Set Doc = CATIA.ActiveDocument
Set rootPrd = Doc.Product
Set sPrd = rootPrd.Products
Set iPrd = sPrd.Item("点焊信息")
Set oSel = Doc.Selection
Dim iType(0)
oSel.Clear
iType(0) = "Product"
Status = oSel.SelectElement3(iType, "选择被连接产品", True,2,False)
If Status = "Normal" and oSel.Count2 <= 3 Then
oName=""
for i=1 to osel.count
	oPn=oPn&"_" & oSel.Item(i).LeafProduct.PartNumber	
next
 iPn="SotWeld_"&oPn
	MsgBox  iPn
End If
Set oPrd = iPrd.Products.AddNewComponent("Part", "")
oPrd.PartNumber=iPn
oSel.Clear
End Sub
 
