Attribute VB_Name = "m_weldSel"
'Attribute VB_Name = "weldSel"
'{GP:51}
'{Caption:产品焊缝}
'{ControlTipText:选择被连接的产品}
'{BackColor:16744703}

Sub CATMain()
Set Doc = CATIA.ActiveDocument
Set rootPrd = Doc.Product
Set sPrd = rootPrd.Products
Set iPrd = sPrd.Item("点焊信息")
Set oSel = Doc.Selection
Dim oPn
Dim iType(0)
oSel.Clear
iType(0) = "Product"
Status = oSel.SelectElement3(iType, "选择被连接产品", True, 2, False)
If Status = "Normal" And oSel.Count2 <= 3 Then
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
 

meDic.Exists(Itm.Name) Then
            Set TgtDoc = ProdsNameDic.Item(Itm.Name)
        Else
            Set TgtDoc = Init_Part(Prods, Itm.Name)
            ProdsNameDic.Add Itm.Name, TgtDoc
        End If
        Call Preparing_Copy(BaseSel, Itm)
        With BaseSel
            .Copy
            .Clear
        End With
        With TopSel
            .Clear
            .Add TgtDoc.Part
            .PasteSpecial PasteType
        End With
    Next
    BaseSel.Clear
    TopSel.Clear
    