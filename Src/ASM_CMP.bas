Attribute VB_Name = "ASM_CMP"
'Attribute VB_Name = "ASM_CMP"
'{GP:3}
'{Ep:myCMP}
'{Caption:数据对比}
'{ControlTipText:对比旧、新数据}
'{BackColor:}

Sub myCMP()
 If Not CanExecute("ProductDocument") Then Exit Sub
Dim rootprd As Product
Dim colls As Products
Dim oWB As OptimizerWorkBench
Dim oComps As PartComps
Dim docs As Documents
    Dim Mt(1), pn2, opath(2), filePath(1), mapName(1)
Dim rtDoc As ProductDocument
Set rtDoc = CATIA.ActiveDocument
Set docs = CATIA.Documents
Set oWB = rtDoc.GetWorkbench("OptimizerWorkBench")
Set oComps = oWB.PartComps
Set rootprd = rtDoc.Product
Set colls = rootprd.Products
 Dim imsg, filter(0)
    imsg = "请依次选择旧版本、新版本零件"
    filter(0) = "Product"
    Set prd1 = KCL.SelectItem(imsg, filter)
    If prd1 Is Nothing Then Exit Sub
    imsg = "请选择新版本零件"
    Set prd2 = KCL.SelectItem(imsg, filter)
      If prd2 Is Nothing Then Exit Sub
    If Not IsNothing(prd1) And Not IsNothing(prd1) Then
            Dim CMPR: Set CMPR = oComps.Add(prd1, prd2, 1#, 1#, 2)
                pn2 = KCL.rmchn(prd2.PartNumber)
                opath(0) = prd2.ReferenceProduct.Parent.path
                opath(2) = "3dmap"
                   Mt(0) = "AddedMaterial"
                    Mt(1) = "RemovedMaterial"
            For i = 0 To 1
             opath(1) = Mt(i)
             filePath(i) = JoinPathName(opath())
             mapName(i) = Mt(i) & ".3dmap"
             KCL.DeleteMe (filePath(i))
            Next
            For i = 0 To 1
                        Set oDoc = docs.item(mapName(i)): oDoc.Activate
                        oDoc.SaveAs filePath(i)
                        oDoc.Close
            Next
                On Error GoTo 0
                    Set Prdvariant = colls
                    Prdvariant.AddComponentsFromFiles filePath(), "*"
                    On Error GoTo 0

End If

End Sub
