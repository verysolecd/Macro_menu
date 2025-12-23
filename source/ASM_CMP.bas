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
Dim Docs As Documents
    Dim Mt(1), pn2, opath(2), filepath(1), mapName(1)
Dim rtDoc As ProductDocument
Set rtDoc = CATIA.ActiveDocument
Set Docs = CATIA.Documents
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
            For I = 0 To 1
             opath(1) = Mt(I)
             filepath(I) = JoinPathName(opath())
             mapName(I) = Mt(I) & ".3dmap"
             KCL.DeleteMe (filepath(I))
            Next
            For I = 0 To 1
                        Set oDoc = Docs.item(mapName(I)): oDoc.Activate
                        oDoc.SaveAs filepath(I)
                        oDoc.Close
            Next
                On Error GoTo 0
                    Set Prdvariant = colls
                    Prdvariant.AddComponentsFromFiles filepath(), "*"
                    On Error GoTo 0

End If

End Sub
