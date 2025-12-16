Attribute VB_Name = "Module3"
Sub CATMain()

Dim rtDoc As ProductDocument
Set rtDoc = CATIA.ActiveDocument

Dim oWB As OptimizerWorkBench
Set oWB = rtDoc.GetWorkbench("OptimizerWorkBench")

Dim Cps As PartComps
Set Cps = oWB.PartComps

Dim prd1 As Product
Set prd1 = rtDoc.Product

Dim products1 As Products
Set products1 = prd1.Products

Dim prd2 As Product
Set prd2 = products1.item("产品3")

Dim product3 As Product
Set product3 = products1.item("键盘造车手出品1003.1.1")




Dim cp1 As PartComp
Set cp1 = Cps.Add(prd2, product3, 1#, 1#, 6)

Dim oDoc As Document
Set oDoc = CATIA.Documents.item("AddedMaterial.3dmap")

oDoc.Activate

oDoc.SaveAs "D:\tt\Desktop\temp\coding test\AddedMaterial.3dmap"

Dim Doc2 As Document
Set Doc2 = CATIA.Documents.item("RemovedMaterial.3dmap")

Doc2.Activate

Dim aryBSTR(0)
aryBSTR(0) = "D:\tt\Desktop\temp\coding test\AddedMaterial.3dmap"


Set products1Variant = products1
products1Variant.AddComponentsFromFiles aryBSTR, "3dmap"

Set prd1 = prd1.ReferenceProduct

Doc2.SaveAs "D:\tt\Desktop\temp\coding test\RemovedMaterial.3dmap"

Dim arrayOfVariantOfBSTR2(0)
arrayOfVariantOfBSTR2(0) = "D:\tt\Desktop\temp\coding test\RemovedMaterial.3dmap"
Set products1Variant = products1
products1Variant.AddComponentsFromFiles arrayOfVariantOfBSTR2, "3dmap"

Set prd1 = prd1.ReferenceProduct
Set rtDoc = CATIA.ActiveDocument

rtDoc.Save

End Sub


