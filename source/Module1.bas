Attribute VB_Name = "Module1"


''=========================
'' Difference percentages :
''=========================
''  added : Added Material /Material in Version1
''  removed : Removed Material /Material in Version1



Sub CATMain()
 Dim added As Double
Dim removed As Double
MinDiff = 0.3

Dim Docs As Documents
Set Docs = CATIA.Documents

''======================
'' New Product Creation

Dim oDoc As Document
Set oDoc = CATIA.ActiveDocument  'Docs.Add("Product")

Dim oPrd As Product
Set oPrd = oDoc.Product


Set ss = oPrd.Products
 
 
Dim colls As Products
Set colls = ss
''======================


''======================================
'' Names of the two products to compare
Dim ary(0)

ary(0) = "D:\tt\Desktop\temp\coding test\Part1.CATPart"
'ary(1) = "D:\tt\Desktop\temp\coding test\Part3.CATPart"

''======================================

''======================
'' Insertion of products

colls.AddComponentsFromFiles ary, "All"
''======================

Dim optimizerWorkBench1 As Workbench
Set optimizerWorkBench1 = oDoc.GetWorkbench("OptimizerWorkBench")

''======================================
'' Products to compare
''=======================================

Dim product2 As Product
Set product2 = colls.item(1)

Dim product3 As Product
Set product3 = colls.item(2)

''=====================================
'' Comparison
''=====================================
Dim partComps1 As PartComps
Set partComps1 = optimizerWorkBench1.PartComps
Dim partComp1 As PartComp
''Set partComp1 = partComps1.GeometricComparison(product2, product3, 2.000000, 2.000000, 2, added, removed)

''======================================
'' Start Comparison
'' Parameters :
''      product2 : first product to compare (Old Version)
''      product3 : second product to compare (New Version)
''      2.000000 : computation accuracy (mm)
''      2.000000 : display accuracy (mm)
''      2 : computation type : 0=Added, 1=Removed, 2=Added+Removed
''========================================
Set partComp1 = partComps1.Add(product2, product3, 2#, 2#, 2)

''=====================================
'' Read computation results
''=====================================

'' Retrieve the percent of added material (value is between 0.0 and 1.0)
Dim PercentAdded As Double
PercentAdded = partComps1.AddedMaterialPercentage

'' Retrieve the percent of removed material (value is between 0.0 and 1.0)
Dim PercentRemoved As Double
PercentRemoved = partComps1.RemovedMaterialPercentage

'' Retrieve the volume of added material (mm3)
Dim VolumeAdded As Double
VolumeAdded = partComps1.AddedMaterialVolume

'' Retrieve the volume of removed material (mm3)
Dim VolumeRemoved As Double
VolumeRemoved = partComps1.RemovedMaterialVolume


''====================================
'' Typical comparison result management
''====================================
If PercentAdded > MinDiff Then
    MsgBox "Difference detected : Added =  " & CStr(PercentAdded) & " , Removed = " & CStr(PercentRemoved) & " VolumeAdded = " & CStr(VolumeAdded) & " VolumeRemoved = " & CStr(VolumeRemoved)

    ''=======================================
    '' Save of added and removed Material
    ''=======================================
    
    Dim document1 As Document
    Set document1 = Docs.item("AddedMaterial.3dmap")
    document1.Activate
    'document1.SaveAs "E:\users\sbc\DemoSMT\Comparison\AddedMaterial.3dmap"
    document1.SaveAs "D:\tt\Desktop\temp\coding test\AddedMaterial.3dmap"

    Dim document2 As Document
    Set document2 = Docs.item("RemovedMaterial.3dmap")
    document2.Activate
   ' document2.SaveAs "E:\users\sbc\DemoSMT\Comparison\RemovedMaterial.3dmap"
    document2.SaveAs "D:\tt\Desktop\temp\coding test\RemovedMaterial.3dmap"
    
    
    document2.Close
    document1.Close
    
    '' =======================================================
    '' Import AddedMaterial Only
    '' =======================================================
    Dim var11(0)
    'var11(0) = "E:\users\sbc\DemoSMT\Comparison\AddedMaterial.3dmap"
    
    
    var11(0) = "D:\tt\Desktop\temp\coding test\AddedMaterial.3dxml"
   colls.AddComponentsFromFiles var11, "*"

    '' =======================================================
    '' Definition du view point
    '' =======================================================

    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutSightDirection Array(1#, 1, 0)
    CATIA.ActiveWindow.ActiveViewer.Viewpoint3D.PutUpDirection Array(0, 0, 1)

    CATIA.ActiveWindow.ActiveViewer.Reframe

  
    'CATIA.ActiveWindow.ActiveViewer.CaptureToFile catCaptureFormatJPEG, "E:\users\sbc\DemoSMT\Comparison\MyImage.jpg"

Else
    MsgBox "No difference detected between products"
End If

oDoc.Activate

End Sub

