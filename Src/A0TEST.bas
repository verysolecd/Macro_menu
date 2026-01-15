Attribute VB_Name = "A0TEST"


Sub CATMain()

Dim productDocument1 As Document
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

Dim groups1 As AnyObject
Set groups1 = product1.GetTechnologicalObject("Groups")

Dim group1 As Group
Set group1 = groups1.Add()

group1.AddExplicit product1

Dim silhouettes1 As AnyObject
Set silhouettes1 = product1.GetTechnologicalObject("Silhouettes")

Dim arrayOfVariantOfDouble1(17)
arrayOfVariantOfDouble1(0) = 1#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = -1#
arrayOfVariantOfDouble1(4) = 0#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 1#
arrayOfVariantOfDouble1(8) = 0#
arrayOfVariantOfDouble1(9) = 0#
arrayOfVariantOfDouble1(10) = -1#
arrayOfVariantOfDouble1(11) = 0#
arrayOfVariantOfDouble1(12) = 0#
arrayOfVariantOfDouble1(13) = 0#
arrayOfVariantOfDouble1(14) = 1#
arrayOfVariantOfDouble1(15) = 0#
arrayOfVariantOfDouble1(16) = 0#
arrayOfVariantOfDouble1(17) = -1#
Dim document1 As Document
Set document1 = silhouettes1.ComputeASilhouette(group1, arrayOfVariantOfDouble1, 20#, 0#)

Dim optimizerWorkBench1 As Workbench
Set optimizerWorkBench1 = productDocument1.GetWorkbench("OptimizerWorkBench")

groups1.Remove group1

document1.Activate

document1.SaveAs "./Product1_SILHOUETTE.cgr"

productDocument1.Activate

productDocument1.Activate

End Sub

