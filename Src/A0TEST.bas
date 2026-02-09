Attribute VB_Name = "A0TEST"
Sub tet()

Dim pub1 As Object
Dim oPrd As Object
Set oPrd = CATIA.ActiveDocument.Product.ReferenceProduct
Set pubs = oPrd.Publications
On Error Resume Next
pubs.Remove ("Density")
Set pub1 = pubs.Add("Density")
CATIA.ActiveDocument.part.Update
Set pub1 = pubs.item("Density")

 Dim oRef As Object: Set oRef = refPrd.CreateReferenceFromName(oPrd.ReferenceProduct.partName & "\" & "Parameters\Partinfo\Density")
        pubs.SetDirect pubName, oRef

On Error GoTo 0
End Sub
