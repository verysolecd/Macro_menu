Attribute VB_Name = "pubtest"
Sub CATMain()
    Set oDoc = CATIA.Activedocument
    Set rootPrd = oDoc.product
    
    Set oPrd = rootPrd
    
    
    Dim refPrd: Set refPrd = oPrd.ReferenceProduct
    Dim oPrt: Set oPrt = refPrd.Parent.Part
    Set Pubs = refPrd.Publications
    
Set Target = refPrd.UserRefProperties.item("Mass")

Set oRef = oPrd.CreateReferenceFromName(Target.Name)


Set oPub = publications1.Add("Mass")

Pubs.SetDirect "Mass", oRef
    
    End Sub

