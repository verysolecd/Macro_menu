Attribute VB_Name = "移除旧的"
Sub CATMain()
Set oprd = CATIA.ActiveDocument.Product
rm oprd
End Sub
Sub rm(oprd)
    On Error Resume Next
     Set refPrd = oprd.ReferenceProduct
     Set oprt = refPrd.Parent.Part
    Set colls = refPrd.Publications
    colls.Remove ("Location")
    colls.Remove ("iMass")
    colls.Remove ("iDensity")
    colls.Remove ("iThickness")
    colls.Remove ("iMaterial")
     Set colls = refPrd.Parent.Part.Parameters.RootParameterSet.ParameterSets
        Set cm = colls.GetItem("cm")
        Set osel = CATIA.ActiveDocument.Selection
        osel.Clear
        osel.Add cm
        osel.Delete
		
     Set colls = refPrd.Parent.Part.relations
     colls.Remove ("CalM")
     colls.Remove ("CMAS")
     colls.Remove ("CTK")
	 
     Set colls = refPrd.UserRefProperties
     colls.Remove ("iMass")
     colls.Remove ("iMaterial")
     colls.Remove ("iThickness")
    If oprd.Products.Count > 0 Then
        For i = 1 To oprd.Products.Count
          rm (oprd.Products.item(i))
        Next
    End If
On Error GoTo 0
End Sub