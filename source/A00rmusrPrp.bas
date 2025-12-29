Attribute VB_Name = "A00rmusrPrp"
Sub CATMain()
Set oprd = CATIA.ActiveDocument.Product
rm oprd
End Sub
Sub rm(oprd)
    On Error Resume Next
     Set refPrd = oprd.ReferenceProduct
     Set oprt = refPrd.Parent.part
    Set colls = refPrd.Publications
    colls.Remove ("Location")
    colls.Remove ("iMass")
    colls.Remove ("iDensity")
    colls.Remove ("iThickness")
    colls.Remove ("iMaterial")
     Set colls = refPrd.Parent.part.Parameters.RootParameterSet.ParameterSets
        Set cm = colls.getItem("cm")
        Set oSel = CATIA.ActiveDocument.Selection
        oSel.Clear
        oSel.Add cm
        oSel.Delete
                
     Set colls = refPrd.Parent.part.relations
     colls.Remove ("CalM")
     colls.Remove ("CMAS")
     colls.Remove ("CTK")
         
     Set colls = refPrd.UserRefProperties
     colls.Remove ("iMass")
     colls.Remove ("iMaterial")
     colls.Remove ("iThickness")
    If oprd.Products.count > 0 Then
        For i = 1 To oprd.Products.count
          rm (oprd.Products.item(i))
        Next
    End If
On Error GoTo 0
End Sub
