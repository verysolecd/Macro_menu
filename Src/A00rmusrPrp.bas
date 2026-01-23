Attribute VB_Name = "A00rmusrPrp"
Private Const mdlname As String = "A00rmusrPrp"
Sub remove_usrP()
Set oPrd = CATIA.ActiveDocument.Product
rm oPrd
End Sub
Sub rm(oPrd)
    On Error Resume Next
     Set refPrd = oPrd.ReferenceProduct
     Set oPrt = refPrd.Parent.part
    Set colls = refPrd.Publications
    colls.Remove ("Location")
    colls.Remove ("iMass")
    colls.Remove ("iDensity")
    colls.Remove ("iThickness")
    colls.Remove ("iMaterial")
     Set colls = refPrd.Parent.part.Parameters.RootParameterSet.ParameterSets
        Set cm = colls.GetItem("cm")
        Set osel = CATIA.ActiveDocument.Selection
        osel.Clear
        osel.Add cm
        osel.Delete
                
     Set colls = refPrd.Parent.part.relations
     colls.Remove ("CalM")
     colls.Remove ("CMAS")
     colls.Remove ("CTK")
         
     Set colls = refPrd.UserRefProperties
     colls.Remove ("iMass")
     colls.Remove ("iMaterial")
     colls.Remove ("iThickness")
    If oPrd.Products.count > 0 Then
        For i = 1 To oPrd.Products.count
          rm (oPrd.Products.item(i))
        Next
    End If
On Error GoTo 0
End Sub
