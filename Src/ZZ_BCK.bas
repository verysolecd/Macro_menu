Attribute VB_Name = "ZZ_BCK"
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Const mdlname As String = "A00rmusrPrp"
Sub remove_usrP()
Set oprd = CATIA.ActiveDocument.Product
rm oprd
End Sub
Sub rm(oprd)
    On Error Resume Next
     Set refPrd = oprd.ReferenceProduct
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
        osel.Clear: osel.Add cm: osel.Delete
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


''==±éÀúµÝ¹é=============================
Sub recurAyo(ayo)
    Dim colls: Set itm = ayo.Products
    For Each itm In colls
        Call recurFunc(itm)
    Next

    If ayo.Products.count > 0 Then
            For Each ctm In ayo.Products
                Call recurAyo(ctm)
             Next
    End If
End Sub

