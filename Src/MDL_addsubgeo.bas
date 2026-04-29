Attribute VB_Name = "MDL_addsubgeo"
'Attribute VB_Name = "m20_newgeotree"
'{GP:4}
'{Ep:addsubgeo}
'{Caption:子几何集}
'{ControlTipText:创建一个子几何集}
'{BackColor: }

Private Const mdlname As String = "MDL_addsubgeo"
Sub addsubgeo()
 If Not CanExecute("Productdocument,PartDocument") Then Exit Sub
    On Error Resume Next
        Dim oDoc: Set oDoc = CATIA.ActiveDocument
        Dim workPrtDoc: Set workPrtDoc = KCL.get_workPartDoc
        Dim oprt: Set oprt = Nothing: Set oprt = workPrtDoc.part
    Err.Clear
    On Error GoTo 0
    If IsNothing(oprt) Then: MsgBox "No activated Part": Exit Sub
    Set igeo = Nothing
    Set colls = oprt.HybridBodies
    itype = TypeName(oprt.InWorkObject)
    If LCase(itype) = LCase("hybridbody") Then
        Set igeo = oprt.InWorkObject
        Set colls = igeo.HybridBodies
    End If
    Set og = colls.Add(): og.Name = "FAXX"
    Set og = colls.Add(): og.Name = "FAXX"
    On Error Resume Next
    If Not igeo Is Nothing Then oprt.InWorkObject = og.HybridBodies.Parent
        Error.Clear
    On Error Resume Next
End Sub

