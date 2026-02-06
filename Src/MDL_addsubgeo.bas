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
    Set odoc = CATIA.ActiveDocument.Product
    Set oprt = KCL.get_inwork_part
    Set colls = oprt.HybridBodies
itype = TypeName(oprt.InWorkObject)
    If LCase(itype) = LCase("hybridbody") Then
        Set colls = oprt.InWorkObject.HybridBodies
    Else
        Exit Sub
    End If
    Set og = colls.Add()
    og.name = "FAXX"
End Sub

