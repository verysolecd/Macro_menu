Sub CATMain()
    If CATIA.Windows.Count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
    
  If Not CanExecute("PartDocument") Then Exit Sub

    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory


'======= 要求选择surface
    Dim imsg, filter(0)
    imsg = "请选择孔所在的面或联合面"
    filter(0) = "Face" '
    Dim oFace
    Set oFace = KCL.SelectElement(imsg, filter).Value
    Set targethb = oPart.HybridBodies.Add()
    targethb.Name = "extracted points"
    If Not oFace Is Nothing Then
    Set oref = oprt.CreateReferenceFromObject(oFace)
    Set Extract = HSF.AddNewExtract(ref)
    end if




Dim hybridShapeSurfaceExplicit1 As HybridShapeSurfaceExplicit
Set hybridShapeSurfaceExplicit1 = parameters1.item("Surface.34")

Dim oref As Reference
Set oref = oprt.CreateReferenceFromObject(hybridShapeSurfaceExplicit1)

Dim oBdry As HybridShapeBoundary
Set oBdry = HSF.AddNewBoundaryOfSurface(oref)

Dim HBS As HybridBodies
Set HBS = oprt.HybridBodies

Set oHB = HBS.item("Geometrical Set.808")

oHB.AppendHybridShape oBdry

oprt.Update

'
'Dim oWindow As SpecsAndGeomWindow
'Set oWindow = CATIA.ActiveWindow
'
'Dim oViewer As Viewer3D
'Set oViewer = oWindow.ActiveViewer
'
'Dim viewpoint3D1 As Viewpoint3D
'Set viewpoint3D1 = oViewer.Viewpoint3D

Dim parameters2 As Parameters
Set parameters2 = oprt.Parameters

Dim iBdry As HybridShapeCircleExplicit
Set iBdry = parameters2.item("Circle.49")

Dim oref As Reference
Set oref = oprt.CreateReferenceFromObject(iBdry)

Dim mycenter As HybridShapePointCenter
Set mycenter = HSF.AddNewPointCenter(oref)

oHB.AppendHybridShape mycenter

oprt.Update

Set opt = HSF.AddNewPointCoord(0#, 0#, 0#)

oHB.AppendHybridShape opt

oprt.Update

End Sub

