Attribute VB_Name = "OTH_unfoldme"
'Attribute VB_Name = "OTH_unfoldme"
'{GP:6}
'{Ep:unfold_children}
'{Caption:展开子图形}
'{ControlTipText:遍历几何图形集的图形并展开}
'{BackColor:16744703}


Private Const mdlname As String = "OTH_unfoldme"
Sub unfold_children()
If Not CanExecute("PartDocument") Then Exit Sub
Dim osel: Set osel = CATIA.ActiveDocument.Selection
osel.Clear
Set oPrt = CATIA.ActiveDocument.part
Set HSF = oPrt.HybridShapeFactory: Set HBS = oPrt.HybridBodies
Dim uFold As HybridShapeUnfold: Set uFold = HSF.AddNewUnfold()
Dim imsg, filter(0)
imsg = "请先选择body，再选择平面"
filter(0) = "HybridBody"
Set itm = KCL.SelectItem(imsg, filter)
If Not itm Is Nothing Then
    Set oHb = itm
    Set oshapes = oHb.HybridShapes
Else
    Exit Sub
End If
    filter(0) = "Plane"
    Set itm = KCL.SelectItem(imsg, filter)
If Not itm Is Nothing Then
    Set oPlane = itm
    Set refplane = oPrt.CreateReferenceFromObject(oPlane)
Else
    Exit Sub
End If
oPrt.Update
Dim targetshape, ref
For i = 1 To oshapes.count
    Set targetshape = oshapes.item(i)
    oPrt.Update

FT = HSF.GetGeometricalFeatureType(targetshape)
If FT <> 5 Then

    oPrt.Update
Else
    Set ref = oPrt.CreateReferenceFromObject(targetshape)
    uFold.SurfaceToUnfold = ref
    Set dir1 = HSF.AddNewDirectionByCoord(1#, 0#, 0#)
    Set dir2 = HSF.AddNewDirectionByCoord(0#, 1#, 0#)
    Set dir3 = HSF.AddNewDirectionByCoord(0#, 0#, 1#)
    Dim extm As HybridShapeExtremum
    Set extm = HSF.AddNewExtremum(ref, dir1, 1)
    extm.Direction2 = dir2
    extm.ExtremumType2 = 1
    extm.Direction3 = dir3
    extm.ExtremumType3 = 1
    Set reforg = oPrt.CreateReferenceFromObject(extm)
    uFold.OriginToUnfold = reforg
    Set refDir = oPrt.CreateReferenceFromObject(dir1)
    uFold.DirectionToUnfold = refDir
    uFold.TargetPlane = refplane
    uFold.SurfaceType = 0 '0
    uFold.TargetOrientationMode = 0
    uFold.EdgeToTearPositioningOrientation = 0
    uFold.Name = "unfold_" & targetshape.Name
    oPrt.Update
    osel.Clear
    osel.Add uFold
    osel.Copy
    osel.Clear
    oPrt.Update
    Set targetHB = HBS.Add()
    targetHB.Name = "unfold result" & i
    oPrt.Update
    osel.Add targetHB
    osel.Paste
    osel.Clear
    oPrt.Update
    oPrt.InWorkObject = targetHB
End If
Next
oPrt.Update

End Sub

