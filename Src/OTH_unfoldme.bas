Attribute VB_Name = "OTH_unfoldme"
'Attribute VB_Name = "OTH_unfoldme"
'{GP:6}
'{Ep:CATMain}
'{Caption:展开子图形}
'{ControlTipText:遍历几何图形集的图形并展开}
'{BackColor:16744703}


Sub CATMain()
If Not CanExecute("PartDocument") Then Exit Sub
Dim oSel: Set oSel = CATIA.ActiveDocument.Selection
oSel.Clear
Set oprt = CATIA.ActiveDocument.part
Set HSF = oprt.HybridShapeFactory: Set HBS = oprt.HybridBodies
Dim uFold As HybridShapeUnfold: Set uFold = HSF.AddNewUnfold()
Dim imsg, filter(0)
imsg = "请先选择body，再选择平面"
filter(0) = "HybridBody"
Set itm = KCL.SelectItem(imsg, filter)
If Not itm Is Nothing Then
    Set oHB = itm
    Set oshapes = oHB.HybridShapes
Else
    Exit Sub
End If
    filter(0) = "Plane"
    Set itm = KCL.SelectItem(imsg, filter)
If Not itm Is Nothing Then
    Set oPlane = itm
    Set refplane = oprt.CreateReferenceFromObject(oPlane)
Else
    Exit Sub
End If
oprt.Update
Dim targetshape, ref
For i = 1 To oshapes.count
    Set targetshape = oshapes.item(i)
    oprt.Update

FT = HSF.GetGeometricalFeatureType(targetshape)
If FT <> 5 Then

    oprt.Update
Else
    Set ref = oprt.CreateReferenceFromObject(targetshape)
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
    Set reforg = oprt.CreateReferenceFromObject(extm)
    uFold.OriginToUnfold = reforg
    Set refDir = oprt.CreateReferenceFromObject(dir1)
    uFold.DirectionToUnfold = refDir
    uFold.TargetPlane = refplane
    uFold.SurfaceType = 0 '0
    uFold.TargetOrientationMode = 0
    uFold.EdgeToTearPositioningOrientation = 0
    uFold.Name = "unfold_" & targetshape.Name
    oprt.Update
    oSel.Clear
    oSel.Add uFold
    oSel.Copy
    oSel.Clear
    oprt.Update
    Set targetHB = HBS.Add()
    targetHB.Name = "unfold result" & i
    oprt.Update
    oSel.Add targetHB
    oSel.Paste
    oSel.Clear
    oprt.Update
    oprt.InWorkObject = targetHB
End If
Next
oprt.Update

End Sub

