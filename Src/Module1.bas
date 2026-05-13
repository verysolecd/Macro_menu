Attribute VB_Name = "Module1"
Sub CATMain()
Set oDoc = CATIA.ActiveDocument
Set oprt = oDoc.part
Set HSF = oprt.HybridShapeFactory
Set HBS = oprt.HybridBodies
Set hb = oprt.InWorkObject
Set ptcoord = HSF.AddNewPointCoord(0#, 0#, 0#)
Set ref = oprt.CreateReferenceFromObject(ptcoord)
hb.AppendHybridShape ptcoord
Dim pln(2)
Set pln(0) = HSF.AddNewPlaneEquation(1#, 0#, 0#, 0)
Set pln(1) = HSF.AddNewPlaneEquation(0#, 1#, 0#, 0)
Set pln(2) = HSF.AddNewPlaneEquation(0#, 0#, 1#, 0)
For i = LBound(pln) To UBound(pln)
    pln(i).SetReferencePoint ref
    hb.AppendHybridShape pln(i)
Next
oprt.Update
Set skts = hb.HybridSketches
Set newSketch = skts.Add(pln(1))
'newSketch.SetAbsoluteAxisData Array(10, 20, 0, 1, 0, 0, 0, 1, 0)
sktabs(0) = ptcoord.X
sktabs(1) = ptcoord.Y
sktabs(2) = ptcoord.Z

 Dim myAxisCoordinate(8)
 newSketch.GetAbsoluteAxisData myAxisCoordinate
  
oprt.InWorkObject = hb
oprt.Update
End Sub
