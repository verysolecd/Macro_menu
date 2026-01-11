Attribute VB_Name = "MDL_newgeotree"
'Attribute VB_Name = "m20_newgeotree"
'{GP:4}
'{Ep:newgeo}
'{Caption:创建几何集}
'{ControlTipText:创建基于模板的几何树}
'{BackColor: }

Private oprt

Sub newgeo()
 If Not CanExecute("PartDocument") Then Exit Sub
 
    Set oDoc = CATIA.ActiveDocument.Product
    Set oprt = oDoc.ReferenceProduct.Parent.part
    Set colls = oprt.HybridBodies
    On Error Resume Next
    Set og = colls.item("Geo_sheet")
    On Error GoTo 0

Set og = colls.Add()
crSkt og
og.Name = "GEO_sheet"
Set colls = og.HybridBodies
arr = Array("01_Profile", "02_Ribs", "03_Assy", "04_trim", "05_Pierce", "06_final part")
For i = 0 To UBound(arr)
    Set og = colls.Add()
    og.Name = arr(i)
    Next
End Sub
Sub crSkt(og)
oprt.InWorkObject = og
Set HSF = oprt.HybridShapeFactory
Set oPoint = HSF.AddNewPointCoord(0#, 0#, 0#)
og.AppendHybridShape oPoint
oprt.InWorkObject = oPoint
oprt.Update
Set oPln = HSF.AddNewPlaneEquation(0#, 0#, 1#, 20#)
Set pref = oPoint
Set oref = oprt.CreateReferenceFromObject(pref)
oPln.SetReferencePoint oPoint  'oref
og.AppendHybridShape oPln
oprt.InWorkObject = oPln
oprt.Update
Set skts = og.HybridSketches
Set oSkt = og.HybridSketches.Add(oPln)
oprt.InWorkObject = oSkt
Set factory2D1 = oSkt.OpenEdition()
Set geometricElements1 = oSkt.GeometricElements
Set axis2D1 = geometricElements1.item("AbsoluteAxis")
Set line2D1 = axis2D1.getItem("HDirection")
line2D1.ReportName = 1
Set line2D2 = axis2D1.getItem("VDirection")
line2D2.ReportName = 2
Set circle2D1 = factory2D1.CreateClosedCircle(0#, 0#, 10#)
Set point2D1 = axis2D1.getItem("Origin")
circle2D1.CenterPoint = point2D1
circle2D1.ReportName = 3
oSkt.CloseEdition
oprt.InWorkObject = og
oprt.Update
''the first 3 being the coordinates of the axis origin,
'Dim arr(0 To 8)
'arr(0) = 0
'arr(1) = 0#
'arr(2) = 0#
'the next 3 being those of the horizontal axis,
'arr(3) = 1#
'arr(4) = 0#
'arr(5) = 0#
'
''and the last 3 those of the vertical axis of the absolute axis.
'arr(6) = 0#
'arr(7) = 1#
'arr(8) = 0#
'oSkt.SetAbsoluteAxisData (arr)
End Sub
