Attribute VB_Name = "MDL_addgeotree"
'{GP:4}
'{Ep:newgeo_Tree}
'{Caption:创建几何集}
'{ControlTipText:创建基于模板的几何树}
'{BackColor: }


Private m_Doc         As Document       ' 当前激活文档
Private m_workPrtDoc   As PartDocument   ' 当前激活的零件文档
Private m_prt         As part           ' 当前激活的Part对象
Private m_sel         As Selection      ' 选择集对象
Private Const mdlname As String = "MDL_addgeotree"
Sub newgeo_Tree()
    If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
    Set colls = m_prt.HybridBodies
    On Error Resume Next
    Set og = colls.item("Geo_sheet")
    On Error GoTo 0
Set og = colls.Add(): og.Name = "GEO_sheet"
crSkt og
Set colls = og.HybridBodies
arr = Array("01_Profile", "02_Ribs", "03_Assy", "04_trim", "05_Pierce", "06_final part")
For i = 0 To UBound(arr)
    Set og = colls.Add()
    og.Name = arr(i)
Next
    
Set og = colls.item(arr(3))
Set subcolls = og.HybridBodies
For i = 1 To 3
    Set og = subcolls.Add(): og.Name = "TR_0" & i
Next
Set og = colls.item(arr(4))
Set subcolls = og.HybridBodies
For i = 1 To 3
    Set og = subcolls.Add(): og.Name = "PI_0" & i
Next

End Sub



Sub crSkt(og)
m_prt.InWorkObject = og
Set HSF = m_prt.HybridShapeFactory
Set oPoint = HSF.AddNewPointCoord(0#, 0#, 0#)
og.AppendHybridShape oPoint
m_prt.InWorkObject = oPoint
m_prt.Update
Set oPln = HSF.AddNewPlaneEquation(0#, 0#, 1#, 20#)
Set pref = oPoint
Set oRef = m_prt.CreateReferenceFromObject(pref)
oPln.SetReferencePoint oPoint  'oref
og.AppendHybridShape oPln
m_prt.InWorkObject = oPln
m_prt.Update
Set skts = og.HybridSketches
Set oSkt = og.HybridSketches.Add(oPln)
m_prt.InWorkObject = oSkt
Set factory2D1 = oSkt.OpenEdition()
Set geometricElements1 = oSkt.GeometricElements
Set axis2D1 = geometricElements1.item("AbsoluteAxis")
Set line2D1 = axis2D1.GetItem("HDirection")
line2D1.ReportName = 1
Set line2D2 = axis2D1.GetItem("VDirection")
line2D2.ReportName = 2
Set circle2D1 = factory2D1.CreateClosedCircle(0#, 0#, 10#)
Set point2D1 = axis2D1.GetItem("Origin")
circle2D1.CenterPoint = point2D1
circle2D1.ReportName = 3
oSkt.CloseEdition
m_prt.InWorkObject = og
m_prt.Update
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
