Attribute VB_Name = "MDL_addPlnSKT"
'{GP:4}
'{Ep:addPlnSKT}
'{Caption:创建面&草图}
'{ControlTipText:创建面&草图}
'{BackColor: }
Private m_Doc         As Document       ' 当前激活文档
Private m_workPrtDoc   As PartDocument   ' 当前激活的零件文档
Private m_prt         As part           ' 当前激活的Part对象
Private m_sel         As Selection      ' 选择集对象
Private Const mdlname As String = "MDL_addPlnSKT"
Sub addPlnSKT()
    If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
    Dim ox As Double, oy As Double, oz As Double
    ox = 10
    oy = 20
    oz = 30
    ' ==============================================
    Set oprt = m_prt
    Set HSF = oprt.HybridShapeFactory
    Set hb = oprt.InWorkObject
    Set hb = hb.HybridBodies.Add
    hb.Name = "Skts"
     Set skts = hb.HybridSketches
    Set ptcoord = HSF.AddNewPointCoord(ox, oy, oz)
    Set refPoint = oprt.CreateReferenceFromObject(ptcoord)
    hb.AppendHybridShape ptcoord
    Dim pln(2)
    Set pln(0) = HSF.AddNewPlaneEquation(1#, 0#, 0#, 0)  ' YZ平面
    Set pln(1) = HSF.AddNewPlaneEquation(0#, 1#, 0#, 0)  ' XZ平面
    Set pln(2) = HSF.AddNewPlaneEquation(0#, 0#, 1#, 0)  ' XY平面
    Dim skt(2)
    Dim plnm
    plnm = Array("pln_x", "pln_y", "pln_z")
    Dim axidata(2)
        axidata(0) = Array(ox, oy, oz, 0, 1, 0, 0, 0, 1)
        axidata(1) = Array(ox, oy, oz, 1, 0, 0, 0, 0, 1)
        axidata(2) = Array(ox, oy, oz, 1, 0, 0, 0, 1, 0)
    For i = 0 To 2
        pln(i).SetReferencePoint refPoint
        pln(i).Name = plnm(i)
        hb.AppendHybridShape pln(i)
    Next
    oprt.Update
    For i = 0 To 2
        Set skt(i) = skts.Add(pln(i))
        skt(i).SetAbsoluteAxisData axidata(i)
        skt(i).Name = "Skt_" & plnm(i)
    Next i
    oprt.Update
    oprt.InWorkObject = hb
    oprt.Update
End Sub
