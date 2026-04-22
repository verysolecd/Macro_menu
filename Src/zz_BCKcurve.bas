Attribute VB_Name = "zz_BCKcurve"
'Private Const TYPE_CURVE As Long = 3
'Private Const TYPE_SWEEP As Long = 7
'
'Sub rmCrv()
'    If CATIA.Windows.count < 1 Then MsgBox "没有打开的窗口": Exit Sub
'    Dim oDoc:  Set oDoc = CATIA.ActiveDocument
'    Dim oprt:  Set oprt = KCL.get_workPartDoc.part
'    Dim osel:  Set osel = oDoc.Selection
'    Dim HSF:   Set HSF = oDoc.part.HybridShapeFactory
'    CATIA.RefreshDisplay = False
'    CATIA.HSOSynchronized = False
'    ' ══ Step 1: Search 获取所有 Sweep，提取引用曲线 ══
'    Dim refSet: Set refSet = KCL.InitDic
'    Dim sweeps: Set sweeps = KCL.Initlst
'    osel.Clear
'    osel.Search "CATGMOSearch.Surface,all" ' Sweep 是曲面的子类型,或许可以使用
'    'osel.Search "CATPrtSearch.HybridShapeSweep,all"
'    Dim i
'    For i = 1 To osel.count
'        Dim shp: Set shp = osel.item(i).Value
'        If HSF.GetGeometricalFeatureType(shp) = TYPE_SWEEP Then sweeps.Add shp
'    Next
'    ' 解析每个 Sweep 引用的引导线（Reference → Shape 仍需 Resolve）
'    Dim sw, crv
'    For Each sw In sweeps
'        Set crv = Resolve(osel, sw.FirstGuideCrv)
'        If Not crv Is Nothing Then refSet(KCL.GetInternalName(crv)) = 1
'    Next
'    ' ══ Step 2: Search 获取所有曲线 ══
'    osel.Clear
'    osel.Search "CATGMOSearch.Curve,all"
'    Dim curves: Set curves = KCL.Initlst
'    For i = 1 To osel.count
'        curves.Add osel.item(i).Value
'    Next
'    ' ══ Step 3: 批量选中未引用曲线 → 一次删除 ══
'    osel.Clear
'    Dim c
'    For Each c In curves
'        If Not refSet.Exists(KCL.GetInternalName(c)) Then osel.Add c
'    Next
'    If osel.count > 0 Then osel.Delete
'    CATIA.HSOSynchronized = True
'    CATIA.RefreshDisplay = True
'End Sub
'
'Private Function Resolve(osel, obj) As Object
'    Set Resolve = Nothing
'    On Error Resume Next
'    osel.Clear: osel.Add obj
'    If Err.Number = 0 Then Set Resolve = osel.item(1).Value
'    osel.Clear: Err.Clear
'    On Error GoTo 0
'End Function
'
'
'
'
'
'
'
'
'
'
'
''===================
'
'
''==============================================================================
'' rmCrv - 删除未被 Sweep 引用的曲线
''
'' 设计哲学:
''   1. 一次遍历，多次过滤 — 递归只发生一次
''   2. 阶段分离 — 收集、分析、行动，各司其职
''   3. 魔法数字命名 — 意图自明
''   4. 副作用隔离 — 删除操作集中在最后一步
''==============================================================================
'Private Const TYPE_CURVE As Long = 3
'Private Const TYPE_SWEEP As Long = 7
'
'Sub rmCrv()
'    If CATIA.Windows.count < 1 Then MsgBox "没有打开的窗口": Exit Sub
'
'    Dim oDoc:  Set oDoc = CATIA.ActiveDocument
'    Dim oprt:  Set oprt = KCL.get_workPartDoc.part
'    Dim osel:  Set osel = oDoc.Selection
'    Dim HSF:   Set HSF = oDoc.part.HybridShapeFactory
'
'    ' ── Phase 1: 一次遍历，收集所有候选 Shape ──
'    Dim allShapes: Set allShapes = KCL.Initlst
'   Gather allShapes, oprt, HSF
'
'    ' ── Phase 2: 从 Sweep 中提取被引用的曲线 ──
'    Dim Dict_ref: Set Dict_ref = KCL.InitDic
'    Dim shp, crv
'
'    For Each shp In allShapes
'        If HSF.GetGeometricalFeatureType(shp) = TYPE_SWEEP Then
'            On Error Resume Next
'            Set shp = Resolve(osel, shp)
'            Set crv = Nothing ' 泛型 → 真实类型
'            Set crv = Resolve(osel, shp.FirstGuideCrv)      ' Reference → Shape
'            If Not crv Is Nothing Then Dict_ref(KCL.GetInternalName(crv)) = 1
'            Error.Clear
'            On Error GoTo 0
'        End If
'    Next
'
'    ' ── Phase 3: 收集未被引用的曲线 ──
'    Dim toDelete: Set toDelete = KCL.Initlst
'    For Each shp In allShapes
'        If HSF.GetGeometricalFeatureType(shp) = TYPE_CURVE Then
'            If Not Dict_ref.Exists(KCL.GetInternalName(shp)) Then toDelete.Add shp
'        End If
'    Next
'
'    ' ── Phase 4: 集中删除 ──
'    CATIA.RefreshDisplay = False
'    Dim itm
'    For Each itm In toDelete
'        osel.Clear: osel.Add itm: osel.Delete
'    Next
'    osel.Clear
'    CATIA.RefreshDisplay = True
'End Sub
'
''──────────────────────────────────────────────────────────────────────────────
'' Gather — 递归收集所有 Curve 和 Sweep，递归只在这一处发生
''──────────────────────────────────────────────────────────────────────────────
'Private Sub Gather(lst, iHB, HSF)
'    On Error Resume Next
'    Dim shps: Set shps = iHB.HybridShapes
'    Dim shp, t
'    If Not shps Is Nothing Then
'        For Each shp In shps
'            t = HSF.GetGeometricalFeatureType(shp)
'            If t = TYPE_CURVE Or t = TYPE_SWEEP Then lst.Add shp
'        Next
'    End If
'    Dim chb
'    For Each chb In iHB.HybridBodies
'        Gather lst, chb, HSF
'    Next
'    On Error GoTo 0
'End Sub
'
''──────────────────────────────────────────────────────────────────────────────
'' Resolve — 将 CATIA 泛型对象/Reference 解析为真实类型对象
''           封装 Selection 黑魔法，使调用处意图清晰
''──────────────────────────────────────────────────────────────────────────────
'Private Function Resolve(osel, obj) As Object
'    Set Resolve = Nothing
'    On Error Resume Next
'    osel.Clear: osel.Add obj
'    If Err.Number = 0 Then Set Resolve = osel.item(1).Value
'    osel.Clear: Err.Clear
'    On Error GoTo 0
'End Function
'
'
'
