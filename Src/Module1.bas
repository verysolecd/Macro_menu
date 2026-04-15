Attribute VB_Name = "Module1"
Private Const TYPE_CURVE As Long = 3
Private Const TYPE_SWEEP As Long = 7

Sub rmCrv()
    If CATIA.Windows.count < 1 Then MsgBox "没有打开的窗口": Exit Sub
    Dim oDoc:  Set oDoc = CATIA.ActiveDocument
    Dim oprt:  Set oprt = KCL.get_workPartDoc.part
    Dim osel:  Set osel = oDoc.Selection
    Dim HSF:   Set HSF = oDoc.part.HybridShapeFactory
    CATIA.RefreshDisplay = False
    CATIA.HSOSynchronized = False
    ' ══ Step 1: Search 获取所有 Sweep，提取引用曲线 ══
    Dim refSet: Set refSet = KCL.InitDic
    Dim sweeps: Set sweeps = KCL.Initlst
    osel.Clear
    osel.Search "CATGMOSearch.Surface,all" ' Sweep 是曲面的子类型,或许可以使用
    'osel.Search "CATPrtSearch.HybridShapeSweep,all"
    Dim i
    For i = 1 To osel.count
        Dim shp: Set shp = osel.item(i).Value
        If HSF.GetGeometricalFeatureType(shp) = TYPE_SWEEP Then sweeps.Add shp
    Next
    ' 解析每个 Sweep 引用的引导线（Reference → Shape 仍需 Resolve）
    Dim sw, crv
    For Each sw In sweeps
        Set crv = Resolve(osel, sw.FirstGuideCrv)
        If Not crv Is Nothing Then refSet(KCL.GetInternalName(crv)) = 1
    Next
    ' ══ Step 2: Search 获取所有曲线 ══
    osel.Clear
    osel.Search "CATGMOSearch.Curve,all"
    Dim curves: Set curves = KCL.Initlst
    For i = 1 To osel.count
        curves.Add osel.item(i).Value
    Next
    ' ══ Step 3: 批量选中未引用曲线 → 一次删除 ══
    osel.Clear
    Dim c
    For Each c In curves
        If Not refSet.Exists(KCL.GetInternalName(c)) Then osel.Add c
    Next
    If osel.count > 0 Then osel.Delete
    CATIA.HSOSynchronized = True
    CATIA.RefreshDisplay = True
End Sub

Private Function Resolve(osel, obj) As Object
    Set Resolve = Nothing
    On Error Resume Next
    osel.Clear: osel.Add obj
    If Err.Number = 0 Then Set Resolve = osel.item(1).Value
    osel.Clear: Err.Clear
    On Error GoTo 0
End Function

