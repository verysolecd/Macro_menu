Attribute VB_Name = "MDL_pt2Hb"
'{GP:4}
'{EP:Copypointsfromset}
'{Caption:复制点到图形集}
'{ControlTipText: 选择几何图形集后导出其下所有点（含阵列）到新图形集}
'{BackColor:12648447}
Private Const mdlname As String = "MDL_pt2Hb"

Sub Copypointsfromset()
    If Not CanExecute("Productdocument,PartDocument") Then Exit Sub
    On Error Resume Next
        Dim oDoc: Set oDoc = CATIA.ActiveDocument
        Dim workPrtDoc: Set workPrtDoc = KCL.get_workPartDoc
        Dim oprt: Set oprt = Nothing: Set oprt = workPrtDoc.part
    Err.Clear
    On Error GoTo 0
    If IsNothing(oprt) Then: MsgBox "No activated Part": Exit Sub
    
    
    Set HSF = oprt.HybridShapeFactory
    '======= 选择源几何图形集
    Dim iSel: Set iSel = Nothing
    Dim imsg: imsg = "请选择一个几何图形集"
    Dim filter(0): filter(0) = "HybridBody"
    On Error Resume Next
        Set iSel = KCL.SelectItem(imsg, filter)
    On Error GoTo 0
    If iSel Is Nothing Then MsgBox "操作取消": Exit Sub
    Dim osel As Selection
    Set osel = CATIA.ActiveDocument.Selection
    Dim oTempHb As HybridBody
    Set oTempHb = oprt.HybridBodies.Add()
    oTempHb.Name = "_temp"
    osel.Clear
    osel.Add iSel
    osel.Copy                          ' 复制整个源图形集

    osel.Clear
    osel.Add oTempHb
    osel.PasteSpecial "CATPrtResultWithOutLink"  ' 粘贴为无链接结果 → 阵列被展开！
    osel.Clear
    oprt.Update

    '======= Step2: 在临时集中搜索所有点（此时阵列已展开为独立点）
    Dim oTargetHb As HybridBody
    Set oTargetHb = oprt.HybridBodies.Add()
    oTargetHb.Name = "extracted points"

    osel.Clear
    osel.Add oTempHb
    osel.Search ".Point,sel"

    Dim i As Integer: i = 1
    Dim j As Integer
    For j = 1 To osel.count
        On Error Resume Next
        Dim oRef As Reference
        Set oRef = oprt.CreateReferenceFromObject(osel.item(j).Value)
        If Not oRef Is Nothing Then
            Dim oPt As HybridShapePointExplicit
            Set oPt = HSF.AddNewPointDatum(oRef)
            oPt.Name = "pt_" & i
            oTargetHb.AppendHybridShape oPt
            i = i + 1
        End If
        Set oRef = Nothing
        On Error GoTo 0
    Next j
    osel.Clear

    '======= Step3: 删除临时集，更新
    osel.Add oTempHb
    osel.Delete
    osel.Clear
    oprt.Update
    
'     '======= Step4: 批量选中目标集中所有点，统一设置符号为圆圈 ?
'    osel.Clear
'    osel.Add oTargetHb
'    osel.Search ".Point,sel"
'    If osel.count > 0 Then
'        Dim oVisProp As VisPropertySet
'        Set oVisProp = osel.VisProperties
'        oVisProp.SetSymbolType catVisPropertySymbolCircle   ' 圆圈符号
'    End If
'    osel.Clear
    MsgBox "完成！共复制 " & (i - 1) & " 个点到 'extracted points' 图形集，符号已设为圆圈。"
    Set iSel = Nothing
    Set osel = Nothing
    
End Sub


