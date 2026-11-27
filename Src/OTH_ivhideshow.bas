Attribute VB_Name = "OTH_ivhideshow"
'{GP:6}
'{Ep:setHideshow}
'{Caption:反选隐藏}
'{ControlTipText:反选并隐藏结构树}
'{BackColor:}

'标题格式为 %Title <Caption/Text>
'%Title 如何配置?
'------控件清单--------------------------------------------------
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lbL_jpzcs       键盘造车手出品
' %UI Button ASMshow_allprd       显示所有产品  #228B22
' %UI Button ASMhide_allprd        隐藏所有产品  #6B7280
' %UI Button Asmshow_selChild 显示选定产品(子树)  #1E88E5
' %UI Button Asmshow_onlySelChild 仅显示选定产品(子树)  #1E88E5
' %UI Label lbL_4 ------
' %UI Button AsmHide_allPlns   隐藏所有平面  #78909C
' %UI Button AsmHide_allaxis   隐藏所有坐标系  #90A4AE
' %UI Button AsmHide_allGS     隐藏所有几何集  #B0BEC5
' %UI Button Asmshow_allGS     显示所有几何集  #B0BEC5
' %UI Label lbL_5          --以下针对零件--
' %UI Button PrtHide_Skt     隐藏所有草图  #5C6BC0
' %UI Button PrtHide_root_GS 隐藏part几何集  #7986CB
' %UI Button PrtShow_allGS      显示part几何集  #43A047
' %UI Button PrtShow_selGS   显示选定及子集  #43A047
' %UI Button PrtShow_onlyselGS   仅显示选定及子集  #43A047

Option Explicit

Private mWD As Cls_DynaWD
Private Const mdlname As String = "OTH_ivhideshow"
Private mQuery As String
Sub setHideshow()
    If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    '==生成UItoolbar-===================
    Set mWD = New Cls_DynaWD
    mWD.ShowToolbar mdlname
End Sub
Private Sub showQuery(bIsVisible As Boolean)
    Dim sel: Set sel = CATIA.ActiveDocument.Selection
    Dim toshow: toshow = IIf(bIsVisible, 0, 1)
    CATIA.RefreshDisplay = False
    On Error Resume Next
        sel.Search mQuery
        If sel.count > 0 Then sel.VisProperties.SetShow toshow: sel.Clear
    On Error GoTo 0
    CATIA.RefreshDisplay = True
End Sub

Sub ASMshow_allprd_click()
    mQuery = "CATProductSearch.Product.Visibility=Hidden,all"
    showQuery (True)
End Sub
Sub ASMhide_allprd_click()
    mQuery = "CATProductSearch.Product.Visibility=Visible,all"
    showQuery (False)
End Sub
Sub AsmHide_allPlns_click()
     mQuery = "CATPrtSearch.Plane,all"
    showQuery (False)
    Asmhide_namedPlane
End Sub
Private Sub Asmhide_namedPlane()
    mQuery = "((CATPrtSearch.Plane + CATPrtSearch.Surface) & (Name='*平面*' + Name='*plane*')),all"
    showQuery (False)
End Sub
Sub AsmHide_allaxis_click()
     mQuery = "CATPrtSearch.AxisSystem,all"
    showQuery (False)
End Sub
Sub AsmHide_allGS_click()
    mQuery = "CATPrtSearch.OpenBodyFeature,all"
    showQuery (False)
End Sub
Sub Asmshow_allGS_click()
    mQuery = "CATPrtSearch.OpenBodyFeature,all"
    showQuery (True)
End Sub

Sub Asmshow_selChild_click()
    Dim vSaved As Variant
    vSaved = GetSavedSelection()
    If Not IsArray(vSaved) Then Exit Sub
    CATIA.RefreshDisplay = False
    Dim i As Integer, parentPrd As Object
    Dim sel As Selection
    Set sel = CATIA.ActiveDocument.Selection
    For i = LBound(vSaved) To UBound(vSaved)
        sel.Add vSaved(i)
        Set parentPrd = vSaved(i).Parent
        Do While TypeName(parentPrd) = "Product" Or TypeName(parentPrd) = "Products"
            If TypeName(parentPrd) = "Product" Then sel.Add parentPrd
            Set parentPrd = parentPrd.Parent
        Loop
    Next i
    If sel.count > 0 Then sel.VisProperties.SetShow 0: sel.Clear
    ' 步：利用底层Search的 ",sel" 条件过滤出其子产品
    For i = LBound(vSaved) To UBound(vSaved)
        sel.Add vSaved(i)
    Next
    mQuery = "CATProductSearch.Product,sel"  '注意此处范围为sel
    showQuery (True)
End Sub


' 隔离功能：仅显示选中的零部件（及其子级），隐藏其他所有无关元素
Sub Asmshow_onlySelChild_click()
    Dim vSaved As Variant
    vSaved = GetSavedSelection()
    If Not IsArray(vSaved) Then Exit Sub
    CATIA.RefreshDisplay = False
    On Error Resume Next
    ' 2. 全场隐藏所有产品节点
    mQuery = "CATProductSearch.Product,all"
    showQuery False
    ' 3. 恢复选中对象及其所有父级装配路径
    Dim sel As Selection
    Set sel = CATIA.ActiveDocument.Selection
    Dim i As Long, parentPrd As Object
    For i = LBound(vSaved) To UBound(vSaved)
        sel.Add vSaved(i)
        Set parentPrd = vSaved(i).Parent
        Do While TypeName(parentPrd) = "Product" Or TypeName(parentPrd) = "Products"
            If TypeName(parentPrd) = "Product" Then sel.Add parentPrd
            Set parentPrd = parentPrd.Parent
        Loop
    Next i
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    ' 4. 若选中的是子装配体，还需向下恢复其所有子零件的显示状态
    sel.Clear
    For i = LBound(vSaved) To UBound(vSaved)
        sel.Add vSaved(i)
    Next i
    mQuery = "CATProductSearch.Product,sel"
    showQuery (True)
End Sub

Sub PrtHide_root_GS_click()
    Dim sel, oprt, itm As Object
    On Error Resume Next
        Set oprt = KCL.get_workPartDoc.part
        If oprt Is Nothing Then Exit Sub
    On Error GoTo 0
    CATIA.RefreshDisplay = False
    Set sel = CATIA.ActiveDocument.Selection
    sel.Clear
    For Each itm In oprt.HybridBodies
        sel.Add itm
    Next
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub PrtShow_allGS_click()
    Dim sel, oprt, itm As Object
    Set sel = CATIA.ActiveDocument.Selection
    On Error Resume Next
        Set oprt = KCL.get_workPartDoc.part
        If oprt Is Nothing Then Exit Sub
    CATIA.RefreshDisplay = False
    sel.Clear
    sel.Add oprt
    sel.Search "CATPrtSearch.OpenBodyFeature,sel"
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub PrtHide_Skt_click()
    Dim sel, oprt, itm As Object
    Set sel = CATIA.ActiveDocument.Selection
    On Error Resume Next
    Set oprt = KCL.get_workPartDoc.part
    If oprt Is Nothing Then Exit Sub
    CATIA.RefreshDisplay = False
    sel.Clear
    sel.Add oprt
    sel.Search "CATPrtSearch.Sketch,sel"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub PrtShow_selGS_click()
    Dim sel, oprt As Object
    Set sel = CATIA.ActiveDocument.Selection
    If sel.Count2 = 0 Then Exit Sub
    On Error Resume Next
    Set oprt = KCL.get_workPartDoc.part
    If oprt Is Nothing Then Exit Sub
    CATIA.RefreshDisplay = False
    Dim mlst, i, gs, itm
    Set mlst = KCL.Initlst
    If sel.Count2 > 0 Then
    For i = 1 To sel.Count2
        Set gs = sel.item(i).Value
        Call recurGEO2lst(gs, mlst)
    Next
    End If
    sel.Clear
    For Each itm In mlst
        sel.Add itm
    Next
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub

Private Sub recurGEO2lst(gs, Optional ByRef lst)
    If lst Is Nothing Then Set lst = KCL.Initlst
    lst.Add gs
    Dim shape, skt, itm
    For Each shape In gs.HybridShapes
        lst.Add shape
    Next
    For Each skt In gs.HybridSketches
         lst.Add skt
    Next
    If gs.HybridBodies.count <> 0 Then
        For Each itm In gs.HybridBodies
            Call recurGEO2lst(itm, lst)
        Next
    End If
End Sub
Sub PrtShow_onlyselGS()
    Dim selArray As Variant
    selArray = GetSavedSelection()
    If IsEmpty(selArray) Then Exit Sub
    Dim oprt As part:    Set oprt = KCL.get_workPartDoc.part
    If oprt Is Nothing Then Exit Sub
    Dim sel As Selection:    Set sel = CATIA.ActiveDocument.Selection
    CATIA.RefreshDisplay = False
    On Error GoTo Cleanup
    ' 1. 隐藏当前零件下的所有几何图形集
    sel.Add oprt
    sel.Search "CATPrtSearch.OpenBodyFeature,sel"
    If sel.Count2 > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    ' 2. 分别收集“父级路径上的 GEO”和“选中项及其子级”
    Dim parentLst, currObj, childLst As Object
    Set parentLst = KCL.Initlst: Set childLst = KCL.Initlst
    Dim i As Long
    On Error Resume Next
    For i = LBound(selArray) To UBound(selArray)
        childLst.Add selArray(i)
        Set currObj = selArray(i).Parent
        Do While Not currObj Is Nothing
            If currObj Is currObj.Parent Then Exit Do  ' 防止死循环（当到达 Application 或 Document 顶层时 Parent 指向自身）
            If TypeName(currObj) = "HybridBodies" Or TypeName(currObj) = "HybridBody" Then parentLst.Add currObj
            If TypeName(currObj) = "Part" Then Exit Do
            Set currObj = currObj.Parent
        Loop
            recurGEO2lst selArray(i), childLst ' 向下递归收集子元素
    Next i
    ' 3. 【关键】先强制显示所有父级几何图形集（从根到当前父级）
    Dim itm As Variant
    sel.Clear
    For Each itm In parentLst
        If Not itm Is Nothing Then sel.Add itm
    Next itm
    If sel.Count2 > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    ' 4. 再显示选中的几何图形集/元素及其内部子项
    For Each itm In childLst
        If Not itm Is Nothing Then sel.Add itm
    Next itm
    If sel.Count2 > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
Cleanup:
    CATIA.RefreshDisplay = True
End Sub

' 辅助函数：备份当前选择集并清空（返回 Variant 数组，无选中时返回 Empty）
Private Function GetSavedSelection() As Variant
    Dim sel As Selection:    Set sel = CATIA.ActiveDocument.Selection
    If sel.Count2 = 0 Then Exit Function
    Dim i As Long
    Dim oselArray() As Object
    ReDim oselArray(sel.Count2 - 1)
    For i = 1 To sel.Count2
        Set oselArray(i - 1) = sel.Item2(i).Value
    Next i
    sel.Clear
    GetSavedSelection = oselArray
End Function


