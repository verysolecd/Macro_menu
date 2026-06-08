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
' %UI Button allshow        显示所有产品  #228B22
' %UI Button allhide        隐藏所有产品  #6B7280
' %UI Button sel_child_show 显示选定(含子树)  #1E88E5
' %UI Button ONLYsel_child_show 仅显示选定(含子树)  #1E88E5
' %UI Button onlyselshow    仅显示选定  #1976D2
' %UI Button onlysel_hide   隐藏选定产品  #757575
' %UI Label lbL_4 ------
' %UI Button AsmHide_Plns   隐藏所有平面  #78909C
' %UI Button AsmHide_axis   隐藏所有坐标系  #90A4AE
' %UI Button AsmHide_GS     隐藏产品几何集  #B0BEC5
' %UI Label lbL_5          --以下针对零件--
' %UI Button PrtHide_Skt     隐藏所有草图  #5C6BC0
' %UI Button PrtHide_root_GS 隐藏part几何集  #7986CB
' %UI Button PrtShow_GS      显示part几何集  #43A047
' %UI Button PrtShow_selGS   显示选定及子集  #43A047
' %UI Button Prtselgeoshow   仅显示选定及子集  #43A047

Option Explicit

Private mWD As Cls_DynaWD
Private Const mdlname As String = "OTH_ivhideshow"
Sub setHideshow()
    If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    '==生成UItoolbar-===================
    Set mWD = New Cls_DynaWD
    mWD.ShowToolbar mdlname
End Sub
Sub sel_child_show_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    If sel.count = 0 Then Exit Sub
    CATIA.RefreshDisplay = False
    On Error Resume Next
    Dim i As Integer, parentPrd As Object
    Dim selArray() As Object
    ReDim selArray(sel.count - 1)
    For i = 1 To sel.count
        Set selArray(i - 1) = sel.item(i).Value
    Next
    sel.Clear
    ' 第一步：全场大范围静默隐藏
'    sel.Search "CATProductSearch.Product,all"
'    If sel.count > 0 Then sel.VisProperties.SetShow 1
'    sel.Clear
    ' 第二步：为了让这几根树苗显现，把它们的树干通道全部显示
    For i = LBound(selArray) To UBound(selArray)
        Set parentPrd = selArray(i).Parent
        Do While TypeName(parentPrd) = "Product" Or TypeName(parentPrd) = "Products"
            If TypeName(parentPrd) = "Product" Then sel.Add parentPrd
            Set parentPrd = parentPrd.Parent
        Loop
    Next
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    ' 第三步：利用底层Search的 ",sel" 条件过滤出其所有几千个子实体
    For i = LBound(selArray) To UBound(selArray)
        sel.Add selArray(i)
    Next
    sel.Search "CATProductSearch.Product,sel"
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub

Sub ONLYsel_child_show()
 Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    If sel.count = 0 Then Exit Sub
    CATIA.RefreshDisplay = False
    On Error Resume Next
    Dim i As Integer, parentPrd As Object
    Dim selArray() As Object
    ReDim selArray(sel.count - 1)
    For i = 1 To sel.count
        Set selArray(i - 1) = sel.item(i).Value
    Next
    sel.Clear
'第一步:      全场大范围静默隐藏
    sel.Search "CATProductSearch.Product,all"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
'第二步：为了让这几根树苗显现，把它们的树干通道全部显示
    For i = LBound(selArray) To UBound(selArray)
        Set parentPrd = selArray(i).Parent
        Do While TypeName(parentPrd) = "Product" Or TypeName(parentPrd) = "Products"
            If TypeName(parentPrd) = "Product" Then sel.Add parentPrd
            Set parentPrd = parentPrd.Parent
        Loop
    Next
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    ' 第三步：利用底层Search的 ",sel" 条件过滤出其所有几千个子实体
    For i = LBound(selArray) To UBound(selArray)
        sel.Add selArray(i)
    Next
    sel.Search "CATProductSearch.Product,sel"
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub

Sub onlyselshow_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    If sel.count = 0 Then Exit Sub
    CATIA.RefreshDisplay = False
    On Error Resume Next
    Dim i As Integer, parentPrd As Object
    Dim selArray() As Object
    ReDim selArray(sel.count - 1)
    For i = 1 To sel.count
        Set selArray(i - 1) = sel.item(i).Value
    Next
    sel.Clear
    ' 全场隐藏
    sel.Search "CATProductSearch.Product,all"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    ' 恢复那几个孤点和它们的父亲
    For i = LBound(selArray) To UBound(selArray)
        sel.Add selArray(i)
        Set parentPrd = selArray(i).Parent
        Do While TypeName(parentPrd) = "Product" Or TypeName(parentPrd) = "Products"
            If TypeName(parentPrd) = "Product" Then sel.Add parentPrd
            Set parentPrd = parentPrd.Parent
        Loop
    Next
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub allshow_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    CATIA.RefreshDisplay = False
    On Error Resume Next
    sel.Clear
    sel.Search "CATProductSearch.Product.Visibility=Hidden,all"
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub allhide_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    CATIA.RefreshDisplay = False
    On Error Resume Next
    sel.Clear
    sel.Search "CATProductSearch.Product.Visibility=Visible,all"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub onlysel_hide_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    If sel.count = 0 Then Exit Sub
    CATIA.RefreshDisplay = False
    On Error Resume Next
    sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub AsmHide_Plns_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    CATIA.RefreshDisplay = False
    On Error Resume Next
    sel.Clear
    sel.Search "CATPrtSearch.Plane,all"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub AsmHide_axis_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    CATIA.RefreshDisplay = False
    On Error Resume Next
    sel.Clear
    sel.Search "CATPrtSearch.AxisSystem,all"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub AsmHide_GS_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    CATIA.RefreshDisplay = False
    On Error Resume Next
    sel.Clear
    sel.Search "CATPrtSearch.OpenBodyFeature,all"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub PrtHide_root_GS_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    On Error Resume Next
    Dim oprt As Object
    Set oprt = KCL.get_workPartDoc.part
    If oprt Is Nothing Then Exit Sub
    CATIA.RefreshDisplay = False
    sel.Clear
    Dim itm As Object
    For Each itm In oprt.HybridBodies
        sel.Add itm
    Next
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub
Sub PrtShow_GS_click()
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    On Error Resume Next
    Dim oprt As Object
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
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    On Error Resume Next
    Dim oprt As Object
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
    Dim sel As Object
    Set sel = CATIA.ActiveDocument.Selection
    If sel.Count2 = 0 Then Exit Sub
    On Error Resume Next
    Dim oprt As Object
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
Sub Prtselgeoshow_click()
    Dim sel: Set sel = CATIA.ActiveDocument.Selection
    If sel.count = 0 Then Exit Sub
    Dim oprt: Set oprt = KCL.get_workPartDoc.part
    If oprt Is Nothing Then Exit Sub
    
    CATIA.RefreshDisplay = False
    On Error Resume Next
    Dim i As Integer, parentGEO As Object
    Dim selArray() As Object
    ReDim selArray(sel.count - 1)
    For i = 1 To sel.count
        Set selArray(i - 1) = sel.item(i).Value
    Next
    sel.Clear
    sel.Add oprt
    sel.Search "CATPrtSearch.OpenBodyFeature,sel"
    If sel.count > 0 Then sel.VisProperties.SetShow 1
    sel.Clear
    ' 恢复那几个孤点和它们的父亲
    Dim mlst: Set mlst = KCL.Initlst
    For i = LBound(selArray) To UBound(selArray)
        sel.Add selArray(i)
        Set parentGEO = selArray(i)
        Do While TypeName(parentGEO) = "HybridBody" Or TypeName(parentGEO) = "HybridBodies"
            mlst.Add parentGEO
            Set parentGEO = parentGEO.Parent
        Loop
    Next
    Dim gs
     For i = LBound(selArray) To UBound(selArray)
        Set gs = selArray(i)
            recurGEO2lst gs, mlst
    Next
    Dim itm
    For Each itm In mlst
        sel.Add itm
    Next
    If sel.count > 0 Then sel.VisProperties.SetShow 0
    sel.Clear
    CATIA.RefreshDisplay = True
End Sub

