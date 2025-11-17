Attribute VB_Name = "MDL_Part2Product"
'Attribute VB_Name = "m2_Part2Product"
'{控件提示文本: 可将零件转换为产品}
' 检查零件文档中是否存在左手坐标系
'{Gp:4}
'{Ep:CATMain}
'{Caption:零件转产品}
'{ControlTipText:此按钮将多实体零件转化为产品}
'{BackColor:}
Option Explicit

Sub CATMain()
    If Not CanExecute("PartDocument") Then Exit Sub
    Dim BaseDoc As PartDocument: Set BaseDoc = CATIA.ActiveDocument
    Dim BasePath As Variant: BasePath = Array(BaseDoc.FullName)
    Dim pt As part: Set pt = BaseDoc.part
    Dim LeafItems As collection: Set LeafItems = Get_LeafItemLst(pt.bodies)
    Dim msg As String
    If LeafItems Is Nothing Then
        msg = "没有可复制的元素！"
        MsgBox msg, vbOKOnly + vbExclamation
        Exit Sub
    End If
    msg = LeafItems.count & " 个可复制的元素。" & vbNewLine & _
          "请指定粘贴的类型" & vbNewLine & vbNewLine & _
          "是 : 带链接的结果(As Result With Link)" & vbNewLine & _
          "否 : 作为结果(As Result)" & vbNewLine & _
          "取消 : 宏中止"
    Dim PasteType As String
    Select Case MsgBox(msg, vbQuestion + vbYesNoCancel)
        Case vbYes
            PasteType = "CATPrtResult"
        Case vbNo
            PasteType = "CATPrtResultWithOutLink"
        Case Else
            Exit Sub
    End Select
    KCL.SW_Start
    Dim BaseScene As Variant: BaseScene = GetScene3D(GetViewPnt3D())
    Dim TopDoc As ProductDocument: Set TopDoc = CATIA.Documents.Add("Product")
    Call ToProduct(TopDoc, LeafItems, PasteType)
    Call UpdateScene(BaseScene)
    TopDoc.Product.Update
    Debug.Print "时间:" & KCL.SW_GetTime & "s"
    MsgBox "完成"
End Sub

Private Sub ToProduct(ByVal TopDoc As ProductDocument, _
                      ByVal LeafItems As collection, _
                      ByVal PasteType As String)
    Dim TopSel As Selection
    Set TopSel = TopDoc.Selection
    Dim BaseSel As Selection
    Set BaseSel = KCL.GetParent_Of_T(LeafItems(1), "PartDocument").Selection
    Dim prods As Products
    Set prods = TopDoc.Product.Products
    Dim itm As AnyObject
    Dim TgtDoc As PartDocument
    Dim ProdsNameDic As Object: Set ProdsNameDic = KCL.InitDic()
    CATIA.HSOSynchronized = False
    For Each itm In LeafItems
        If ProdsNameDic.Exists(itm.Name) Then
            Set TgtDoc = ProdsNameDic.item(itm.Name)
        Else
            Set TgtDoc = Init_Part(prods, itm.Name)
            ProdsNameDic.Add itm.Name, TgtDoc
        End If
        Call Preparing_Copy(BaseSel, itm)
        With BaseSel
            .Copy
            .Clear
        End With
        With TopSel
            .Clear
            .Add TgtDoc.part
            .PasteSpecial PasteType
        End With
    Next
    BaseSel.Clear
    TopSel.Clear
    CATIA.HSOSynchronized = True
End Sub

Private Sub Preparing_Copy(ByVal sel As Selection, ByVal itm As AnyObject)
    sel.Clear
    If TypeName(itm) = "Body" Then
        sel.Add itm
        Exit Sub
    End If
    Dim ShpsLst As collection: Set ShpsLst = New collection
    ShpsLst.Add itm.HybridShapes
    Select Case TypeName(itm)
        Case "HybridBody"
            Set ShpsLst = Get_All_HbShapes(itm, ShpsLst)
        Case "OrderedGeometricalSet"
            Set ShpsLst = Get_All_OdrGeoSetShapes(itm, ShpsLst)
    End Select
    Dim Shps As HybridShapes, Shp As HybridShape
    For Each Shps In ShpsLst
        For Each Shp In Shps
            sel.Add Shp
        Next
    Next
End Sub

Private Function Get_All_OdrGeoSetShapes(ByVal OdrGeoSet As OrderedGeometricalSet, _
                                         ByVal lst As collection) As collection
    Dim child As OrderedGeometricalSet
    For Each child In OdrGeoSet.OrderedGeometricalSets
        lst.Add child.HybridShapes
        If child.OrderedGeometricalSets.count > 0 Then
            Set lst = Get_All_OdrGeoSetShapes(child, lst)
        End If
    Next
    Set Get_All_OdrGeoSetShapes = lst
End Function

Private Function Get_All_HbShapes(ByVal Hbdy As HybridBody, _
                                  ByVal lst As collection) As collection
    Dim child As HybridBody
    For Each child In Hbdy.HybridBodies
        lst.Add child.HybridShapes
        If child.HybridBodies.count > 0 Then
            Set lst = Get_All_HbShapes(child, lst)
        End If
    Next
    Set Get_All_HbShapes = lst
End Function

Private Function Get_LeafItemLst(ByVal pt As part) As collection
    Set Get_LeafItemLst = Nothing
    Dim sel As Selection: Set sel = pt.Parent.Selection
    Dim TmpLst As collection: Set TmpLst = New collection
    Dim i As Long
    Dim filter As String
    filter = "(CATPrtSearch.BodyFeature.Visibility=Shown " & _
            "+ CATPrtSearch.OpenBodyFeature.Visibility=Shown" & _
            "+ CATPrtSearch.MMOrderedGeometricalSet.Visibility=Shown),sel"
    CATIA.HSOSynchronized = False
    With sel
        .Clear
        .Add pt
        .Search filter
        For i = 1 To .Count2
            TmpLst.Add .item(i).value
        Next
        .Clear
    End With
    CATIA.HSOSynchronized = True
    If TmpLst.count < 1 Then Exit Function
    Dim LeafHBdys As Object: Set LeafHBdys = KCL.InitDic()
    Dim Hbdy As AnyObject
    For Each Hbdy In pt.HybridBodies
        LeafHBdys.Add Hbdy, 0
    Next
    For Each Hbdy In pt.OrderedGeometricalSets
        LeafHBdys.Add Hbdy, 0
    Next
    Dim itm As AnyObject
    Dim lst As collection: Set lst = New collection
    For Each itm In TmpLst
        Select Case TypeName(itm)
            Case "Body"
                If Is_LeafBody(itm) Then lst.Add itm
            Case Else
                If Is_LeafHybridBody(itm, LeafHBdys) Then lst.Add itm
        End Select
    Next
    If lst.count < 1 Then Exit Function
    Set Get_LeafItemLst = lst
End Function

Private Function Is_LeafBody(ByVal bdy As body) As Boolean
    Is_LeafBody = bdy.InBooleanOperation = False And bdy.Shapes.count > 0
End Function

Private Function Is_LeafHybridBody(ByVal Hbdy As AnyObject, _
                                   ByVal Dic As Object) As Boolean
    Is_LeafHybridBody = False
    If Not Dic.Exists(Hbdy) Then Exit Function
    CATIA.HSOSynchronized = False
    Dim sel As Selection
    Set sel = KCL.GetParent_Of_T(Hbdy, "PartDocument").Selection
    Dim cnt As Long
    With sel
        .Clear
        .Add Hbdy
        .Search "Visibility=Shown,sel"
        cnt = .Count2
        .Clear
    End With
    CATIA.HSOSynchronized = True
    If cnt > 1 Then Is_LeafHybridBody = True
End Function

Private Function Init_Part(ByVal prods As Variant, _
                           ByVal PtNum As String) As PartDocument
    Dim prod As Product
    On Error Resume Next
        Set prod = prods.AddNewComponent("Part", PtNum)
    On Error GoTo 0
    Set Init_Part = prods.item(prods.count).ReferenceProduct.Parent
End Function

Private Sub UpdateScene(ByVal Scene As Variant)
    Dim viewer As Viewer3D: Set viewer = CATIA.ActiveWindow.ActiveViewer
    Dim VPnt3D As Variant
    Set VPnt3D = viewer.Viewpoint3D
    Dim ary As Variant
    ary = GetRangeAry(Scene, 0, 2)
    Call VPnt3D.PutOrigin(ary)
    ary = GetRangeAry(Scene, 3, 5)
    Call VPnt3D.PutSightDirection(ary)
    ary = GetRangeAry(Scene, 6, 8)
    Call VPnt3D.PutUpDirection(ary)
    VPnt3D.FieldOfView = Scene(9)
    VPnt3D.FocusDistance = Scene(10)
    Call viewer.Update
End Sub

Private Function GetScene3D(ViewPnt3D As Viewpoint3D) As Variant
    Dim vp As Variant: Set vp = ViewPnt3D
    Dim origin(2) As Variant: Call vp.GetOrigin(origin)
    Dim sight(2) As Variant: Call vp.GetSightDirection(sight)
    GetScene3D = KCL.JoinAry(origin, sight)
    Dim up(2) As Variant: Call vp.GetUpDirection(up)
    GetScene3D = KCL.JoinAry(GetScene3D, up)
    Dim FieldOfView(0) As Variant: FieldOfView(0) = vp.FieldOfView
    GetScene3D = KCL.JoinAry(GetScene3D, FieldOfView)
    Dim FocusDist(0) As Variant: FocusDist(0) = vp.FocusDistance
    GetScene3D = KCL.JoinAry(GetScene3D, FocusDist)
End Function

Private Function GetViewPnt3D() As Viewpoint3D
    Set GetViewPnt3D = CATIA.ActiveWindow.ActiveViewer.Viewpoint3D
End Function




