Attribute VB_Name = "m21_Part2Product"
'Attribute VB_Name = "m2_Part2Product"
'{控件提示文本: 可将零件转换为产品}
' 检查零件文档中是否存在左手坐标系
'{Gp:2}
'{Ep:CATMain}
'{Caption:零件转产品}
'{ControlTipText:此按钮将多实体零件转化为产品}
'{BackColor:33023}
Option Explicit

Sub CATMain()
    If Not CanExecute("PartDocument") Then Exit Sub
    Dim BaseDoc As PartDocument: Set BaseDoc = CATIA.Activedocument
    Dim BasePath As Variant: BasePath = Array(BaseDoc.FullName)
    Dim Pt As Part: Set Pt = BaseDoc.Part
    Dim LeafItems As collection: Set LeafItems = Get_LeafItemLst(Pt.Bodies)
    Dim Msg As String
    If LeafItems Is Nothing Then
        Msg = "没有可复制的元素！"
        MsgBox Msg, vbOKOnly + vbExclamation
        Exit Sub
    End If
    Msg = LeafItems.Count & " 个可复制的元素。" & vbNewLine & _
          "请指定粘贴的类型" & vbNewLine & vbNewLine & _
          "是 : 带链接的结果(As Result With Link)" & vbNewLine & _
          "否 : 作为结果(As Result)" & vbNewLine & _
          "取消 : 宏中止"
    Dim PasteType As String
    Select Case MsgBox(Msg, vbQuestion + vbYesNoCancel)
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
    TopDoc.product.Update
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
    Dim Prods As Products
    Set Prods = TopDoc.product.Products
    Dim Itm As AnyObject
    Dim TgtDoc As PartDocument
    Dim ProdsNameDic As Object: Set ProdsNameDic = KCL.InitDic()
    CATIA.HSOSynchronized = False
    For Each Itm In LeafItems
        If ProdsNameDic.Exists(Itm.Name) Then
            Set TgtDoc = ProdsNameDic.item(Itm.Name)
        Else
            Set TgtDoc = Init_Part(Prods, Itm.Name)
            ProdsNameDic.Add Itm.Name, TgtDoc
        End If
        Call Preparing_Copy(BaseSel, Itm)
        With BaseSel
            .Copy
            .Clear
        End With
        With TopSel
            .Clear
            .Add TgtDoc.Part
            .PasteSpecial PasteType
        End With
    Next
    BaseSel.Clear
    TopSel.Clear
    CATIA.HSOSynchronized = True
End Sub

Private Sub Preparing_Copy(ByVal Sel As Selection, ByVal Itm As AnyObject)
    Sel.Clear
    If TypeName(Itm) = "Body" Then
        Sel.Add Itm
        Exit Sub
    End If
    Dim ShpsLst As collection: Set ShpsLst = New collection
    ShpsLst.Add Itm.HybridShapes
    Select Case TypeName(Itm)
        Case "HybridBody"
            Set ShpsLst = Get_All_HbShapes(Itm, ShpsLst)
        Case "OrderedGeometricalSet"
            Set ShpsLst = Get_All_OdrGeoSetShapes(Itm, ShpsLst)
    End Select
    Dim Shps As HybridShapes, Shp As HybridShape
    For Each Shps In ShpsLst
        For Each Shp In Shps
            Sel.Add Shp
        Next
    Next
End Sub

Private Function Get_All_OdrGeoSetShapes(ByVal OdrGeoSet As OrderedGeometricalSet, _
                                         ByVal lst As collection) As collection
    Dim child As OrderedGeometricalSet
    For Each child In OdrGeoSet.OrderedGeometricalSets
        lst.Add child.HybridShapes
        If child.OrderedGeometricalSets.Count > 0 Then
            Set lst = Get_All_OdrGeoSetShapes(child, lst)
        End If
    Next
    Set Get_All_OdrGeoSetShapes = lst
End Function

Private Function Get_All_HbShapes(ByVal Hbdy As HybridBody, _
                                  ByVal lst As collection) As collection
    Dim child As HybridBody
    For Each child In Hbdy.hybridBodies
        lst.Add child.HybridShapes
        If child.hybridBodies.Count > 0 Then
            Set lst = Get_All_HbShapes(child, lst)
        End If
    Next
    Set Get_All_HbShapes = lst
End Function

Private Function Get_LeafItemLst(ByVal Pt As Part) As collection
    Set Get_LeafItemLst = Nothing
    Dim Sel As Selection: Set Sel = Pt.Parent.Selection
    Dim TmpLst As collection: Set TmpLst = New collection
    Dim i As Long
    Dim Filter As String
    Filter = "(CATPrtSearch.BodyFeature.Visibility=Shown " & _
            "+ CATPrtSearch.OpenBodyFeature.Visibility=Shown" & _
            "+ CATPrtSearch.MMOrderedGeometricalSet.Visibility=Shown),sel"
    CATIA.HSOSynchronized = False
    With Sel
        .Clear
        .Add Pt
        .Search Filter
        For i = 1 To .Count2
            TmpLst.Add .item(i).Value
        Next
        .Clear
    End With
    CATIA.HSOSynchronized = True
    If TmpLst.Count < 1 Then Exit Function
    Dim LeafHBdys As Object: Set LeafHBdys = KCL.InitDic()
    Dim Hbdy As AnyObject
    For Each Hbdy In Pt.hybridBodies
        LeafHBdys.Add Hbdy, 0
    Next
    For Each Hbdy In Pt.OrderedGeometricalSets
        LeafHBdys.Add Hbdy, 0
    Next
    Dim Itm As AnyObject
    Dim lst As collection: Set lst = New collection
    For Each Itm In TmpLst
        Select Case TypeName(Itm)
            Case "Body"
                If Is_LeafBody(Itm) Then lst.Add Itm
            Case Else
                If Is_LeafHybridBody(Itm, LeafHBdys) Then lst.Add Itm
        End Select
    Next
    If lst.Count < 1 Then Exit Function
    Set Get_LeafItemLst = lst
End Function

Private Function Is_LeafBody(ByVal Bdy As Body) As Boolean
    Is_LeafBody = Bdy.InBooleanOperation = False And Bdy.Shapes.Count > 0
End Function

Private Function Is_LeafHybridBody(ByVal Hbdy As AnyObject, _
                                   ByVal Dic As Object) As Boolean
    Is_LeafHybridBody = False
    If Not Dic.Exists(Hbdy) Then Exit Function
    CATIA.HSOSynchronized = False
    Dim Sel As Selection
    Set Sel = KCL.GetParent_Of_T(Hbdy, "PartDocument").Selection
    Dim Cnt As Long
    With Sel
        .Clear
        .Add Hbdy
        .Search "Visibility=Shown,sel"
        Cnt = .Count2
        .Clear
    End With
    CATIA.HSOSynchronized = True
    If Cnt > 1 Then Is_LeafHybridBody = True
End Function

Private Function Init_Part(ByVal Prods As Variant, _
                           ByVal PtNum As String) As PartDocument
    Dim Prod As product
    On Error Resume Next
        Set Prod = Prods.AddNewComponent("Part", PtNum)
    On Error GoTo 0
    Set Init_Part = Prods.item(Prods.Count).ReferenceProduct.Parent
End Function

Private Sub UpdateScene(ByVal Scene As Variant)
    Dim Viewer As Viewer3D: Set Viewer = CATIA.ActiveWindow.ActiveViewer
    Dim VPnt3D As Variant
    Set VPnt3D = Viewer.Viewpoint3D
    Dim ary As Variant
    ary = GetRangeAry(Scene, 0, 2)
    Call VPnt3D.PutOrigin(ary)
    ary = GetRangeAry(Scene, 3, 5)
    Call VPnt3D.PutSightDirection(ary)
    ary = GetRangeAry(Scene, 6, 8)
    Call VPnt3D.PutUpDirection(ary)
    VPnt3D.FieldOfView = Scene(9)
    VPnt3D.FocusDistance = Scene(10)
    Call Viewer.Update
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




