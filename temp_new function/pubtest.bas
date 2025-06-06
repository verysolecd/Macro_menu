Attribute VB_Name = "pubtest"
Attribute VB_Name = "m63_minbox"
Option Explicit
Private Const MINBODYNAME = "MinimumBox" '��С��������
Private Const DMYLNG = 1000000# 'Ĭ�������
Private Enum MINMAX
    MinX = 0
    MaxX = 1
    MinY = 2
    MaxY = 3
    MinZ = 4
    Maxz = 5
End Enum
Sub CATMain()
    ' ����Ʒ
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    ' ��Ʒָ��
    Dim msg As String
    msg = "��ѡ��һ����Ʒ"
    Dim prod As Product
    Set prod = KCL.SelectItem(msg, "Product")
    If prod Is Nothing Then Exit Sub
    ' ��ȡʵ��
    Dim targetBodies As collection
    Set targetBodies = getBodies(prod)
    If targetBodies Is Nothing Then Exit Sub
    ' ������ҵ�����
    Dim workDoc As PartDocument
    Set workDoc = initPartDoc(prod)
    Dim workPt As Part
    Set workPt = workDoc.Part
    ' ����ϵ
    Dim ax As AxisSystem
    Set ax = getAxis(workDoc)
    ' �������
    Dim maxBox As Variant
    maxBox = getMaxSize_Bodies(workPt, targetBodies, ax)
    ' ������С����
    Dim minBody As body
    Set minBody = workPt.bodies.Add
    minBody.Name = "MinimumBox"
    Call changeColor(minBody)
    ' ��ȡ�ο�
    Dim supportRef As Reference
    If ax Is Nothing Then
        Set supportRef = workPt.CreateReferenceFromGeometry(workPt.OriginElements.PlaneXY)
    Else
        Dim AxPlnRefs As Variant
        AxPlnRefs = getAxisPlaneRefs(ax)
        Set supportRef = AxPlnRefs(0)
    End If
    Dim skt As Sketch
    Set skt = initSketch(minBody.Sketches, supportRef, ax)
    Call initBox2D(skt, maxBox)
    ' ����͹̨
    Call initPad(minBody, skt, maxBox)
    workPt.Update
    MsgBox "���"
End Sub
' ��Ҫ�ؼ����
Private Function getMimLength( _
    ByVal pt As Part, _
    ByVal body As AnyObject, _
    ByVal axRef As Reference, _
    vec As Variant) _
    As Double
    Dim bdyPt As Part
    Set bdyPt = KCL.GetParent_Of_T(body, "Part")
    Dim pln As HybridShapePlaneEquation
    Set pln = createPlane(pt, axRef, vec(0), vec(1), vec(2))
    Dim spa As AnyObject
    Set spa = pt.Parent.GetWorkbench("SPAWorkbench")
    getMimLength = _
        spa.GetMeasurable(bdyPt.CreateReferenceFromObject(body)) _
        .GetMinimumDistance(pt.CreateReferenceFromObject(pln))
End Function
' �ϲ���������
Private Function updateBox( _
    ByVal newBox As Variant, _
    ByVal maxBox As Variant) _
    As Variant
    If IsEmpty(maxBox) Then
        updateBox = newBox
        Exit Function
    End If
    If maxBox(MINMAX.MinX) > newBox(MINMAX.MinX) Then _
        maxBox(MINMAX.MinX) = newBox(MINMAX.MinX)
    If maxBox(MINMAX.MaxX) < newBox(MINMAX.MaxX) Then _
        maxBox(MINMAX.MaxX) = newBox(MINMAX.MaxX)
    If maxBox(MINMAX.MinY) > newBox(MINMAX.MinY) Then _
        maxBox(MINMAX.MinY) = newBox(MINMAX.MinY)
    If maxBox(MINMAX.MaxY) < newBox(MINMAX.MaxY) Then _
        maxBox(MINMAX.MaxY) = newBox(MINMAX.MaxY)
    If maxBox(MINMAX.MinZ) > newBox(MINMAX.MinZ) Then _
        maxBox(MINMAX.MinZ) = newBox(MINMAX.MinZ)
    If maxBox(MINMAX.Maxz) < newBox(MINMAX.Maxz) Then _
        maxBox(MINMAX.Maxz) = newBox(MINMAX.Maxz)
    updateBox = maxBox
End Function
' ��ȡ6�������������
Private Function getMaxSize_Bodies( _
    ByVal pt As Part, _
    ByVal bodies As collection, _
    ByVal ax As AxisSystem) _
    As Variant
    ' ���������������������ӦEnum MinMax
    Dim vec As Variant
    vec = Array( _
        Array(-1#, 0#, 0#), _
        Array(1#, 0#, 0#), _
        Array(0#, -1#, 0#), _
        Array(0#, 1#, 0#), _
        Array(0#, 0#, -1#), _
        Array(0#, 0#, 1#))
    Dim axRef As Reference
    If Not ax Is Nothing Then
        Set axRef = pt.CreateReferenceFromObject(ax)
    End If
    Dim tmpBox() As Double
    ReDim tmpBox(UBound(vec))
    Dim maxBox As Variant
    Dim i As Long
    Dim bdy As body
    For Each bdy In bodies
        For i = 0 To UBound(vec)
            tmpBox(i) = _
                (DMYLNG - getMimLength( _
                    pt, bdy, axRef, vec(i))) * IIf(i Mod 2 = 0, -1, 1)
        Next
        maxBox = updateBox(tmpBox, maxBox)
    Next
    getMaxSize_Bodies = maxBox
End Function
' ��ȡ����ϵ��û���򴴽�
Private Function getAxis( _
    ByVal doc As PartDocument) _
    As AxisSystem
    Dim pt As Part
    Set pt = doc.Part
    Dim axiss As AxisSystems
    Set axiss = pt.AxisSystems
    If axiss.Count > 0 Then
        Set getAxis = axiss.item(1)
    Else
        Set getAxis = initAxis(pt)
    End If
End Function
' ��������ϵ
Private Function initAxis( _
    ByVal pt As Part) _
    As AxisSystem
    Dim axiss As AxisSystems
    Set axiss = pt.AxisSystems
    Dim ax As Variant ' AxisSystem
    Set ax = axiss.Add()
    Dim ary As Variant
    ary = Array(0#, 0#, 0#)
    ax.OriginType = catAxisSystemOriginByCoordinates
    Set ax = ax
    ax.PutOrigin ary
    ary = Array(1#, 0#, 0#)
    ax.XAxisType = catAxisSystemAxisByCoordinates
    Set ax = ax
    ax.PutXAxis ary
    ary = Array(0#, 1#, 0#)
    ax.YAxisType = catAxisSystemAxisByCoordinates
    Set ax = ax
    ax.PutYAxis ary
    ax.IsCurrent = True
    pt.Update
    Set initAxis = ax
End Function
' �������
Private Function initPartDoc( _
    ByVal prod As Product) _
    As PartDocument
    Dim belongProd As Product
    If prod.Products.Count < 1 Then
        Set belongProd = prod.Parent.Parent
    Else
        Set belongProd = prod
    End If
    Dim prods As Products
    Set prods = belongProd.Products
    Dim newProd As Product
    Set newProd = prods.AddNewComponent("Part", "")
    Set initPartDoc = newProd.ReferenceProduct.Parent
End Function
' ��ȡ��ѡ��Ʒ����ʾ��ʵ��
Private Function getBodies( _
    ByVal prod As Product) _
    As collection
    Set getBodies = Nothing
    Dim sel As Selection
    Set sel = CATIA.ActiveDocument.Selection
    CATIA.HSOSynchronized = False
    sel.Clear
    sel.Add prod
    sel.Search "CATPrtSearch.BodyFeature.Visibility=Shown,sel"
    Dim lst As collection
    Set lst = New collection
    Dim i As Long
    Dim bdy As body
    For i = 1 To sel.Count2
        Set bdy = sel.Item2(i).Value
        If bdy.Shapes.Count > 0 Then
            lst.Add bdy
        End If
    Next
    sel.Clear
    CATIA.HSOSynchronized = True
    Dim msg As String
    If lst.Count < 1 Then
        msg = "û����ʾ��ʵ�壡"
        MsgBox msg, vbExclamation
        Exit Function
    End If
    Set getBodies = lst
End Function
' ����ƽ��
Private Function createPlane( _
    ByVal pt As Part, _
    ByVal axRef As Reference, _
    ByVal A As Double, _
    ByVal B As Double, _
    ByVal C As Double) _
    As HybridShapePlaneEquation
    Dim fact As HybridShapeFactory
    Set fact = pt.HybridShapeFactory
    Set createPlane = fact.AddNewPlaneEquation(A, B, C, DMYLNG)
    If Not axRef Is Nothing Then
        createPlane.RefAxisSystem = axRef
    End If
    Call pt.UpdateObject(createPlane)
End Function
' ��ȡ����ϵ��ƽ��Ĳο�
'Return : 0-XY,1-YZ,2-ZY �Ĳο�
Private Function getAxisPlaneRefs( _
    ByVal ax As AxisSystem) _
    As Variant ' Reference()
    Dim pt As Part
    Set pt = KCL.GetParent_Of_T(ax, "Part")
    Dim PlaneRef(2) As Reference
    Dim i As Long
    For i = 0 To UBound(PlaneRef)
        Set PlaneRef(i) = _
            pt.CreateReferenceFromBRepName( _
                getAxisPlaneBrepName(ax, i), ax)
    Next
    getAxisPlaneRefs = PlaneRef
End Function
' ��ȡ����ϵBrep����
' PlaneN0 : 0-XY,1-YZ,2-ZY �е�����һ��
Private Function getAxisPlaneBrepName( _
    ByVal ax As AxisSystem, _
    ByVal planeNo As Long) _
    As String
    Dim intName As String
    intName = ax.GetItem("ModelElement").InternalName
    getAxisPlaneBrepName = _
        "RSur:(Face:(Brp:(" + intName + ";" + CStr(planeNo + 1) + ");None:();Cf11:());" + _
        "WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
End Function
'***** ��ͼ��� *****
' ������ͼ
Private Function initSketch( _
    ByVal skts As Sketches, _
    ByVal supportRef As Reference, _
    ByVal ax As AxisSystem) _
    As Sketch
    Dim skt As Sketch
    Set skt = skts.Add(supportRef)
    Set initSketch = skt
    If ax Is Nothing Then Exit Function
    Dim axVar As Variant
    Set axVar = ax
    Dim ori(2) As Variant
    Call axVar.GetOrigin(ori)
    Dim vecX(2) As Variant, vecY(2) As Variant
    Call axVar.GetVectors(vecX, vecY)
    Dim settingAbsData As Variant
    settingAbsData = KCL.JoinAry(ori, vecX)
    settingAbsData = KCL.JoinAry(settingAbsData, vecY)
    Dim sktVar As Variant
    Set sktVar = skt
    Call sktVar.SetAbsoluteAxisData(settingAbsData)
End Function
' ��������
Private Sub initBox2D( _
    ByVal skt As Sketch, _
    ByVal poss As Variant)
    If Not UBound(poss) = 5 Then Exit Sub
    Dim fact2D As Factory2D
    Set fact2D = skt.OpenEdition()
    Dim pnt2D(3) As Point2D
    Set pnt2D(0) = fact2D.CreatePoint(poss(MINMAX.MinX), poss(MINMAX.MinY))
    Set pnt2D(1) = fact2D.CreatePoint(poss(MINMAX.MinX), poss(MINMAX.MaxY))
    Set pnt2D(2) = fact2D.CreatePoint(poss(MINMAX.MaxX), poss(MINMAX.MaxY))
    Set pnt2D(3) = fact2D.CreatePoint(poss(MINMAX.MaxX), poss(MINMAX.MinY))
    Dim consts As Constraints
    Set consts = skt.Constraints
    Call initLine2D(fact2D, consts, pnt2D(0), pnt2D(1))
    Call initLine2D(fact2D, consts, pnt2D(1), pnt2D(2))
    Call initLine2D(fact2D, consts, pnt2D(2), pnt2D(3))
    Call initLine2D(fact2D, consts, pnt2D(3), pnt2D(0))
    skt.CloseEdition
End Sub
' ������ - ��������Ӵ�ֱˮƽԼ��
Private Sub initLine2D( _
    ByVal fact2D As Factory2D, _
    ByVal csts As Constraints, _
    ByVal pntSt As Point2D, _
    ByVal pntEd As Point2D)
    Dim pntStVri As Variant
    Set pntStVri = pntSt
    Dim posSt(1) As Variant
    Call pntStVri.GetCoordinates(posSt)
    Dim pntEdVri As Variant
    Set pntEdVri = pntEd
    Dim posEd(1) As Variant
    Call pntEdVri.GetCoordinates(posEd)
    If dist2D_Ary2Ary(posSt, posEd) < 0.001 Then Exit Sub
    Dim line As Line2D
    Set line = fact2D.CreateLine(posSt(0), posSt(1), posEd(0), posEd(1))
    With line
        .StartPoint = pntSt
        .EndPoint = pntEd
    End With
    Dim ax2D As Axis2D
    Set ax2D = KCL.GetParent_Of_T(csts, "Sketch").GeometricElements.item(1)
    Select Case True
        Case Abs(posSt(0) - posEd(0)) < 0.001
            Call initConstraint( _
                csts, catCstTypeVerticality, _
                line, ax2D.VerticalReference) '��3,4�������ܵߵ�
            Call initConstraint( _
                csts, catCstTypeDistance, _
                ax2D.VerticalReference, line, posSt(0))
        Case Abs(posSt(1) - posEd(1)) < 0.001
            Call initConstraint( _
                csts, catCstTypeHorizontality, _
                line, ax2D.HorizontalReference) '��3,4�������ܵߵ�
            Call initConstraint( _
                csts, catCstTypeDistance, _
                ax2D.HorizontalReference, line, posSt(1))
    End Select
End Sub
' ���Լ��
Private Sub initConstraint( _
    ByVal csts As Constraints, _
    ByVal cstType As CatConstraintType, _
    ByVal itm1 As AnyObject, _
    ByVal itm2 As AnyObject, _
    Optional ByVal dist# = -1)
    Dim pt As Part
    Set pt = KCL.GetParent_Of_T(csts, "Part")
    Dim Cst As Constraint
    Set Cst = csts.AddBiEltCst( _
        cstType, _
        pt.CreateReferenceFromObject(itm1), _
        pt.CreateReferenceFromObject(itm2))
    Cst.Mode = catCstModeDrivingDimension
    If dist < 0.001 Then Exit Sub 'IsMissing(Dist)������
    Dim Leng As Length
    Set Leng = Cst.Dimension
    Leng.Value = dist
End Sub
'***** ʵ����� *****
' ����͹̨
Private Sub initPad( _
    ByVal bdy As body, _
    ByVal skt As Sketch, _
    ByVal poss As Variant)
    If Not UBound(poss) = 5 Then Exit Sub
    Dim pt As Part
    Set pt = KCL.GetParent_Of_T(bdy, "Part")
    Dim fact As ShapeFactory
    Set fact = pt.ShapeFactory
    Dim pad As pad
    Set pad = fact.AddNewPad(skt, poss(MINMAX.Maxz))
    pad.DirectionOrientation = catRegularOrientation
    Dim MinZ As Length
    Set MinZ = pad.SecondLimit.Dimension
    MinZ.Value = poss(MINMAX.MinZ) * -1
End Sub
' ������ɫ��
Private Sub changeColor( _
    ByVal itm As AnyObject)
    Dim doc As PartDocument
    Set doc = KCL.GetParent_Of_T(itm, "PartDocument")
    Dim sel As Selection
    Set sel = doc.Selection
    Dim vis As VisPropertySet
    Set vis = sel.VisProperties
    sel.Clear
    sel.Add itm
    Call vis.SetRealColor(128, 64, 64, 1)
    Call vis.SetRealOpacity(128, 1)
    Call vis.SetRealWidth(1, 1)
    Call vis.SetRealLineType(4, 1)
    sel.Clear
End Sub
'***** ������� *****
' ��ά����-��������
Private Function dist2D_Ary2Ary( _
    ByVal XY1 As Variant, _
    ByVal XY2 As Variant) _
    As Double
    dist2D_Ary2Ary = _
        Sqr((XY2(0) - XY1(0)) * (XY2(0) - XY1(0)) + _
            (XY2(1) - XY1(1)) * (XY2(1) - XY1(1)))
End Function

