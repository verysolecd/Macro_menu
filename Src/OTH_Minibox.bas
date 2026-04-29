Attribute VB_Name = "OTH_Minibox"
'{GP:6}
'{Ep:cRMinibox}
'{Caption:最小包络体}
'{ControlTipText:一键创建最小包络体}
'{BackColor:}

' %UI Label lbL_jpzcs  键盘造车手出品
' %UI txt txt_size  null
' %UI Button btnOK  创建最小包络体
' %UI Button btn_onlySize 只输出尺寸
' %UI Button btncancel  取消

Option Explicit
Private Const MINBODYNAME = "MinimumBox"
Private Const dynaExtrum = 1000000#
Private Const mdlname As String = "OTH_Minibox"
'types
Private Type mCoord
    X As Double
    Y As Double
    Z As Double
End Type
Private Type Box3D
    Min As mCoord
    Max As mCoord
    IsValid As Boolean
End Type

' Main Entry Point
Sub cRMinibox()
'If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
'    On Error GoTo ErrorHandler
    Dim workDoc As Document
    Dim prod As Product
    Dim msg As String: msg = "请选择要创建包络尺寸的产品"
    Dim osel: Set osel = CATIA.ActiveDocument.Selection
            If KCL.checkDocType("PartDocument") Then
                Set prod = CATIA.ActiveDocument.Product
            ElseIf KCL.checkDocType("ProductDocument") Then
                Set prod = KCL.SelectItem(msg, "Product")
                If prod Is Nothing Then Exit Sub
            Else
                Exit Sub
            End If
   Dim toCbox: toCbox = False
   Dim oEng: Set oEng = KCL.newEngine(mdlname)
    oEng.Show
    Select Case oEng.ClickedButton
        Case "btnOK": '要创建minibox则
            On Error Resume Next
                Set workDoc = prod.ReferenceProduct.Parent.part: Error.Clear
            On Error GoTo 0
            If workDoc Is Nothing Then Set workDoc = InitPartDoc(prod)
            toCbox = True
         Case "btn_onlySize":
            Set workDoc = KCL.get_aPartDoc
             toCbox = False
         Case Else: Exit Sub
    End Select
    
    Dim workPt As part: Set workPt = workDoc.part
    Dim targetBodies As Collection ' Gather Bodies
    Set targetBodies = GetBodies(prod)
    If targetBodies Is Nothing Then Exit Sub
    
    Dim ax As AxisSystem ' Get Axis System
    Set ax = GetAxis(workDoc)

    Dim totalBox As Box3D
    totalBox = GetMaxSize_Bodies(workPt, targetBodies, ax)
    If Not totalBox.IsValid Then
         MsgBox "无法计算包络体 (No valid geometry found).", vbExclamation
         Exit Sub
    End If
        

    
    If toCbox Then
        Dim minBody As body
        Dim bdy As AnyObject
        Set bdy = KCL.getItm(MINBODYNAME, workPt.bodies)
        If Not KCL.IsNothing(bdy) Then
            osel.Clear: osel.Add bdy
            osel.Delete: osel.Clear
        End If
        Set minBody = workPt.bodies.Add
        minBody.Name = MINBODYNAME
        Call ChangeColor(minBody)
        Dim supportRef As Reference
        If ax Is Nothing Then
            Set supportRef = workPt.CreateReferenceFromGeometry(workPt.OriginElements.PlaneXY)
        Else
            Dim AxPlnRefs As Variant
            AxPlnRefs = GetAxisPlaneRefs(ax)
            Set supportRef = AxPlnRefs(0) ' XY Plane of Axis
        End If
        Dim skt As Sketch: Set skt = InitSketch(minBody.Sketches, supportRef, ax)
        Call InitBox2D(skt, totalBox)
        Call InitPad(minBody, skt, totalBox)
        workPt.Update
        workPt.InWorkObject = minBody
        MsgBox "Done", vbInformation
    End If
        Dim iSize As mCoord
        With totalBox
            iSize.X = .Max.X - .Min.X
            iSize.Y = .Max.Y - .Min.Y
            iSize.Z = .Max.Z - .Min.Z
        End With
        Call oEng.Alert(Format(iSize.X, "0.00") & " x " _
                 & Format(iSize.Y, "0.00") & " x " _
                 & Format(iSize.Z, "0.00"))
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in OTH_Minibox: " & Err.Description, vbCritical
End Sub

' -------------------------------------------------------------------------
' Core Logic
' -------------------------------------------------------------------------

Private Function GetMaxSize_Bodies( _
    ByVal pt As part, _
    ByVal bodies As Collection, _
    ByVal ax As AxisSystem) _
    As Box3D
    
    ' Define 6 Analysis Directions (Normal Vectors)
    ' -X, +X, -Y, +Y, -Z, +Z
    Dim Dirs(5) As mCoord
    Dirs(0) = CreatePoint3D(-1, 0, 0)
    Dirs(1) = CreatePoint3D(1, 0, 0)
    Dirs(2) = CreatePoint3D(0, -1, 0)
    Dirs(3) = CreatePoint3D(0, 1, 0)
    Dirs(4) = CreatePoint3D(0, 0, -1)
    Dirs(5) = CreatePoint3D(0, 0, 1)
                
    Dim axRef As Reference
    If Not ax Is Nothing Then
        Set axRef = pt.CreateReferenceFromObject(ax)
    End If
    
    Dim globalBox As Box3D
    ' Initialize as Invalid
    globalBox.IsValid = False
    
    Dim currentBodyBox As Box3D
    Dim bdy As body
    Dim i As Long
    Dim dist As Double
    Dim limits(5) As Double ' Store measured limits for 6 directions
    
    For Each bdy In bodies
        For i = 0 To 5
            dist = GetMimLength(pt, bdy, axRef, Dirs(i))
           ' -X, +X, -Y, +Y, -Z, +Z ，'此处0，2，4为负向;1,3,5 为正向
            limits(i) = (dynaExtrum - dist) * IIf(i Mod 2 = 0, -1, 1)
        Next i
        
        ' Construct Box3D for this body
        currentBodyBox.Min.X = limits(0)
        currentBodyBox.Max.X = limits(1)
        currentBodyBox.Min.Y = limits(2)
        currentBodyBox.Max.Y = limits(3)
        currentBodyBox.Min.Z = limits(4)
        currentBodyBox.Max.Z = limits(5)
        currentBodyBox.IsValid = True
        
        ' Merge into Global Box
        globalBox = UpdateBox(globalBox, currentBodyBox)
    Next
    
    GetMaxSize_Bodies = globalBox
End Function

Private Function UpdateBox(CurrentBox As Box3D, NewBox As Box3D) As Box3D
    If Not CurrentBox.IsValid Then
        UpdateBox = NewBox
        Exit Function
    End If
    
    Dim Result As Box3D
    Result = CurrentBox
    
    ' Simplify comparison using Min/Max helpers
    Result.Min.X = Min(Result.Min.X, NewBox.Min.X)
    Result.Max.X = Max(Result.Max.X, NewBox.Max.X)
    
    Result.Min.Y = Min(Result.Min.Y, NewBox.Min.Y)
    Result.Max.Y = Max(Result.Max.Y, NewBox.Max.Y)
    
    Result.Min.Z = Min(Result.Min.Z, NewBox.Min.Z)
    Result.Max.Z = Max(Result.Max.Z, NewBox.Max.Z)
    
    Result.IsValid = True
    UpdateBox = Result
End Function

Private Function GetMimLength( _
    ByVal pt As part, _
    ByVal body As AnyObject, _
    ByVal axRef As Reference, _
    Direction As mCoord) _
    As Double
    
    Dim bdyPt As part
    Set bdyPt = KCL.GetParent_Of_T(body, "Part")
    
    Dim pln As HybridShapePlaneEquation
    Set pln = CreatePlane(pt, axRef, Direction.X, Direction.Y, Direction.Z)
    
    Dim spa As AnyObject
    Set spa = pt.Parent.GetWorkbench("SPAWorkbench")
    
    GetMimLength = spa.GetMeasurable(bdyPt.CreateReferenceFromObject(body)) _
                      .GetMinimumDistance(pt.CreateReferenceFromObject(pln))
End Function

' -------------------------------------------------------------------------
' Geometry Creation
' -------------------------------------------------------------------------

Private Sub InitBox2D(ByVal skt As Sketch, box As Box3D)
    Dim fact2D As Factory2D
    Set fact2D = skt.OpenEdition()
    Dim pnt2D(3) As Point2D
    ' Use Box properties directly - Clean and Readable
    Set pnt2D(0) = fact2D.CreatePoint(box.Min.X, box.Min.Y)
    Set pnt2D(1) = fact2D.CreatePoint(box.Min.X, box.Max.Y)
    Set pnt2D(2) = fact2D.CreatePoint(box.Max.X, box.Max.Y)
    Set pnt2D(3) = fact2D.CreatePoint(box.Max.X, box.Min.Y)
    Dim consts As Constraints
    Set consts = skt.Constraints
    Call InitLine2D(fact2D, consts, pnt2D(0), pnt2D(1))
    Call InitLine2D(fact2D, consts, pnt2D(1), pnt2D(2))
    Call InitLine2D(fact2D, consts, pnt2D(2), pnt2D(3))
    Call InitLine2D(fact2D, consts, pnt2D(3), pnt2D(0))
    skt.CloseEdition
End Sub

Private Sub InitPad(ByVal bdy As body, ByVal skt As Sketch, box As Box3D)
    Dim pt As part
    Set pt = KCL.GetParent_Of_T(bdy, "Part")
    Dim Fact As ShapeFactory
    Set Fact = pt.ShapeFactory
    
    Dim pad As pad
    Set pad = Fact.AddNewPad(skt, box.Max.Z)
    pad.DirectionOrientation = catRegularOrientation
    
    Dim MinZ As Length
    Set MinZ = pad.SecondLimit.Dimension
    ' Pad SecondLimit is opposite direction, so we typically use negative value or positive depending on orientation.
    ' Original code: MinZ.Value = poss(MINMAX.MinZ) * -1
    ' If MinZ is -10, we want Pad to go down 10.
    MinZ.Value = box.Min.Z * -1
End Sub

Private Sub InitLine2D( _
    ByVal fact2D As Factory2D, _
    ByVal csts As Constraints, _
    ByVal pntSt As Point2D, _
    ByVal pntEd As Point2D)
    
    Dim pntStVri As Variant: Set pntStVri = pntSt
    Dim posSt(1) As Variant
    Call pntStVri.GetCoordinates(posSt)
    
    Dim pntEdVri As Variant: Set pntEdVri = pntEd
    Dim posEd(1) As Variant
    Call pntEdVri.GetCoordinates(posEd)
    
    If Dist2D_Ary2Ary(posSt, posEd) < 0.001 Then Exit Sub
    
    Dim line As Line2D
    Set line = fact2D.CreateLine(posSt(0), posSt(1), posEd(0), posEd(1))
    line.StartPoint = pntSt
    line.EndPoint = pntEd
    
    Dim ax2D As Axis2D
    Set ax2D = KCL.GetParent_Of_T(csts, "Sketch").GeometricElements.item(1)
   ' Automatic Constraint Creation
    Select Case True
        Case Abs(posSt(0) - posEd(0)) < 0.001 ' Vertical
            Call InitConstraint(csts, catCstTypeVerticality, line, ax2D.VerticalReference)
            Call InitConstraint(csts, catCstTypeDistance, ax2D.VerticalReference, line, posSt(0))
            
        Case Abs(posSt(1) - posEd(1)) < 0.001 ' Horizontal
            Call InitConstraint(csts, catCstTypeHorizontality, line, ax2D.HorizontalReference)
            Call InitConstraint(csts, catCstTypeDistance, ax2D.HorizontalReference, line, posSt(1))
    End Select
End Sub


' -------------------------------------------------------------------------
' Helpers & Utilities
' -------------------------------------------------------------------------

Private Function CreatePoint3D(ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As mCoord
    CreatePoint3D.X = X
    CreatePoint3D.Y = Y
    CreatePoint3D.Z = Z
End Function

Private Function Min(ByVal A As Double, ByVal B As Double) As Double
    If A < B Then Min = A Else Min = B
End Function

Private Function Max(ByVal A As Double, ByVal B As Double) As Double
    If A > B Then Max = A Else Max = B
End Function

Private Function GetAxis(ByVal doc As PartDocument) As AxisSystem
    Dim pt As part: Set pt = doc.part
    Dim axiss As AxisSystems: Set axiss = pt.AxisSystems
    If axiss.count > 0 Then
        Set GetAxis = axiss.item(1)
    Else
        Set GetAxis = InitAxis(pt)
    End If
End Function

Private Function InitAxis(ByVal pt As part) As AxisSystem
    Dim axiss As AxisSystems: Set axiss = pt.AxisSystems
    Dim ax As Variant: Set ax = axiss.Add()
    
    ax.OriginType = catAxisSystemOriginByCoordinates
    ax.PutOrigin Array(0#, 0#, 0#)
    
    ax.XAxisType = catAxisSystemAxisByCoordinates
    ax.PutXAxis Array(1#, 0#, 0#)
    
    ax.YAxisType = catAxisSystemAxisByCoordinates
    ax.PutYAxis Array(0#, 1#, 0#)
    
    ax.IsCurrent = True
    pt.Update
    Set InitAxis = ax
End Function

Private Function InitPartDoc(ByVal prod As Product) As PartDocument
    Dim belongProd As Product
    If prod.Products.count < 1 Then
        Set belongProd = prod.Parent.Parent
    Else
        Set belongProd = prod
    End If
    Dim newProd As Product
    Set newProd = belongProd.Products.AddNewComponent("Part", "")
    newProd.partNumber = "Mini_box_" & prod.partNumber
    Set InitPartDoc = newProd.ReferenceProduct.Parent
End Function

Private Function GetBodies(ByVal prod As Product) As Collection
    Set GetBodies = Nothing
    Dim sel As Selection
    Set sel = CATIA.ActiveDocument.Selection
    CATIA.HSOSynchronized = False
    sel.Clear
    sel.Add prod
    sel.Search "CATPrtSearch.BodyFeature.Visibility=Shown,sel"
    
    Dim lst As New Collection
    Dim i As Long
    Dim bdy As body
    For i = 1 To sel.Count2
        Set bdy = sel.Item2(i).Value
        If bdy.Shapes.count > 0 And bdy.Name <> MINBODYNAME Then
            lst.Add bdy
        End If
    Next
    sel.Clear
    CATIA.HSOSynchronized = True
    
    If lst.count < 1 Then
        MsgBox "No visible bodies found", vbExclamation
        Exit Function
    End If
    Set GetBodies = lst
End Function

Private Function CreatePlane( _
    ByVal pt As part, _
    ByVal axRef As Reference, _
    ByVal A As Double, _
    ByVal B As Double, _
    ByVal c As Double) _
    As HybridShapePlaneEquation
    
    Dim Fact As HybridShapeFactory
    Set Fact = pt.HybridShapeFactory
    Set CreatePlane = Fact.AddNewPlaneEquation(A, B, c, dynaExtrum)
    If Not axRef Is Nothing Then
        CreatePlane.RefAxisSystem = axRef
    End If
    pt.UpdateObject CreatePlane
End Function

Private Function GetAxisPlaneRefs(ByVal ax As AxisSystem) As Variant
    Dim pt As part
    Set pt = KCL.GetParent_Of_T(ax, "Part")
    Dim PlaneRef(2) As Reference
    Dim i As Long
    For i = 0 To UBound(PlaneRef)
        Set PlaneRef(i) = pt.CreateReferenceFromBRepName(GetAxisPlaneBrepName(ax, i), ax)
    Next
    GetAxisPlaneRefs = PlaneRef
End Function

Private Function GetAxisPlaneBrepName(ByVal ax As AxisSystem, ByVal planeNo As Long) As String
    Dim intName As String
    intName = ax.GetItem("ModelElement").internalName
    GetAxisPlaneBrepName = _
        "RSur:(Face:(Brp:(" + intName + ";" + CStr(planeNo + 1) + ");None:();Cf11:());" + _
        "WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
End Function

Private Function InitSketch( _
    ByVal skts As Sketches, _
    ByVal supportRef As Reference, _
    ByVal ax As AxisSystem) _
    As Sketch
    
    Dim skt As Sketch
    Set skt = skts.Add(supportRef)
    Set InitSketch = skt
    If ax Is Nothing Then Exit Function
    
    Dim axVar As Variant: Set axVar = ax
    Dim ori(2) As Variant
    Call axVar.GetOrigin(ori)
    
    Dim vecX(2) As Variant, vecY(2) As Variant
    Call axVar.GetVectors(vecX, vecY)
    
    Dim settingAbsData As Variant
    settingAbsData = KCL.JoinAry(ori, vecX)
    settingAbsData = KCL.JoinAry(settingAbsData, vecY)
    
    Dim sktVar As Variant: Set sktVar = skt
    Call sktVar.SetAbsoluteAxisData(settingAbsData)
End Function

Private Sub InitConstraint( _
    ByVal csts As Constraints, _
    ByVal cstType As CatConstraintType, _
    ByVal itm1 As AnyObject, _
    ByVal itm2 As AnyObject, _
    Optional ByVal dist As Double = -1)
    
    Dim pt As part
    Set pt = KCL.GetParent_Of_T(csts, "Part")
    Dim Cst As Constraint
    Set Cst = csts.AddBiEltCst( _
        cstType, _
        pt.CreateReferenceFromObject(itm1), _
        pt.CreateReferenceFromObject(itm2))
    Cst.Mode = catCstModeDrivingDimension
    If dist >= 0 Then
        Dim Leng As Length
        Set Leng = Cst.Dimension
        Leng.Value = dist
    End If
End Sub

Private Sub ChangeColor(ByVal itm As AnyObject)
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
    Call vis.SetRealWidth(1, 1) ' Fixed: Width is usually integer 1-5
    Call vis.SetRealLineType(4, 1)
    sel.Clear
End Sub

Private Function Dist2D_Ary2Ary(ByVal XY1 As Variant, ByVal XY2 As Variant) As Double
    Dist2D_Ary2Ary = Sqr((XY2(0) - XY1(0)) ^ 2 + (XY2(1) - XY1(1)) ^ 2)
End Function











