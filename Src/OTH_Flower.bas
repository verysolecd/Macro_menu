Attribute VB_Name = "OTH_Flower"
'Attribute VB_Name = "OTH_flower"
'{GP:4446}
'{Ep:DF}
'{Caption:天杀的小花}
'{ControlTipText: 可以画一些花出来}
'{BackColor:16744703}

Private X0 As Double
Private Y0 As Double
Private Z0 As Double
Sub DF()
 If Not CanExecute("PartDocument") Then Exit Sub
    FrmFlower.Show
End Sub
Sub iPos(Optional ByVal pp As Double)

    Z0 = 0
    Dim idst
    idstx = Val(FrmFlower.TxtDeltaX.Text) * 2.3
   idsty = Val(FrmFlower.TxtDeltaX.Text) * 2.6
    
     If pp > 1 Then '循环画图
    
       For i = 1 To pp
          For j = 1 To pp
                X0 = idstx * i
               Y0 = idsty * j
               
               fs = i * j
               If fs Mod 2 = 0 Then
                Z = 50
                Else
                Z = -50
                End If
               Call Flower
           Next
       Next
    Else
       
        X0 = 100
        X0 = Val(FrmFlower.TxtX0.Text)
        Y0 = 100
        Y0 = Val(FrmFlower.TxtY0.Text)
        Z0 = 0
        
        Call Flower
     End If
End Sub
    
Sub Flower(Optional ByVal pp As Double)
'Created by Alireza Reihani

'Step 1: Create Geometrical Sets
'Step 2: Create Sketch in YZ Plane:
    '2-1: Create two  Arcs with Radius
        '2-1-1: Create Sketch Object
        '2-1-2: Create Arc Radius Dimension
        '2-1-3: Create Arc Center H,V Position
'Step 3: Sweep Circle
    'Step 3-1: Join Sketch
    'Step 3-2: Sweep Circle
    'Step 3-3: Sweep Color
'Step 4: Create Ovary
    'Step 4-1: Create Extremum Point of Sketch
    'Step 4-2: Create Plane perp to Sketch
    'Step 4-3: Create Parallel Plane
    'Step 4-4: Create Intersection (Instead of Boundary of Sweep)
    'Step 4-5: Create Circle Sketch
    'Step 4-6: Create Blend
    'Step 4-7: Create Sketch2 Projection
    'Step 4-8: Create Fill
    'Step 4-9: Create Bump
    'Step 4-10: Create Hole
        'Step 4-10-1: Create Hole Center Point
        'Step 4-10-2: Project Hole Circle on the Bump
        'Step 4-10-3: Split
    'Step 4-11: Create Fillet
    'Step 4-12: Fillet Color
'Step 5: Flower
    'Step 5-1: Create Extremum Point of Intersection
    'Step 5-2: Create Plane Perp. to Sketch2 through Extremum Point
    'Step 5-3: Create End Point of Flower
    'Step 5-4: Create Flower Line
    'Step 5-5: Create Plane Perp. to Flower Line through Start Point
    'Step 5-6: Create Plane Perp. to Flower Line through End Point
    Dim Pi As Double
    Pi = 3.14159265358979
    
    Dim FlowerDirection As Boolean
    FlowerDirection = True
    
    Dim R1 As Double
    Dim R12 As Double
    Dim H1 As Double
    Dim H2 As Double
    Dim H As Double
    Dim SweepR As Double
    Dim OffsetPlane As Double
    Dim R2 As Double
    Dim RHole As Double
    Dim DistHole As Double
    Dim QtyHole As Integer
    Dim RFillet As Double
    Dim DeltaX As Double
    Dim DeltaY As Double
    Dim RS1 As Double
    Dim RS2 As Double
    Dim RS3 As Double
    Dim RS4 As Double
    Dim RS5 As Double
    Dim RS6 As Double
    Dim RS7 As Double
    Dim XS1 As Double
    Dim YS1 As Double
    Dim XS2 As Double
    Dim YS2 As Double
    Dim XS3 As Double
    Dim YS3 As Double
    Dim XS4 As Double
    Dim YS4 As Double
    Dim XS5 As Double
    Dim YS5 As Double
    Dim XS6 As Double
    Dim YS6 As Double
    Dim XS7 As Double
    Dim YS7 As Double
    Dim Alfa1Start As Double
    Dim Alfa1End As Double
    Dim Alfa2Start As Double
    Dim Alfa2End As Double
    Dim Alfa3Start As Double
    Dim Alfa3End As Double
    Dim Alfa4Start As Double
    Dim Alfa4End As Double
    Dim Alfa5Start As Double
    Dim Alfa5End As Double
    Dim Alfa6Start As Double
    Dim Alfa6End As Double
    Dim Alfa7Start As Double
    Dim Alfa7End As Double
    
    Dim QtyPetal As Integer
    QtyPetal = 7
    QtyPetal = Val(FrmFlower.TxtQtyPetal.Text)
    Dim AlfaPetal2 As Double
    AlfaPetal2 = 25
    AlfaPetal2 = Val(FrmFlower.TxtAlfaPetal2.Text)
    Dim ScaleFactor As Double
    ScaleFactor = 0.75
    ScaleFactor = Val(FrmFlower.TxtScaleFactor.Text)

    
    Dim R_Stem As Byte
    Dim G_Stem As Byte
    Dim B_Stem As Byte
    R_Stem = 0
    G_Stem = 255
    B_Stem = 0
    R_Stem = FrmFlower.ScrR1.value
    G_Stem = FrmFlower.ScrG1.value
    B_Stem = FrmFlower.ScrB1.value
    
    Dim R_Ovary As Byte
    Dim G_Ovary As Byte
    Dim B_Ovary As Byte
    R_Ovary = 255
    G_Ovary = 255
    B_Ovary = 0
    R_Ovary = FrmFlower.ScrR2.value
    G_Ovary = FrmFlower.ScrG2.value
    B_Ovary = FrmFlower.ScrB2.value
    
    Dim R_Petal1 As Byte
    Dim G_Petal1 As Byte
    Dim B_Petal1 As Byte
    R_Petal1 = 105
    G_Petal1 = 0
    B_Petal1 = 105
    R_Petal1 = FrmFlower.ScrR3.value
    G_Petal1 = FrmFlower.ScrG3.value
    B_Petal1 = FrmFlower.ScrB3.value
    
    Dim R_Petal2 As Byte
    Dim G_Petal2 As Byte
    Dim B_Petal2 As Byte
    R_Petal2 = 255
    G_Petal2 = 0
    B_Petal2 = 255
    R_Petal2 = FrmFlower.ScrR4.value
    G_Petal2 = FrmFlower.ScrG4.value
    B_Petal2 = FrmFlower.ScrB4.value
    
    R1 = 150    'Arc1 in Sketch.1
    R1 = Val(FrmFlower.TxtR1.Text)

    R12 = 150   'Arc2 in Sketch.1
    R12 = R1
    
    H1 = 100
    H2 = 100
    H = Val(FrmFlower.TxtH.Text)
    H1 = H / 2
    H2 = H1
    
    SweepR = 10
    SweepR = Val(FrmFlower.TxtSweepR.Text) / 2
    
    OffsetPlane = 20
    OffsetPlane = Val(FrmFlower.TxtOffsetPlane.Text)

    R2 = 20     'Full Circle in Sketch.2
    R2 = Val(FrmFlower.TxtR2.Text)

    RHole = 2
    RHole = Val(FrmFlower.TxtRHole.Text)

    DistHole = 12
    DistHole = Val(FrmFlower.TxtDistHole.Text) / 2

    QtyHole = 10
    QtyHole = Val(FrmFlower.TxtQtyHole.Text)

    RFillet = 1
    DeltaX = -143
    DeltaX = -Val(FrmFlower.TxtDeltaX.Text)

    DeltaY = 30
    DeltaY = Val(FrmFlower.TxtDeltaY.Text)

    
    RS1 = 0.6
    RS2 = 0.5
    RS3 = 27.11
    RS4 = 32.88
    RS5 = 29.86
    RS6 = 21.36
    RS7 = 4
    XS1 = -0.1
    YS1 = -0.1
    XS2 = -0.02
    YS2 = -1.42
    XS3 = -1.34
    YS3 = 19.07
    XS4 = -2.36
    YS4 = 14.58
    XS5 = -2.42
    YS5 = 6.89
    XS6 = -0.24
    YS6 = 5.92
    XS7 = -0.17
    YS7 = 1.99
    Alfa1Start = 150.14
    Alfa1End = 379.26
    Alfa2Start = 176.83
    Alfa2End = 363.18
    Alfa3Start = 223.57
    Alfa3End = 312.88
    Alfa4Start = 208.38
    Alfa4End = 328.86
    Alfa5Start = 195.28
    Alfa5End = 342.77
    Alfa6Start = 198.45
    Alfa6End = 336.21
    Alfa7Start = 184.39
    Alfa7End = 357.19
    
    Dim myPart As part
    Set myPart = CATIA.ActiveDocument.part
    Dim HB1s As HybridBodies
    Set HB1s = myPart.HybridBodies
    Dim mySF As ShapeFactory
    Set mySF = myPart.ShapeFactory
' 1: Create Geometrical Sets---------------------
    Dim myHB As HybridBody
    Set myHB = HB1s.Add
    myHB.Name = "GS_Flower_Wireframe"
    
    Dim myHB2 As HybridBody
    Set myHB2 = HB1s.Add
    myHB2.Name = "GS_Flower_Hidden-Surface"
    
    Dim myHB3 As HybridBody
    Set myHB3 = HB1s.Add
    myHB3.Name = "GS_Flower_Surface"
    
    Dim HSF As HybridShapeFactory
    Set HSF = myPart.HybridShapeFactory
    ''''''''''''''''''''''''''''''''''''''''''''''''''
'Step 2: Create Sketch
    
    ' Create Sketches List Object in Hybrid Body -------------
    
    Dim MySketches As Sketches
    Set MySketches = myHB.HybridSketches
    
    ' Create Reference Plane ------------------
    
    Dim OriginElement, myPlane
    Set OriginElement = myPart.OriginElements
    Set myPlane = OriginElement.PlaneYZ
    
    Dim myPlaneX As HybridShapePlaneOffset
    Set myPlaneX = HSF.AddNewPlaneOffset(myPlane, X0, False)
    myHB.AppendHybridShape myPlaneX
    
    Dim RefmyPlaneX As Reference
    Set RefmyPlaneX = myPart.CreateReferenceFromObject(myPlaneX)
    
    Dim myPlaneXobj
    Set myPlaneXobj = RefmyPlaneX
    
    myPart.Update
    ' Create Sketch Object --------------------
    
    Dim mySketch As Sketch
    Set mySketch = MySketches.Add(myPlaneXobj)
    
     ' Create 2D Tool Box and Open Sketch ---------------------------
    
    Dim F2D As Factory2D
    Set F2D = mySketch.OpenEdition
        
    ' Create Geometry ----------------------------------------------
   
    'Arc1
    
    Dim alfaRad As Double
    alfaRad = Atn((H1 / 2) / Sqr(R1 ^ 2 - (H1 / 2) ^ 2))
    Dim Xcen1 As Double
    If FlowerDirection = True Then
        Xcen1 = Y0 - Sqr(R1 ^ 2 - (H1 / 2) ^ 2)
    Else
        Xcen1 = Y0 + Sqr(R1 ^ 2 - (H1 / 2) ^ 2)
    End If
    Dim Ycen1 As Double
    Ycen1 = Z0 + H1 / 2
    
    Dim Circle1 As Circle2D
    If FlowerDirection = True Then
        Set Circle1 = F2D.CreateCircle(Xcen1, Ycen1, R1, 2 * Pi - alfaRad, 2 * Pi + alfaRad) '0<=start<2*pi, start<end<=4*pi
    Else
        Set Circle1 = F2D.CreateCircle(Xcen1, Ycen1, R1, Pi - alfaRad, Pi + alfaRad) '0<=start<2*pi, start<end<=4*pi
    End If
    
    Dim PtCtr As Point2D
    Set PtCtr = F2D.CreatePoint(Xcen1, Ycen1)
    Circle1.CenterPoint = PtCtr
    
    ' Create R Constraint
    Dim MyConstraints As Constraints
    Dim RadiusCst As Constraint
    Dim ArcRef As Reference
    Set ArcRef = myPart.CreateReferenceFromObject(Circle1)
    Set MyConstraints = mySketch.Constraints
    Set RadiusCst = MyConstraints.AddMonoEltCst(catCstTypeRadius, ArcRef)
    RadiusCst.Dimension.value = R1
    
    ' Create H,V constraint of Arc1 Center Point
    Set Axis2D = mySketch.GeometricElements.item("AbsoluteAxis")
    Set Hdir = Axis2D.getItem("HDirection")
    Set Vdir = Axis2D.getItem("VDirection")
    
    
    Dim RefCenter1 As Reference
    Set RefCenter1 = myPart.CreateReferenceFromObject(PtCtr)
    Dim RefH As Reference
    Set RefH = myPart.CreateReferenceFromObject(Hdir)
    Dim RefV As Reference
    Set RefV = myPart.CreateReferenceFromObject(Vdir)
    
    Dim HCst1 As Constraint
    Dim VCst1 As Constraint
    Set HCst1 = MyConstraints.AddBiEltCst(catCstTypeDistance, RefH, RefCenter1)
    HCst1.Dimension.value = Ycen1
    
    Set VCst1 = MyConstraints.AddBiEltCst(catCstTypeDistance, RefV, RefCenter1)
    VCst1.Dimension.value = Xcen1

    'Arc2
    Dim alfaRad2 As Double
    alfaRad2 = Atn((H2 / 2) / Sqr(R12 ^ 2 - (H2 / 2) ^ 2))
    Dim Xcen2 As Double
    If FlowerDirection = True Then
        Xcen2 = Y0 + Sqr(R1 ^ 2 - (H2 / 2) ^ 2)
    Else
        Xcen2 = Y0 - Sqr(R1 ^ 2 - (H2 / 2) ^ 2)
    End If
    
    Dim Ycen2 As Double
    Ycen2 = Z0 + H1 + H2 / 2
    
    Dim Circle2 As Circle2D
    If FlowerDirection = True Then
        Set Circle2 = F2D.CreateCircle(Xcen2, Ycen2, R12, Pi - alfaRad2, Pi + alfaRad2)   '0<=start<2*pi, start<end<=4*pi
    Else
        Set Circle2 = F2D.CreateCircle(Xcen2, Ycen2, R12, 2 * Pi - alfaRad2, 2 * Pi + alfaRad2)   '0<=start<2*pi, start<end<=4*pi
    End If
    Dim PtCtr2 As Point2D
    Set PtCtr2 = F2D.CreatePoint(Xcen2, Ycen2)
    Circle2.CenterPoint = PtCtr2

   ' Create R Constraint
    Dim RadiusCst2 As Constraint
    Dim ArcRef2 As Reference
    Set ArcRef2 = myPart.CreateReferenceFromObject(Circle2)
    Set RadiusCst2 = MyConstraints.AddMonoEltCst(catCstTypeRadius, ArcRef2)
    RadiusCst2.Dimension.value = R12
    
    ' Create H,V constraint of Arc2 Center Point
    Dim RefCenter2 As Reference
    Set RefCenter2 = myPart.CreateReferenceFromObject(PtCtr2)
    Dim HCst2 As Constraint
    Dim VCst2 As Constraint
    Set HCst2 = MyConstraints.AddBiEltCst(catCstTypeDistance, RefH, RefCenter2)
    HCst2.Dimension.value = Ycen2
    
    Set VCst2 = MyConstraints.AddBiEltCst(catCstTypeDistance, RefV, RefCenter2)
    VCst2.Dimension.value = Xcen2
    ' Close Sketch and Update MyComponent ----------------------------
    
    mySketch.CloseEdition
    
    
    Set Refskt = myPart.CreateReferenceFromObject(mySketch)
           HSF.GSMVisibility Refskt, 0
    
    myPart.Update
    
    
'Step 3: Sweep Circle
    'Step 3-1: Join Sketch '''''''''''
        'Dim myCurve
        'Set myCurve = myHB.HybridSketches.Item("Sketch.4")
        Dim Ref1 As Reference
        Set Ref1 = myPart.CreateReferenceFromObject(mySketch)
        Dim myJoin As HybridShapeAssemble
        Set myJoin = HSF.AddNewJoin(Ref1, Ref1)
        myHB.AppendHybridShape myJoin
    
    
    'Step 3 - 2: Sweep Circle
        Dim RefCurve As Reference
        Set RefCurve = myPart.CreateReferenceFromObject(myJoin)
        
        
        
        Dim mySweepCircle As HybridShapeSweepCircle
        Set mySweepCircle = HSF.AddNewSweepCircle(RefCurve)
        mySweepCircle.SetRadius 1, SweepR    'SweepR: Radius
        mySweepCircle.Mode = 6          '6: Center and Radius
        ' Assign Arc to Geometrical Set --------------------------------
        
        myHB3.AppendHybridShape mySweepCircle
        myPart.Update
   
    'Step 3 - 3: Sweep Color
        Dim MyList As Selection
        Set MyList = CATIA.ActiveDocument.Selection
        MyList.Clear
        'Dim myObj As Object
        'Set myObj = myRef
        MyList.Add mySweepCircle
        MyList.VisProperties.SetRealColor R_Stem, G_Stem, B_Stem, 1
    
'Step 4: Create Ovary
    'Step 4-1: Create Extremum Point of Sketch
        Dim RefExt As Reference
        Set RefExt = myPart.CreateReferenceFromObject(myJoin)
        Dim mydir1 As HybridShapeDirection
        Set mydir1 = HSF.AddNewDirectionByCoord(0#, 0#, 1#)  'Z direction
        
        Dim myExtremumPoint As HybridShapeExtremum
        Set myExtremumPoint = HSF.AddNewExtremum(RefExt, mydir1, 1)  '1:Max
        myHB.AppendHybridShape myExtremumPoint
        
    'Step 4-2: Create Plane perp to Sketch
        
        Dim RefCurveNormal1 As Reference
        Set RefCurveNormal1 = myPart.CreateReferenceFromObject(myJoin)
        Dim RefPtNormal1 As Reference
        Set RefPtNormal1 = myPart.CreateReferenceFromObject(myExtremumPoint)
        
        Dim myNormalPlane As HybridShapePlaneNormal
        Set myNormalPlane = HSF.AddNewPlaneNormal(RefCurveNormal1, RefPtNormal1)
        myHB.AppendHybridShape myNormalPlane
    
    
    'Step 4-3: Create Parallel Plane
        Dim RefPlane1 As Reference
        Set RefPlane1 = myPart.CreateReferenceFromObject(myNormalPlane)
        Dim myOffsetPlane As HybridShapePlaneOffset
        Set myOffsetPlane = HSF.AddNewPlaneOffset(RefPlane1, OffsetPlane, False)
        myHB.AppendHybridShape myOffsetPlane
    
        
    'Step 4-4: Create Intersection (Instead of Boundary of Sweep)
        
        Dim RefSweep1 As Reference
        Set RefSweep1 = myPart.CreateReferenceFromObject(mySweepCircle)
        Dim myHybridShapeIntersection As HybridShapeIntersection
        Set myHybridShapeIntersection = HSF.AddNewIntersection(RefPlane1, RefSweep1)
        myHB.AppendHybridShape myHybridShapeIntersection
        
    myPart.Update
        
    'Step 4-5: Create Circle Sketch
        
        Dim RefPlanemySketch2 As Reference
        Set RefPlanemySketch2 = myPart.CreateReferenceFromObject(myOffsetPlane)
        Dim mySketch2 As Sketch
        Set mySketch2 = MySketches.Add(RefPlanemySketch2)
        Set F2D = mySketch2.OpenEdition
        Dim Circlefull1 As Circle2D
        Set Circlefull1 = F2D.CreateClosedCircle(0, 0, R2)
        
         mySketch2.CloseEdition
          Set Refskt = myPart.CreateReferenceFromObject(mySketch2)
           HSF.GSMVisibility Refskt, 0
         
         
         myPart.Update
     'Step 4-6: Create Blend
     
        'Step 4-6-1: Create Extremum Point of Intersection
        
        Dim mydirX As HybridShapeDirection
        Set mydirX = HSF.AddNewDirectionByCoord(1#, 0#, 0#)  'X direction
        
        Dim RefExtIntersection As Reference
        Set RefExtIntersection = myPart.CreateReferenceFromObject(myHybridShapeIntersection)
        Dim myExtremumPointIntersection As HybridShapeExtremum
        Set myExtremumPointIntersection = HSF.AddNewExtremum(RefExtIntersection, mydirX, 1)  '1:Max
        myHB.AppendHybridShape myExtremumPointIntersection
        myPart.Update
        
        Dim RefExtSketch2 As Reference
        Set RefExtSketch2 = myPart.CreateReferenceFromObject(mySketch2)
        Dim myExtremumPointSketch2 As HybridShapeExtremum
        Set myExtremumPointSketch2 = HSF.AddNewExtremum(RefExtSketch2, mydirX, 1)  '1:Max
        myHB.AppendHybridShape myExtremumPointSketch2
        
        
        myPart.Update
        
        '''''''Create Blend
        Dim RefCurveBlend1 As Reference
        Set RefCurveBlend1 = myPart.CreateReferenceFromObject(myHybridShapeIntersection)
        Dim RefCurveBlend2 As Reference
        Set RefCurveBlend2 = myPart.CreateReferenceFromObject(mySketch2)
        Dim RefmyExtremumPointIntersection As Reference
        Set RefmyExtremumPointIntersection = myPart.CreateReferenceFromObject(myExtremumPointIntersection)
        Dim RefmyExtremumPointSketch2 As Reference
        Set RefmyExtremumPointSketch2 = myPart.CreateReferenceFromObject(myExtremumPointSketch2)
        Dim RefSupport1 As Reference
        Set RefSupport1 = myPart.CreateReferenceFromObject(mySweepCircle)
         
        
        Dim myBlend As HybridShapeBlend
        Set myBlend = HSF.AddNewBlend
        myBlend.SetCurve 1, RefCurveBlend1
        myBlend.SetCurve 2, RefCurveBlend2
        myBlend.SetSupport 1, RefSupport1
        myBlend.SetClosingPoint 1, RefmyExtremumPointIntersection
        myBlend.SetClosingPoint 2, RefmyExtremumPointSketch2
        myBlend.SetOrientation 1, 1
        myBlend.SetOrientation 2, -1
        
        myBlend.SetContinuity 1, 1  '1:1,2;   1:0 Point,1 Tangency ,2 Curvature
        myHB2.AppendHybridShape myBlend
    myPart.Update
    
    'Step 4-7: Create Sketch2 Projection
    
        Dim RefPlaneProject As Reference
        Set RefPlaneProject = myPart.CreateReferenceFromObject(myOffsetPlane)
        Dim RefmySketch2Project As Reference
        Set RefmySketch2Project = myPart.CreateReferenceFromObject(mySketch2)
        Dim myProjectSketch2 As HybridShapeProject
        Set myProjectSketch2 = HSF.AddNewProject(RefmySketch2Project, RefPlaneProject)
        myProjectSketch2.Normal = True 'False: Type=Along a direction; 'True: Type=Normal
        'myProjectSketch2.Direction = HSF.AddNewDirectionByCoord(0, 0, 1)  'if .Normal = False
        myHB.AppendHybridShape myProjectSketch2
    myPart.Update
    
    'Step 4-8: Create Fill
         Dim FillBoundary1 As Reference
         Set FillBoundary1 = myPart.CreateReferenceFromObject(myProjectSketch2)
         Dim myFill As HybridShapeFill
         Set myFill = HSF.AddNewFill
         myFill.AddBound FillBoundary1
         'myFill.RemoveSupportAtPosition 1
         myHB2.AppendHybridShape myFill
        
    myPart.Update
    
    'Step 4-9: Create Bump
        
    
    ''     Bump
        Dim RefBoundaryBump As Reference
        Set RefBoundaryBump = myPart.CreateReferenceFromObject(myProjectSketch2)
                'Step 4-9-1: Create Center Point Bump
                Dim myCenterBumpPoint As HybridShapePointCenter
                Set myCenterBumpPoint = HSF.AddNewPointCenter(RefBoundaryBump)
                myHB.AppendHybridShape myCenterBumpPoint
                myPart.Update
         
        Dim RefBumpSurface As Reference
        Set RefBumpSurface = myPart.CreateReferenceFromObject(myFill)
        Dim RefCenterBump As Reference
        Set RefCenterBump = myPart.CreateReferenceFromObject(myCenterBumpPoint)
        Dim myBump As HybridShapeBump
        Set myBump = HSF.AddNewBump(RefBumpSurface)
        myBump.LimitCurve = RefBoundaryBump
        myBump.DeformationDistValue = 0.007
        myBump.ContinuityType = 0
        myBump.DeformationCenter = myCenterBumpPoint
        HSF.GSMVisibility RefBumpSurface, 0
        myHB2.AppendHybridShape myBump
        
        
    myPart.Update
    
    'Step 4-10: Create Hole
        'Step 4-10-1: Create Hole Center Point
            Dim RefPlaneHole As Reference
            Dim RefOriginPtHole As Reference
            Dim myCenterHolePoint As HybridShapePointOnPlane
            Dim RefPtCenterHole As Reference
            Dim myCircleHole As HybridShapeCircleCtrRad
            Dim RefSurfaceProject As Reference
            Dim RefmyHoleProject As Reference
            Dim myProjectHole As HybridShapeProject
            Dim RefSurfaceCut As Reference
            Dim RefCutter As Reference
            Dim mySplit As HybridShapeSplit
            
            
            '''''''''''''''''''''''''
            
            Set RefPlaneHole = myPart.CreateReferenceFromObject(myOffsetPlane)
            Set RefOriginPtHole = myPart.CreateReferenceFromObject(myOffsetPlane)
            
            
            
            
            Set myCenterHolePoint = HSF.AddNewPointOnPlane(RefPlaneHole, DistHole, 0)
            myCenterHolePoint.Point = RefCenterBump
            myHB.AppendHybridShape myCenterHolePoint
            
            myPart.Update
        
        'Step 4-10-2: Create Hole Circle
            
            Set RefPtCenterHole = myPart.CreateReferenceFromObject(myCenterHolePoint)
            
            Set myCircleHole = HSF.AddNewCircleCtrRad(RefPtCenterHole, RefPlaneHole, True, RHole)   'True: Geometry on Support
            myHB.AppendHybridShape myCircleHole
        myPart.Update
        myPart.Update
        
        'Step 4-10-2: Project Hole Circle on the Bump
            
            Set RefSurfaceProject = myPart.CreateReferenceFromObject(myBump)
            
            Set RefmyHoleProject = myPart.CreateReferenceFromObject(myCircleHole)
            
            Set myProjectHole = HSF.AddNewProject(RefmyHoleProject, RefSurfaceProject)
            myProjectHole.Normal = True 'False: Type=Along a direction; 'True: Type=Normal
            'myProjectSketch2.Direction = HSF.AddNewDirectionByCoord(0, 0, 1)  'if .Normal = False
            myHB.AppendHybridShape myProjectHole
        myPart.Update
        
        'Step 4-10-3: Split
        
        Set RefSurfaceCut = myPart.CreateReferenceFromObject(myBump)
        
        Set RefCutter = myPart.CreateReferenceFromObject(myProjectHole)
        
        Set mySplit = HSF.AddNewHybridSplit(RefSurfaceCut, RefCutter, -1) '1,-1: Orientation
    
        HSF.GSMVisibility RefSurfaceCut, 0   'makhfi kardan
        myHB2.AppendHybridShape mySplit
        myPart.Update
        
        HSF.GSMVisibility RefmyHoleProject, 0
        HSF.GSMVisibility RefCutter, 0
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''Create 7 Holes
        Dim Alfa As Double
            Alfa = 0
            
            For i = 1 To QtyHole - 1
                Alfa = Alfa + 360 / QtyHole
                Set myCenterHolePoint = HSF.AddNewPointOnPlane(RefPlaneHole, DistHole * Cos(Alfa * Pi / 180), DistHole * Sin(Alfa * Pi / 180))
                myCenterHolePoint.Point = RefCenterBump
                myHB.AppendHybridShape myCenterHolePoint
                        
                myPart.Update
                    
                'Step 4-10-2: Create Hole Circle
                        
                    Set RefPtCenterHole = myPart.CreateReferenceFromObject(myCenterHolePoint)
                        
                    Set myCircleHole = HSF.AddNewCircleCtrRad(RefPtCenterHole, RefPlaneHole, True, RHole)   'True: Geometry on Support
                    myHB.AppendHybridShape myCircleHole
                    myPart.Update
                    myPart.Update
                    
                    'Step 4-10-2: Project Hole Circle on the Bump
                        
                    Set RefSurfaceProject = myPart.CreateReferenceFromObject(mySplit)
                        
                    Set RefmyHoleProject = myPart.CreateReferenceFromObject(myCircleHole)
                        
                    Set myProjectHole = HSF.AddNewProject(RefmyHoleProject, RefSurfaceProject)
                    myProjectHole.Normal = True 'False: Type=Along a direction; 'True: Type=Normal
                    'myProjectSketch2.Direction = HSF.AddNewDirectionByCoord(0, 0, 1)  'if .Normal = False
                    myHB.AppendHybridShape myProjectHole
                    myPart.Update
                    
                    'Step 4-10-3: Split
                    
                    Set RefSurfaceCut = myPart.CreateReferenceFromObject(mySplit)
                    
             
                    
                    Set RefCutter = myPart.CreateReferenceFromObject(myProjectHole)
                    
                    Set mySplit = HSF.AddNewHybridSplit(RefSurfaceCut, RefCutter, -1) '1,-1: Orientation
                
                    HSF.GSMVisibility RefSurfaceCut, 0   'makhfi kardan
                    myHB2.AppendHybridShape mySplit
                    myPart.Update
                    'MsgBox Alfa
                HSF.GSMVisibility RefmyHoleProject, 0
                HSF.GSMVisibility RefCutter, 0
            Next i
            
            'Step 4-11: Create Fillet
                Dim RefFilletUp As Reference
                Set RefFilletUp = myPart.CreateReferenceFromObject(mySplit)
                Dim RefFilletDown As Reference
                Set RefFilletDown = myPart.CreateReferenceFromObject(myBlend)
                Dim myFillet 'kar nemikone
                Set myFillet = HSF.AddNewFilletBiTangent(RefFilletUp, RefFilletDown, RFillet, -1, 1, 1, 0) ''''-1,1 BOODE 1,-1
                HSF.GSMVisibility RefFilletUp, 0
                HSF.GSMVisibility RefFilletDown, 0
                myHB3.AppendHybridShape myFillet
                myPart.Update
            
            'Step 4-12: Fillet Color
                Dim MyList2 As Selection
                Set MyList2 = CATIA.ActiveDocument.Selection
                MyList2.Clear
                'Dim myObj As Object
                'Set myObj = myRef
                MyList2.Add myFillet
                MyList2.VisProperties.SetRealColor R_Ovary, G_Ovary, B_Ovary, 1
                myPart.Update
                
'Step 5: Flower
     
        'Step 5-1: Create Extremum Point of Intersection
        
        Dim mydirY As HybridShapeDirection
        Set mydirY = HSF.AddNewDirectionByCoord(0#, 1#, 0#)  'Y direction
        
        Dim RefExtSketch2Y As Reference
        Set RefExtSketch2Y = myPart.CreateReferenceFromObject(mySketch2)
        Dim myExtremumPointSketch2Y As HybridShapeExtremum
        Set myExtremumPointSketch2Y = HSF.AddNewExtremum(RefExtSketch2Y, mydirY, 1)  '1:Max
        myHB.AppendHybridShape myExtremumPointSketch2Y
        
        myPart.Update
        
        'Step 5-2: Create Plane Perp. to Sketch2 through Extremum Point
        Dim RefCurveSketch2 As Reference
        Set RefCurveSketch2 = myPart.CreateReferenceFromObject(mySketch2)
        Dim RefPtExtYSketch2 As Reference
        Set RefPtExtYSketch2 = myPart.CreateReferenceFromObject(myExtremumPointSketch2Y)
        Dim myPlaneNormaltoSketch2 As HybridShapePlaneNormal
        Set myPlaneNormaltoSketch2 = HSF.AddNewPlaneNormal(RefCurveSketch2, RefPtExtYSketch2)
        myHB.AppendHybridShape myPlaneNormaltoSketch2
               
        myPart.Update
        
        'Step 5-3: Create End Point of Flower
        Dim RefPlaneFlowerLine As Reference
        Set RefPlaneFlowerLine = myPart.CreateReferenceFromObject(myPlaneNormaltoSketch2)
        Dim myEndFlowerPoint As HybridShapePointOnPlane
        Set myEndFlowerPoint = HSF.AddNewPointOnPlane(RefPlaneFlowerLine, DeltaX, DeltaY)
        myEndFlowerPoint.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myEndFlowerPoint
                
        myPart.Update
        
        'Step 5-4: Create Flower Line
        Dim RefPtEndFlower As Reference
        Set RefPtEndFlower = myPart.CreateReferenceFromObject(myEndFlowerPoint)
        Dim RefPtStartFlower As Reference
        Set RefPtStartFlower = myPart.CreateReferenceFromObject(myExtremumPointSketch2Y)
        Dim myFlowerLine As HybridShapeLinePtPt
        Set myFlowerLine = HSF.AddNewLinePtPt(RefPtStartFlower, RefPtEndFlower)
        myHB.AppendHybridShape myFlowerLine
                
        myPart.Update
        
        'Step 5-5: Create Plane Perp. to Flower Line through Start Point
        Dim RefmyFlowerLine As Reference
        Set RefmyFlowerLine = myPart.CreateReferenceFromObject(myFlowerLine)
        Dim myPlaneNormaltoFlowerLine As HybridShapePlaneNormal
        Set myPlaneNormaltoFlowerLine = HSF.AddNewPlaneNormal(RefmyFlowerLine, RefPtExtYSketch2)
        myHB.AppendHybridShape myPlaneNormaltoFlowerLine
        
        myPart.Update
        
        'Step 5-6: Create Plane Perp. to Flower Line through End Point
        Dim myPlaneNormaltoFlowerLineEnd As HybridShapePlaneNormal
        Set myPlaneNormaltoFlowerLineEnd = HSF.AddNewPlaneNormal(RefmyFlowerLine, RefPtEndFlower)
        myHB.AppendHybridShape myPlaneNormaltoFlowerLineEnd
        
        myPart.Update
        
        'Step 5-7: Create Plane Perp to Flower Line
        Dim LenFlowerLine As Double
        LenFlowerLine = Sqr(DeltaX ^ 2 + DeltaY ^ 2)
        
        
        Dim myPlaneOffset1 As HybridShapePlaneOffset
        Set myPlaneOffset1 = HSF.AddNewPlaneOffset(myPlaneNormaltoFlowerLine, LenFlowerLine / 5, False)
        myHB.AppendHybridShape myPlaneOffset1
        
        Dim myPlaneOffset2 As HybridShapePlaneOffset
        Set myPlaneOffset2 = HSF.AddNewPlaneOffset(myPlaneNormaltoFlowerLine, 2 * LenFlowerLine / 5, False)
        myHB.AppendHybridShape myPlaneOffset2
        
        Dim myPlaneOffset3 As HybridShapePlaneOffset
        Set myPlaneOffset3 = HSF.AddNewPlaneOffset(myPlaneNormaltoFlowerLine, 3 * LenFlowerLine / 5, False)
        myHB.AppendHybridShape myPlaneOffset3
        
        Dim myPlaneOffset4 As HybridShapePlaneOffset
        Set myPlaneOffset4 = HSF.AddNewPlaneOffset(myPlaneNormaltoFlowerLine, 4 * LenFlowerLine / 5, False)
        myHB.AppendHybridShape myPlaneOffset4
        
        Dim myPlaneOffset5 As HybridShapePlaneOffset
        Set myPlaneOffset5 = HSF.AddNewPlaneOffset(myPlaneNormaltoFlowerLine, 4.5 * LenFlowerLine / 5, False)
        myHB.AppendHybridShape myPlaneOffset5
        
        myPart.Update
        
        'Step 5-8: Create CenterPoint of Arc1~7
        Dim RefPlaneS1 As Reference
        Set RefPlaneS1 = myPart.CreateReferenceFromObject(myPlaneNormaltoFlowerLine)
        Dim RefPlaneS2 As Reference
        Set RefPlaneS2 = myPart.CreateReferenceFromObject(myPlaneNormaltoFlowerLineEnd)
        Dim RefPlaneS3 As Reference
        Set RefPlaneS3 = myPart.CreateReferenceFromObject(myPlaneOffset1)
        Dim RefPlaneS4 As Reference
        Set RefPlaneS4 = myPart.CreateReferenceFromObject(myPlaneOffset2)
        Dim RefPlaneS5 As Reference
        Set RefPlaneS5 = myPart.CreateReferenceFromObject(myPlaneOffset3)
        Dim RefPlaneS6 As Reference
        Set RefPlaneS6 = myPart.CreateReferenceFromObject(myPlaneOffset4)
        Dim RefPlaneS7 As Reference
        Set RefPlaneS7 = myPart.CreateReferenceFromObject(myPlaneOffset5)
       
        Dim mydirZ As HybridShapeDirection
        Set mydirZ = HSF.AddNewDirectionByCoord(0#, 0#, 1#)  'Z direction
        
        
        Dim RefPointofArcs As Reference
        Set RefPointofArcs = RefPtExtYSketch2
        
        'Step 5-9: Create Arc1~7 and its Points
        
        Dim myCenterPointArc1 As HybridShapePointOnPlane
        Set myCenterPointArc1 = HSF.AddNewPointOnPlane(RefPlaneS1, XS1, YS1)
        myCenterPointArc1.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myCenterPointArc1
        
        myPart.Update
        
        Dim RefmyCenterPointArc1 As Reference
        Set RefmyCenterPointArc1 = myPart.CreateReferenceFromObject(myCenterPointArc1)
        Dim myArc1 As HybridShapeCircleCtrRad
        Set myArc1 = HSF.AddNewCircleCtrRadWithAngles(RefmyCenterPointArc1, RefPlaneS1, True, RS1, Alfa1Start, Alfa1End) 'True: Geometry on Support, StartAngle,EndAngle
        myHB.AppendHybridShape myArc1
        
        myPart.Update
            'Create Extremum Point of Arc1
            Dim RefExtArc1 As Reference
            Set RefExtArc1 = myPart.CreateReferenceFromObject(myArc1)
           
            Dim myExtremumPointArc1 As HybridShapeExtremum
            Set myExtremumPointArc1 = HSF.AddNewExtremum(RefExtArc1, mydirZ, 0)  '1:Max
            myHB.AppendHybridShape myExtremumPointArc1
        
            myPart.Update
            
            'Create Left and Right Point of Arc1
            Dim myLeftPtArc1 As HybridShapePointOnCurve
            Set myLeftPtArc1 = HSF.AddNewPointOnCurveFromPercent(RefExtArc1, 0, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myLeftPtArc1
            
            Dim myRightPtArc1 As HybridShapePointOnCurve
            Set myRightPtArc1 = HSF.AddNewPointOnCurveFromPercent(RefExtArc1, 1, 0) '0,1 Direction reverse
            myHB.AppendHybridShape myRightPtArc1
        
            myPart.Update
        
        'Arc2
        Dim myCenterPointArc2 As HybridShapePointOnPlane
        Set myCenterPointArc2 = HSF.AddNewPointOnPlane(RefPlaneS2, XS2, YS2)
        myCenterPointArc2.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myCenterPointArc2
                
        myPart.Update
        
        Dim RefmyCenterPointArc2 As Reference
        Set RefmyCenterPointArc2 = myPart.CreateReferenceFromObject(myCenterPointArc2)
        Dim myArc2 As HybridShapeCircleCtrRad
        Set myArc2 = HSF.AddNewCircleCtrRadWithAngles(RefmyCenterPointArc2, RefPlaneS2, True, RS2, Alfa2Start, Alfa2End) 'True: Geometry on Support, StartAngle,EndAngle
        myHB.AppendHybridShape myArc2
        
        myPart.Update
        
            'Create Extremum Point of Arc2
            Dim RefExtArc2 As Reference
            Set RefExtArc2 = myPart.CreateReferenceFromObject(myArc2)
           
            Dim myExtremumPointArc2 As HybridShapeExtremum
            Set myExtremumPointArc2 = HSF.AddNewExtremum(RefExtArc2, mydirZ, 0)  '1:Max
            myHB.AppendHybridShape myExtremumPointArc2
        
            myPart.Update
            
            'Create Left and Right Point of Arc2
            Dim myLeftPtArc2 As HybridShapePointOnCurve
            Set myLeftPtArc2 = HSF.AddNewPointOnCurveFromPercent(RefExtArc2, 0, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myLeftPtArc2
            
            Dim myRightPtArc2 As HybridShapePointOnCurve
            Set myRightPtArc2 = HSF.AddNewPointOnCurveFromPercent(RefExtArc2, 1, 0) '0,1 Direction reverse
            myHB.AppendHybridShape myRightPtArc2
        
            myPart.Update
        
        'Arc3
        Dim myCenterPointArc3 As HybridShapePointOnPlane
        Set myCenterPointArc3 = HSF.AddNewPointOnPlane(RefPlaneS3, XS3, YS3)
        myCenterPointArc3.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myCenterPointArc3
                
        myPart.Update
        
        Dim RefmyCenterPointArc3 As Reference
        Set RefmyCenterPointArc3 = myPart.CreateReferenceFromObject(myCenterPointArc3)
        Dim myArc3 As HybridShapeCircleCtrRad
        Set myArc3 = HSF.AddNewCircleCtrRadWithAngles(RefmyCenterPointArc3, RefPlaneS3, True, RS3, Alfa3Start, Alfa3End) 'True: Geometry on Support, StartAngle,EndAngle
        myHB.AppendHybridShape myArc3
        
        myPart.Update
        
            'Create Extremum Point of Arc3
            Dim RefExtArc3 As Reference
            Set RefExtArc3 = myPart.CreateReferenceFromObject(myArc3)
           
            Dim myExtremumPointArc3 As HybridShapeExtremum
            Set myExtremumPointArc3 = HSF.AddNewExtremum(RefExtArc3, mydirZ, 0)  '1:Max
            myHB.AppendHybridShape myExtremumPointArc3
        
            myPart.Update
            
            'Create Left and Right Point of Arc3
            Dim myLeftPtArc3 As HybridShapePointOnCurve
            Set myLeftPtArc3 = HSF.AddNewPointOnCurveFromPercent(RefExtArc3, 0, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myLeftPtArc3
            
            Dim myRightPtArc3 As HybridShapePointOnCurve
            Set myRightPtArc3 = HSF.AddNewPointOnCurveFromPercent(RefExtArc3, 1, 0) '0,1 Direction reverse
            myHB.AppendHybridShape myRightPtArc3
        
            myPart.Update
        
        
        'Arc4
        Dim myCenterPointArc4 As HybridShapePointOnPlane
        Set myCenterPointArc4 = HSF.AddNewPointOnPlane(RefPlaneS4, XS4, YS4)
        myCenterPointArc4.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myCenterPointArc4
                
        myPart.Update
        
        Dim RefmyCenterPointArc4 As Reference
        Set RefmyCenterPointArc4 = myPart.CreateReferenceFromObject(myCenterPointArc4)
        Dim myArc4 As HybridShapeCircleCtrRad
        Set myArc4 = HSF.AddNewCircleCtrRadWithAngles(RefmyCenterPointArc4, RefPlaneS4, True, RS4, Alfa4Start, Alfa4End) 'True: Geometry on Support, StartAngle,EndAngle
        myHB.AppendHybridShape myArc4
        
        myPart.Update
        
        'Create Extremum Point of Arc4
            Dim RefExtArc4 As Reference
            Set RefExtArc4 = myPart.CreateReferenceFromObject(myArc4)
           
            Dim myExtremumPointArc4 As HybridShapeExtremum
            Set myExtremumPointArc4 = HSF.AddNewExtremum(RefExtArc4, mydirZ, 0)  '1:Max
            myHB.AppendHybridShape myExtremumPointArc4
        
            myPart.Update
            
        'Create Left and Right Point of Arc4
            Dim myLeftPtArc4 As HybridShapePointOnCurve
            Set myLeftPtArc4 = HSF.AddNewPointOnCurveFromPercent(RefExtArc4, 0, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myLeftPtArc4
            
            Dim myRightPtArc4 As HybridShapePointOnCurve
            Set myRightPtArc4 = HSF.AddNewPointOnCurveFromPercent(RefExtArc4, 1, 0) '0,1 Direction reverse
            myHB.AppendHybridShape myRightPtArc4
        
            myPart.Update
        
        'Arc5
        Dim myCenterPointArc5 As HybridShapePointOnPlane
        Set myCenterPointArc5 = HSF.AddNewPointOnPlane(RefPlaneS5, XS5, YS5)
        myCenterPointArc5.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myCenterPointArc5
                
        myPart.Update
        
        Dim RefmyCenterPointArc5 As Reference
        Set RefmyCenterPointArc5 = myPart.CreateReferenceFromObject(myCenterPointArc5)
        Dim myArc5 As HybridShapeCircleCtrRad
        Set myArc5 = HSF.AddNewCircleCtrRadWithAngles(RefmyCenterPointArc5, RefPlaneS5, True, RS5, Alfa5Start, Alfa5End) 'True: Geometry on Support, StartAngle,EndAngle
        myHB.AppendHybridShape myArc5
        
        myPart.Update
        
            'Create Extremum Point of Arc5
            Dim RefExtArc5 As Reference
            Set RefExtArc5 = myPart.CreateReferenceFromObject(myArc5)
           
            Dim myExtremumPointArc5 As HybridShapeExtremum
            Set myExtremumPointArc5 = HSF.AddNewExtremum(RefExtArc5, mydirZ, 0)  '1:Max
            myHB.AppendHybridShape myExtremumPointArc5
        
            myPart.Update
            
            'Create Left and Right Point of Arc5
            Dim myLeftPtArc5 As HybridShapePointOnCurve
            Set myLeftPtArc5 = HSF.AddNewPointOnCurveFromPercent(RefExtArc5, 0, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myLeftPtArc5
            
            Dim myRightPtArc5 As HybridShapePointOnCurve
            Set myRightPtArc5 = HSF.AddNewPointOnCurveFromPercent(RefExtArc5, 1, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myRightPtArc5
            
            myPart.Update

        
        'Arc6
        Dim myCenterPointArc6 As HybridShapePointOnPlane
        Set myCenterPointArc6 = HSF.AddNewPointOnPlane(RefPlaneS6, XS6, YS6)
        myCenterPointArc6.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myCenterPointArc6
                
        myPart.Update
        
        Dim RefmyCenterPointArc6 As Reference
        Set RefmyCenterPointArc6 = myPart.CreateReferenceFromObject(myCenterPointArc6)
        Dim myArc6 As HybridShapeCircleCtrRad
        Set myArc6 = HSF.AddNewCircleCtrRadWithAngles(RefmyCenterPointArc6, RefPlaneS6, True, RS6, Alfa6Start, Alfa6End) 'True: Geometry on Support, StartAngle,EndAngle
        myHB.AppendHybridShape myArc6
        
        myPart.Update
        
        
        
        'Create Extremum Point of Arc6
            Dim RefExtArc6 As Reference
            Set RefExtArc6 = myPart.CreateReferenceFromObject(myArc6)
           
            Dim myExtremumPointArc6 As HybridShapeExtremum
            Set myExtremumPointArc6 = HSF.AddNewExtremum(RefExtArc6, mydirZ, 0)  '1:Max
            myHB.AppendHybridShape myExtremumPointArc6
        
            myPart.Update
            
        'Create Left and Right Point of Arc6
            Dim myLeftPtArc6 As HybridShapePointOnCurve
            Set myLeftPtArc6 = HSF.AddNewPointOnCurveFromPercent(RefExtArc6, 0, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myLeftPtArc6
            
            Dim myRightPtArc6 As HybridShapePointOnCurve
            Set myRightPtArc6 = HSF.AddNewPointOnCurveFromPercent(RefExtArc6, 1, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myRightPtArc6
            
            myPart.Update
        
        'Arc7
        Dim myCenterPointArc7 As HybridShapePointOnPlane
        Set myCenterPointArc7 = HSF.AddNewPointOnPlane(RefPlaneS7, XS7, YS7)
        myCenterPointArc7.Point = RefPtExtYSketch2
        myHB.AppendHybridShape myCenterPointArc7
                
        myPart.Update
        
        Dim RefmyCenterPointArc7 As Reference
        Set RefmyCenterPointArc7 = myPart.CreateReferenceFromObject(myCenterPointArc7)
        Dim myArc7 As HybridShapeCircleCtrRad
        Set myArc7 = HSF.AddNewCircleCtrRadWithAngles(RefmyCenterPointArc7, RefPlaneS7, True, RS7, Alfa7Start, Alfa7End) 'True: Geometry on Support, StartAngle,EndAngle
        myHB.AppendHybridShape myArc7
        
        myPart.Update
        
         'Create Extremum Point of Arc7
            Dim RefExtArc7 As Reference
            Set RefExtArc7 = myPart.CreateReferenceFromObject(myArc7)
           
            Dim myExtremumPointArc7 As HybridShapeExtremum
            Set myExtremumPointArc7 = HSF.AddNewExtremum(RefExtArc7, mydirZ, 0)  '1:Max
            myHB.AppendHybridShape myExtremumPointArc7
        
            myPart.Update
            
            'Create Left and Right Point of Arc7
            Dim myLeftPtArc7 As HybridShapePointOnCurve
            Set myLeftPtArc7 = HSF.AddNewPointOnCurveFromPercent(RefExtArc7, 0, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myLeftPtArc7
            
            Dim myRightPtArc7 As HybridShapePointOnCurve
            Set myRightPtArc7 = HSF.AddNewPointOnCurveFromPercent(RefExtArc7, 1, 0)  '0,1 Direction reverse
            myHB.AppendHybridShape myRightPtArc7
            
            myPart.Update
            
        'Step 5-10: Create Guide Splines for Loft (Multi-Section)
        
               Dim myDownSpline As HybridShapeSpline
               Set myDownSpline = HSF.AddNewSpline
            
               myDownSpline.AddPoint myExtremumPointArc1
               myDownSpline.AddPoint myExtremumPointArc3
               myDownSpline.AddPoint myExtremumPointArc4
               myDownSpline.AddPoint myExtremumPointArc5
               myDownSpline.AddPoint myExtremumPointArc6
               myDownSpline.AddPoint myExtremumPointArc7
               myDownSpline.AddPoint myExtremumPointArc2
               
               myHB.AppendHybridShape myDownSpline
                   
                myPart.Update
                
                
               Dim myLeftSpline As HybridShapeSpline
               Set myLeftSpline = HSF.AddNewSpline
            
               myLeftSpline.AddPoint myLeftPtArc1
               myLeftSpline.AddPoint myLeftPtArc3
               myLeftSpline.AddPoint myLeftPtArc4
               myLeftSpline.AddPoint myLeftPtArc5
               myLeftSpline.AddPoint myLeftPtArc6
               myLeftSpline.AddPoint myLeftPtArc7
               myLeftSpline.AddPoint myLeftPtArc2
               
               myHB.AppendHybridShape myLeftSpline
                   
                myPart.Update
                
              
               Dim myRightSpline As HybridShapeSpline
               Set myRightSpline = HSF.AddNewSpline
            
               myRightSpline.AddPoint myRightPtArc1
               myRightSpline.AddPoint myRightPtArc3
               myRightSpline.AddPoint myRightPtArc4
               myRightSpline.AddPoint myRightPtArc5
               myRightSpline.AddPoint myRightPtArc6
               myRightSpline.AddPoint myRightPtArc7
               myRightSpline.AddPoint myRightPtArc2
               
               myHB.AppendHybridShape myRightSpline
                   
                myPart.Update
                
            'Step 5-11: Create Loft
                Dim RefGuideLeft As Reference
                Dim RefGuideDown As Reference
                Dim RefGuideRight As Reference
                Set RefGuideLeft = myPart.CreateReferenceFromObject(myLeftSpline)
                Set RefGuideRight = myPart.CreateReferenceFromObject(myRightSpline)
                Set RefGuideDown = myPart.CreateReferenceFromObject(myDownSpline)


            
                Dim myFlowerLoft As HybridShapeLoft
                Set myFlowerLoft = HSF.AddNewLoft
                
                myFlowerLoft.AddSectionToLoft RefExtArc1, 1, Nothing   '1:orientation, Nothing: End Ref Point
                myFlowerLoft.AddSectionToLoft RefExtArc3, 1, Nothing
                myFlowerLoft.AddSectionToLoft RefExtArc4, 1, Nothing
                myFlowerLoft.AddSectionToLoft RefExtArc5, 1, Nothing
                myFlowerLoft.AddSectionToLoft RefExtArc6, 1, Nothing
                myFlowerLoft.AddSectionToLoft RefExtArc7, 1, Nothing
                myFlowerLoft.AddSectionToLoft RefExtArc2, 1, Nothing
                myFlowerLoft.AddGuide RefGuideDown
                myFlowerLoft.AddGuide RefGuideRight
                myFlowerLoft.AddGuide RefGuideLeft
                
                myHB3.AppendHybridShape myFlowerLoft
                   
                myPart.Update
            'Step 5 - 11-1: Petal  Color
                Dim MyListPetal0 As Selection
                Set MyListPetal0 = CATIA.ActiveDocument.Selection
                MyListPetal0.Clear
                MyListPetal0.Add myFlowerLoft
                MyListPetal0.VisProperties.SetRealColor R_Petal1, G_Petal1, B_Petal1, 1
                
            myPart.Update
                
            'Step 5-12: Create Circular Pattern of Petal
                Dim anyObject1 As AnyObject
                Set anyObject1 = myOffsetPlane
                
                Set reference1 = myPart.CreateReferenceFromName("")
                Set reference2 = myPart.CreateReferenceFromName("")
                Dim myCircPattern1
                Set myCircPattern1 = mySF.AddNewSurfacicCircPattern(anyObject1, 1, QtyPetal, 20#, 360 / QtyPetal, 1, 1, reference1, reference2, True, 0#, True, False)
                myCircPattern1.CircularPatternParameters = catInstancesandAngularSpacing
                
                Dim anyObject2 As AnyObject
                Set anyObject2 = myFlowerLoft
                
                myCircPattern1.ItemToCopy = anyObject2
                Set hybridShapePlaneOffset1 = myOffsetPlane
                Set reference3 = myPart.CreateReferenceFromObject(hybridShapePlaneOffset1)
                myCircPattern1.SetRotationAxis reference3
            
             myPart.UpdateObject myCircPattern1
             
             'Step 5 - 12-1: Petal array Color
                Dim MyListPetal1 As Selection
                Set MyListPetal1 = CATIA.ActiveDocument.Selection
                MyListPetal1.Clear
                MyListPetal1.Add myCircPattern1
                MyListPetal1.VisProperties.SetRealColor R_Petal1, G_Petal1, B_Petal1, 1
                
            myPart.Update
             
             'Step 5-13: Create Circular Pattern of Petal
                Dim anyObject11 As AnyObject
                Set anyObject11 = myPlaneNormaltoSketch2
                
                Set reference11 = myPart.CreateReferenceFromName("")
                Set Reference22 = myPart.CreateReferenceFromName("")
                Dim myCircPattern11
                Set myCircPattern11 = mySF.AddNewSurfacicCircPattern(anyObject11, 1, 2, 20#, AlfaPetal2, 1, 1, reference11, Reference22, False, 0#, True, False)
                myCircPattern11.CircularPatternParameters = catInstancesandAngularSpacing
                
                Dim anyObject22 As AnyObject
                Set anyObject22 = myFlowerLoft
                
                myCircPattern11.ItemToCopy = anyObject22
                Set hybridShapePlaneOffset11 = myPlaneNormaltoSketch2
                Set Reference33 = myPart.CreateReferenceFromObject(hybridShapePlaneOffset11)
                myCircPattern11.SetRotationAxis Reference33

            
             myPart.UpdateObject myCircPattern11
             
             
             
             'Step 5-14: Create Scale of Petal
             Dim RefSurfScaling As Reference
             Set RefSurfScaling = myPart.CreateReferenceFromObject(myCircPattern11)
             Dim myScaling As HybridShapeScaling
             Set myScaling = HSF.AddNewHybridScaling(RefSurfScaling, RefPtExtYSketch2, ScaleFactor)
             myHB2.AppendHybridShape myScaling
             HSF.GSMVisibility RefSurfScaling, 0
                   
                myPart.Update
                
                
            'Step 5-15: Create Rotate of Scaled Petal
            Dim RefPtDir1 As Reference
            Set RefPtDir1 = myPart.CreateReferenceFromObject(myExtremumPoint)
            Dim RefPtDir2 As Reference
            Set RefPtDir2 = myPart.CreateReferenceFromObject(myCenterBumpPoint)
            Dim myAXISLine As HybridShapeLinePtPt
            Set myAXISLine = HSF.AddNewLinePtPt(RefPtDir1, RefPtDir2)
            myHB.AppendHybridShape myAXISLine
            
            Dim RefmyAXISLine As Reference
            Set RefmyAXISLine = myPart.CreateReferenceFromObject(myAXISLine)
            
            Dim myDirROTATE As HybridShapeDirection
            Set myDirROTATE = HSF.AddNewDirection(RefmyAXISLine)

            
            Dim ReftoRotate As Reference
            
            
            Set ReftoRotate = myPart.CreateReferenceFromObject(myScaling)
            'Set mydirZ = HSF.AddNewDirectionByCoord(0#, 0#, 1#)
            Dim RefDirRotate As Reference
            'Set RefDirRotate = myPart.CreateReferenceFromObject(mydirZ)
            Set RefDirRotate = myPart.CreateReferenceFromObject(myDirROTATE)
            Dim HybridShapePetalRotate As HybridShapeRotate
            Set HybridShapePetalRotate = HSF.AddNewRotate(ReftoRotate, RefmyAXISLine, 180 / QtyPetal)
            myHB3.AppendHybridShape HybridShapePetalRotate
            myPart.InWorkObject = HybridShapePetalRotate
            
            HSF.GSMVisibility ReftoRotate, 0
            myPart.Update
            
         'Step 5 - 16: Petal2 Color
                Dim MyListPetal2 As Selection
                Set MyListPetal2 = CATIA.ActiveDocument.Selection
                MyListPetal2.Clear
                MyListPetal2.Add HybridShapePetalRotate
                MyListPetal2.VisProperties.SetRealColor R_Petal2, G_Petal2, B_Petal2, 1
                
            myPart.Update
        
        'Step 5-17: Create Circular Pattern of Scaled Petal
                Dim anyObject333 As AnyObject
                Set anyObject333 = myOffsetPlane
                
                Set Reference111 = myPart.CreateReferenceFromName("")
                Set Reference222 = myPart.CreateReferenceFromName("")
                Dim myCircPattern3
                Set myCircPattern3 = mySF.AddNewSurfacicCircPattern(anyObject333, 1, QtyPetal, 20#, 360 / QtyPetal, 1, 1, Reference111, Reference222, True, 0#, True, False)
                myCircPattern3.CircularPatternParameters = catInstancesandAngularSpacing
                
                Dim anyObject222 As AnyObject
                Set anyObject222 = HybridShapePetalRotate
                
                myCircPattern3.ItemToCopy = anyObject222
                Set hybridShapePlaneOffset1 = myOffsetPlane
                Set reference3 = myPart.CreateReferenceFromObject(hybridShapePlaneOffset1)
                myCircPattern3.SetRotationAxis reference3
            
             myPart.UpdateObject myCircPattern3
             
            'Step 5 - 18: Petal2 array Color
                Dim MyListPetal3 As Selection
                Set MyListPetal3 = CATIA.ActiveDocument.Selection
                MyListPetal3.Clear
                MyListPetal3.Add myCircPattern3
                MyListPetal3.VisProperties.SetRealColor R_Petal2, G_Petal2, B_Petal2, 1
                
            myPart.Update
        
                
       'HSF.GSMVisibility myHB, 0
       Dim MyListGS As Selection
                Set MyListGS = CATIA.ActiveDocument.Selection
                MyListGS.Clear
                MyListGS.Add myHB
                MyListGS.VisProperties.SetShow catVisPropertyNoShowAttr
                
            myPart.Update
 myPart.Update
        
End Sub
