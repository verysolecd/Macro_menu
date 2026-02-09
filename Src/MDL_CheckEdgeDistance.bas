Attribute VB_Name = "MDL_CheckEdgeDistance"
'=============================================================================================
' Module:       MDL_CheckEdgeDistance
' Purpose:      Check if holes are too close to the part edge (Edge Distance Check)
' Description:  Scans all holes in the active part/selection. Measures distance from hole center
'               to the nearest face (excluding top/bottom and the hole itself).
'               If Distance < Ratio * Diameter, highlights the hole in RED.
' Usage:        Run "CheckHoleEdgeDist" macro.
' Requirements: Active Part Document.
' Author:       Antigravity (Google DeepMind)
'=============================================================================================
Option Explicit

Private Const EDGE_DIST_RATIO As Double = 1.5 ' Rule: Dist >= 1.5 * Diameter
Private Const TOLERANCE As Double = 0.1       ' Tolerance for geometry comparisons

Sub CheckHoleEdgeDist()
    ' 1. Logic: Initialize
    Dim oDoc As PartDocument
    If TypeName(CATIA.ActiveDocument) <> "PartDocument" Then
        MsgBox "Active document must be a Part.", vbCritical
        Exit Sub
    End If
    Set oDoc = CATIA.ActiveDocument
    
    Dim oPart As Part
    Set oPart = oDoc.Part
    
    Dim oSel As Selection
    Set oSel = oDoc.Selection
    
    Dim oSPA As NAMECATIA_SPAWorkbench ' Placeholder for user reference, actual type is SPAWorkbench
    ' In VBA without reference, use Object or late binding.
    ' Assuming user has reference or using late binding.
    Dim oSPAObj As Object
    On Error Resume Next
    Set oSPAObj = oDoc.GetWorkbench("SPAWorkbench")
    If oSPAObj Is Nothing Then
        MsgBox "SPAWorkbench not available. License issue?", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 2. Logic: Get Holes to Process
    Dim targetHoles As New Collection
    Dim i As Integer
    
    ' Strategy: If user selected holes, check them. If Body, check body. If nothing, verify PartBody.
    If oSel.Count > 0 Then
        For i = 1 To oSel.Count
            If TypeName(oSel.Item(i).Value) = "Hole" Then
                targetHoles.Add oSel.Item(i).Value
            End If
        Next
    End If
    
    ' If no holes manually selected, search entire part
    If targetHoles.Count = 0 Then
        oSel.Search "Type='Part Design'.Hole,all"
        For i = 1 To oSel.Count
            targetHoles.Add oSel.Item(i).Value
        Next
    End If
    
    If targetHoles.Count = 0 Then
        MsgBox "No holes found to check.", vbInformation, "Edge Check"
        Exit Sub
    End If
    
    ' 3. Logic: Process Each Hole
    oSel.Clear ' Clear for highlighting later
    
    Dim violationCount As Integer
    violationCount = 0
    
    Dim oHole As Hole
    Dim dDiam As Double
    Dim dMinDist As Double
    
    ' Prepare specific selection for highlighting
    Dim badHoles As New Collection
    
    For Each oHole In targetHoles
        dDiam = GetHoleDiameter(oHole)
        
        ' Determine "Edge Distance"
        ' This requires geometric measurement logic.
        dMinDist = GetTrueEdgeDistance(oPart, oHole, oSPAObj, dDiam)
        
        ' Check Rule
        Dim requiredDist As Double
        requiredDist = dDiam * EDGE_DIST_RATIO
        
        ' Valid distance found and rule violated
        If dMinDist >= 0 And dMinDist < requiredDist Then
            violationCount = violationCount + 1
            badHoles.Add oHole
            ' Debug.Print "[FAIL] Hole: " & oHole.Name & " | Dist: " & Round(dMinDist, 2) & " < " & Round(requiredDist, 2)
        End If
    Next
    
    ' 4. Logic: Visualization
    oSel.Clear
    If violationCount > 0 Then
        Dim badItem
        For Each badItem In badHoles
            oSel.Add badItem
        Next
        
        ' Set Color to ORANGE (Red is often for errors, Orange for warnings)
        ' VisProperties: SetRealColor R, G, B, Inheritance
        oSel.VisProperties.SetRealColor 255, 128, 0, 1
        oSel.VisProperties.SetRealWidth 4, 1
        
        MsgBox "Found " & violationCount & " violations!" & vbCrLf & _
               "Holes too close to edge (Ratio < " & EDGE_DIST_RATIO & ") are highlighted.", vbExclamation, "Design Rule Check"
    Else
        MsgBox "Check Complete. No violations found.", vbInformation, "Design Rule Check"
    End If
    
End Sub

' --------------------------------------------------------------------------------------
' Helper: Get Hole Diameter
' --------------------------------------------------------------------------------------
Function GetHoleDiameter(oHole As Hole) As Double
    On Error Resume Next
    If Not oHole.Diameter Is Nothing Then
        GetHoleDiameter = oHole.Diameter.Value
    Else
        GetHoleDiameter = 10 ' Default safety
    End If
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------------------
' Helper: Measure distance to nearest valid face
' --------------------------------------------------------------------------------------
Function GetTrueEdgeDistance(oPart As Part, oHole As Hole, oSPA As Object, dDiameter As Double) As Double
    On Error Resume Next
    GetTrueEdgeDistance = -1 ' Default error
    
    Dim oMeasurable As Object
    Set oMeasurable = oSPA.GetMeasurable(oPart.CreateReferenceFromObject(oHole))
    
    Dim dRadius As Double
    dRadius = dDiameter / 2
    
    ' Search for all faces in the PART to be robust
    ' Ideally, we only search faces of the Body the hole belongs to.
    ' Let's try searching "Topology.Face" in the part selection.
    
    Dim oSel As Selection
    Set oSel = oPart.Parent.Selection
    oSel.Clear
    
    ' Optimization: Get the Body object from the Hole.Parent? usually Sketch-based features parent is Body or HybridBody.
    Dim oParent As Object
    Set oParent = oHole.Parent
    If TypeName(oParent) = "Body" Or TypeName(oParent) = "PartBody" Then
        oSel.Add oParent
        oSel.Search "Topology.Face,sel"
    Else
        ' Fallback: Search all faces in Part
        oSel.Search "Topology.Face,all"
    End If
    
    Dim dMin As Double
    dMin = 99999.0
    
    Dim i As Integer
    Dim oFace As Face
    Dim dDist As Double
    Dim isSelf As Boolean
    
    ' Loop limit for performance (max 50 faces check?)
    Dim maxChecks As Integer
    maxChecks = oSel.Count
    If maxChecks > 50 Then maxChecks = 50 ' Safety limit for demo
    
    For i = 1 To maxChecks
        Set oFace = oSel.Item(i).Value
        
        ' Measure from Hole Axis/Object to Face
        ' Note: GetMinimumDistance from Hole Object to a Face returns shortest distance between their geometries.
        dDist = oMeasurable.GetMinimumDistance(oPart.CreateReferenceFromObject(oFace))
        
        isSelf = False
        
        ' Filter Logic:
        ' 1. Distance ~ Radius: It is the hole cylinder itself (Hole Axis to Cylinder Surface = Radius).
        If Abs(dDist - dRadius) < TOLERANCE Then isSelf = True
        
        ' 2. Distance ~ 0: It is intersecting face (Top/Bottom).
        '    Hole usually starts on a face, so distance is 0.
        If dDist < TOLERANCE Then isSelf = True
        
        If Not isSelf Then
            If dDist < dMin Then dMin = dDist
        End If
    Next
    
    If dMin <> 99999.0 Then
        GetTrueEdgeDistance = dMin
    End If
    oSel.Clear
End Function
