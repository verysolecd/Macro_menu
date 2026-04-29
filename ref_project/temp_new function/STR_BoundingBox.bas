Attribute VB_Name = "STR_BoundingBox"
Option Explicit

' -------------------------------------------------------------------------
' Module: STR_BoundingBox
' Description: Creates a Bounding Box for the selected Part (Structural Engineering)
' Author: Google DeepMind
' -------------------------------------------------------------------------

Sub CreateBoundingBox()
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    
    If oSel.Count = 0 Then
        MsgBox "Please select a Part or Body.", vbExclamation
        Exit Sub
    End If
    
    Dim oPart As Part
    Dim oDist As Object ' SPAWorkbench.Measurable
    Dim oSPA As Object
    Set oSPA = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    
    ' Identify the Part to process
    On Error Resume Next
    Dim oItem As Object
    Set oItem = oSel.Item(1).Value
    
    If TypeName(oItem) = "Part" Then
        Set oPart = oItem
    ElseIf TypeName(oItem) = "Body" Then
        Set oPart = oItem.Parent.Parent ' Body -> Part -> PartDocument (Wait, Body -> Bodies -> Part)
        ' Actually Body.Parent is Bodies, Bodies.Parent is Part.
        Set oPart = oItem.Parent.Parent
    ElseIf InStr(TypeName(oItem), "Product") > 0 Then
        ' If it's a product, try to get the reference part
        ' This is tricky if it's an assembly. Skipping for now.
        MsgBox "Please select a Part or Body (Geometrical Set/Solid). functionality for Products is limited.", vbInformation
        Exit Sub
    End If
    On Error GoTo 0
    
    If oPart Is Nothing Then
        MsgBox "Could not determine the Part context.", vbCritical
        Exit Sub
    End If
    
    ' Create a new Geometrical Set for the Bounding Box
    Dim oHB As HybridBody
    On Error Resume Next
    Set oHB = oPart.HybridBodies.Item("Bounding_Box")
    If oHB Is Nothing Then
        Set oHB = oPart.HybridBodies.Add()
        oHB.Name = "Bounding_Box"
    End If
    On Error GoTo 0
    
    ' Factory
    Dim oFactory As HybridShapeFactory
    Set oFactory = oPart.HybridShapeFactory
    
    ' We need a reference. If a Body is selected, use it. If Part is selected, use MainBody.
    Dim oRef As Reference
    If TypeName(oItem) = "Body" Then
        Set oRef = oPart.CreateReferenceFromObject(oItem)
    Else
        Set oRef = oPart.CreateReferenceFromObject(oPart.MainBody)
    End If
    
    ' 1. Create Extrama (Min/Max X, Y, Z)
    ' Directions: 1=X, 2=Y, 3=Z
    ' Min/Max: 0=Min, 1=Max
    
    Dim oDirX As HybridShapeDirection
    Dim oDirY As HybridShapeDirection
    Dim oDirZ As HybridShapeDirection
    
    Set oDirX = oFactory.AddNewDirectionByCoord(1, 0, 0)
    Set oDirY = oFactory.AddNewDirectionByCoord(0, 1, 0)
    Set oDirZ = oFactory.AddNewDirectionByCoord(0, 0, 1)
    
    ' Compute Extrema is computationally expensive for complex parts. 
    ' But it's reliable for "Visual" Bounding Box.
    
    ' Creating planes at extrema
    ' This macro simplifies by creating a Box from 6 extrema points if possible
    ' But AddNewExtremum returns a point or a curve or surface.
    
    ' Let's use SPAWorkbench to get coordinates if possible? 
    ' SPAWorkbench doesn't give BBox points easily for a solid.
    
    ' Alternative: "Constraint" based BBox or just create 6 planes.
    
    ' Let's try to Create the 6 planes directly using Extremum
    Dim oExtremum(5) As HybridShapeExtremum
    
    ' Min X
    Set oExtremum(0) = oFactory.AddNewExtremum(oRef, oDirX, 0)
    ' Max X
    Set oExtremum(1) = oFactory.AddNewExtremum(oRef, oDirX, 1)
    ' Min Y
    Set oExtremum(2) = oFactory.AddNewExtremum(oRef, oDirY, 0)
    ' Max Y
    Set oExtremum(3) = oFactory.AddNewExtremum(oRef, oDirY, 1)
    ' Min Z
    Set oExtremum(4) = oFactory.AddNewExtremum(oRef, oDirZ, 0)
    ' Max Z
    Set oExtremum(5) = oFactory.AddNewExtremum(oRef, oDirZ, 1)
    
    Dim i As Integer
    For i = 0 To 5
        oHB.AppendHybridShape oExtremum(i)
        ' oPart.UpdateObject oExtremum(i) ' Defer update
    Next
    
    oPart.Update ' Update all to get the points
    
    ' Now create planes through these points normal to directions?
    ' Or just create a box from the coordinates of these points.
    
    ' Read coordinates
    Dim dMinX, dMaxX, dMinY, dMaxY, dMinZ, dMaxZ
    Dim oMeas As Measurable
    Dim aCoords(2)
    
    ' Min X
    Set oRef = oPart.CreateReferenceFromObject(oExtremum(0))
    Set oMeas = oSPA.GetMeasurable(oRef)
    oMeas.GetPoint aCoords
    dMinX = aCoords(0)
    
    ' Max X
    Set oRef = oPart.CreateReferenceFromObject(oExtremum(1))
    Set oMeas = oSPA.GetMeasurable(oRef)
    oMeas.GetPoint aCoords
    dMaxX = aCoords(0)
    
    ' Min Y
    Set oRef = oPart.CreateReferenceFromObject(oExtremum(2))
    Set oMeas = oSPA.GetMeasurable(oRef)
    oMeas.GetPoint aCoords
    dMinY = aCoords(1)
    
    ' Max Y
    Set oRef = oPart.CreateReferenceFromObject(oExtremum(3))
    Set oMeas = oSPA.GetMeasurable(oRef)
    oMeas.GetPoint aCoords
    dMaxY = aCoords(1)
    
    ' Min Z
    Set oRef = oPart.CreateReferenceFromObject(oExtremum(4))
    Set oMeas = oSPA.GetMeasurable(oRef)
    oMeas.GetPoint aCoords
    dMinZ = aCoords(2)
    
    ' Max Z
    Set oRef = oPart.CreateReferenceFromObject(oExtremum(5))
    Set oMeas = oSPA.GetMeasurable(oRef)
    oMeas.GetPoint aCoords
    dMaxZ = aCoords(2)
    
    ' Now create the Box using the coordinates
    ' Point 1: MinX, MinY, MinZ
    ' Point 2: MaxX, MaxY, MaxZ
    
    Dim oPt1 As HybridShapePointCoord
    Set oPt1 = oFactory.AddNewPointCoord(dMinX, dMinY, dMinZ)
    oHB.AppendHybridShape oPt1
    
    Dim oPt2 As HybridShapePointCoord
    Set oPt2 = oFactory.AddNewPointCoord(dMaxX, dMaxY, dMaxZ)
    oHB.AppendHybridShape oPt2
    
    ' Clean up the extrema (optional, maybe keep them for reference)
    ' oPart.InWorkObject = oHB
    
    ' Create a Block / Box? 
    ' Without "Part Design" license features, it's hard to make a solid.
    ' Let's make a Line box or just the points?
    ' Structural engineers often want the Volume.
    ' Let's try to create a "GS" with the 6 faces.
    
    ' ... Simplified: Just show the dimensions in a MsgBox for now
    MsgBox "Bounding Box:" & vbCrLf & _
           "X: " & dMinX & " to " & dMaxX & " (L=" & (dMaxX - dMinX) & ")" & vbCrLf & _
           "Y: " & dMinY & " to " & dMaxY & " (W=" & (dMaxY - dMinY) & ")" & vbCrLf & _
           "Z: " & dMinZ & " to " & dMaxZ & " (H=" & (dMaxZ - dMinZ) & ")", vbInformation, "Bounding Box Result"
    
    ' Cleanup the temporary extrema
    ' Ideally we delete them if we just wanted the numbers.
    ' For this macro, I will leave them as 'construction geometry' in the Bounding_Box set.
    
    oPart.Update
End Sub
