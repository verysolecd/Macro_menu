Private mdict

Sub tube()
    ' Initialize global dictionary
    Set mdict = KCL.InitDic
    
    Dim oDoc ' As Document
    Set oDoc = CATIA.ActiveDocument
    
    If TypeName(oDoc) <> "PartDocument" Then
        MsgBox "Active document is not a Part."
        Exit Sub
    End If
    
    Dim oprt ' As Part
    Set oprt = oDoc.Part
    
    Dim HSF ' As HybridShapeFactory
    Set HSF = oprt.HybridShapeFactory
    
    ' --- Method 1: Get Shapes by Recursion (Agent Recommended) ---
    ' This method uses recursion and interface casting via Selection to ensure properties like TopOffset are available.
    Dim oSel ' As Selection
    Set oSel = oDoc.Selection
    oSel.Clear
    
    ' Traverse all container types
    ' HybridBodies (Geometrical Sets)
    For Each hb In oprt.HybridBodies
        Call GetShapesByRecursion(hb, HSF, oSel)
    Next
    ' Bodies (Solid Bodies)
    For Each b In oprt.Bodies
        Call GetShapesByRecursion(b, HSF, oSel)
    Next
    ' OrderedGeometricalSets
    For Each ogs In oprt.OrderedGeometricalSets
        Call GetShapesByRecursion(ogs, HSF, oSel)
    Next
    
    ' --- Method 2: Get Shapes by Parameters (User Original Logic) ---
    ' This method scans all parameters to find their parents.
    Call GetShapesByParameters(oprt, HSF)


    ' --- Process Results ---
    Dim lst
    Set lst = KCL.InitLst
    
    ' Helper for measurement
    Dim spa ' As SPAWorkbench
    Set spa = oDoc.GetWorkbench("SPAWorkbench")

    Debug.Print "--- Measurement Results ---"
    For Each key In mdict.Keys
        Dim shp
        Set shp = mdict(key)
        lst.Add shp
        
        ' --- New Request: Check for ThickSurface and Measure Parent Sweep Spine ---
        Dim parentSweep
        Set parentSweep = GetParentSweep(shp)
        
        If Not parentSweep Is Nothing Then
            Dim spineLen As Double
            spineLen = GetSpineLength(parentSweep, spa)
            Debug.Print "ThickSurface: " & shp.Name & " | Parent Sweep: " & parentSweep.Name & " | Spine Length: " & FormatNumber(spineLen, 2)
        Else
            ' Debug.Print "ThickSurface: " & shp.Name & " | Parent Sweep: Not Found"
        End If
    Next

    MsgBox "Total Unique Shapes Found: " & mdict.Count
End Sub

' --- Helper to find the Parent Sweep of a ThickSurface ---
Function GetParentSweep(oThickSurface)
    On Error Resume Next
    Set GetParentSweep = Nothing
    
    ' Try to get the support surface. 
    ' HybridShapeThickSurface usually has a property named 'Surface' or can be accessed via Reference.
    ' Note: In some API versions, this might require getting the 'Feature' parent or specific method.
    ' We attempt the direct property '.Surface' which is common for offsetting features.
    
    Dim support
    Set support = NOTHING
    
    ' Attempt 1: Direct .Surface property (Most likely for Offset/ThickSurface)
    Set support = oThickSurface.Surface
    
    ' Attempt 2: If .Surface fails, check if we can get it via inputs (if HSF support GetInputs - usually no)
    
    If Not support Is Nothing Then
        ' Check if the support is a Sweep (Type check or Loop inputs)
        ' Simplest check: Name contains "Sweep" or try to measure its spine
        Set GetParentSweep = support
    End If
    On Error GoTo 0
End Function
Sub GetShapesByRecursion(container, HSF, oSel)
    On Error Resume Next
    Dim Shps ' As HybridShapes
    Set Shps = container.HybridShapes
    
    If Err.Number = 0 Then
        For Each Shp In Shps
            If HSF.GetGeometricalFeatureType(Shp) = 7 Then
                Dim internalName As String
                internalName = KCL.GetInternalName(Shp)
                
                If Not mdict.Exists(internalName) Then
                    ' FIX: Use Selection to cast the HybridShape to its specific interface (e.g. ThickSurface)
                    oSel.Add Shp
                    Dim realShp
                    Set realShp = oSel.Item(1).Value
                    oSel.Clear
                    
                    mdict.Add internalName, realShp
                End If
            End If
        Next
    End If
    Err.Clear
    
    ' Recursive Calls
    ' 1. HybridBodies
    If container.HybridBodies.Count > 0 Then
        For Each childHB In container.HybridBodies
            Call GetShapesByRecursion(childHB, HSF, oSel)
        Next
    End If
    ' 2. Bodies (if nested)
    ' Note: Bodies usually don't nest Bodies in same collection, but for safety
    ' standard container structure usually sufficient.
    ' 3. OrderedGeometricalSets
    If container.OrderedGeometricalSets.Count > 0 Then
        For Each childOGS In container.OrderedGeometricalSets
            Call GetShapesByRecursion(childOGS, HSF, oSel)
        Next
    End If
    On Error GoTo 0
End Sub

' --- Method 2 Implementation ---
Sub GetShapesByParameters(oprt, HSF)
    Dim paras ' As Parameters
    Set paras = oprt.Parameters
    
    For Each p In paras
        On Error Resume Next
        Dim parentObj
        Set parentObj = p.Parent
        
        If Err.Number = 0 Then
            If TypeName(parentObj) <> "Parameters" Then
                If HSF.GetGeometricalFeatureType(parentObj) = 7 Then
                    Dim oname As String
                    oname = KCL.GetInternalName(parentObj)
                    If Not mdict.Exists(oname) Then
                        mdict.Add oname, parentObj
                        ' Debug.Print "Found by Param: " & oname
                    End If
                End If
            End If
        End If
        Err.Clear
    Next
    On Error GoTo 0
End Sub

' --- Helper for Sweep Spine Measurement ---
Function GetSpineLength(oSweep, spa)
    ' Calculates the length of the spine/center curve of a sweep surface
    ' Requires SPAWorkbench
    On Error Resume Next
    GetSpineLength = -1
    
    Dim spineRef ' As Reference
    Set spineRef = Nothing
    
    ' Try common spine properties based on Sweep type
    ' 1. Explicit / Line / Circle sweeps often have a 'Spine' property
    Set spineRef = oSweep.Spine
    
    ' 2. If no Spine property, try CenterCurve (common in some Explicit sweeps)
    If spineRef Is Nothing Then
        Set spineRef = oSweep.CenterCurve
    End If
    
    ' 3. If still nothing, it might be a Line sweep with GuideCurve as spine equivalent
    If spineRef Is Nothing Then
        Set spineRef = oSweep.GuideCurve
    End If
    
    If Not spineRef Is Nothing Then
        Dim meas ' As Measurable
        Set meas = spa.GetMeasurable(spineRef)
        GetSpineLength = meas.Length
    End If
    On Error GoTo 0
End Function

' Function to attempt retrieval of Sweep from ThickSurface (Limited VBA capabilities)
Function GetParentSweepFromThickSurface(oThickSurface)
    ' This is difficult in pure VBA without specific links.
    ' Usually relies on naming convention or checking inputs if accessible.
    ' Placeholder for user logic.
    Set GetParentSweepFromThickSurface = Nothing
End Function
