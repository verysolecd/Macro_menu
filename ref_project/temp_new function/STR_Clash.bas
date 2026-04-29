Attribute VB_Name = "STR_Clash"
Option Explicit

' -------------------------------------------------------------------------
' Module: STR_Clash
' Description: Quick creation of CATIA DMU Clash Analysis (SpaceAnalysis)
' Author: Google DeepMind
' -------------------------------------------------------------------------

Sub QuickClash_SelectionVsAll()
    ' 1. Get Group 1 from current selection
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    
    If oSel.Count = 0 Then
        MsgBox "Please select the components for the 'Main Group' (Group 1) first.", vbExclamation
        Exit Sub
    End If
    
    Dim colGroup1 As New Collection
    Dim i As Integer
    Dim oProd As Object
    
    On Error Resume Next
    For i = 1 To oSel.Count
        Set oProd = oSel.Item(i).Value
        If TypeName(oProd) = "Product" Then
            colGroup1.Add oProd
        End If
    Next
    On Error GoTo 0
    
    If colGroup1.Count = 0 Then
        MsgBox "No valid products selected.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Get Group 2 (Rest of the visible products in the active level)
    Dim colGroup2 As New Collection
    Dim oRoot As Product
    Set oRoot = GetRootProduct()
    
    If oRoot Is Nothing Then Exit Sub
    
    Dim oChild As Product
    Dim vItem As Variant
    Dim bFound As Boolean
    
    For i = 1 To oRoot.Products.Count
        Set oChild = oRoot.Products.Item(i)
        bFound = False
        For Each vItem In colGroup1
           If vItem.Name = oChild.Name Then
               bFound = True
               Exit For
           End If
        Next
        
        If Not bFound Then
            colGroup2.Add oChild
        End If
    Next
    
    If colGroup2.Count = 0 Then
        MsgBox "No other products found to check against.", vbExclamation
        Exit Sub
    End If
    
    ' 3. Ask for Clearance
    Dim sClearance As String
    sClearance = InputBox("Enter clearance value (mm) or leave empty for Contact+Clash only:", "Clash Clearance", "0")
    
    Dim bClearance As Boolean
    Dim dValidClearance As Double
    If IsNumeric(sClearance) And Val(sClearance) > 0 Then
        bClearance = True
        dValidClearance = Val(sClearance)
    Else
        bClearance = False
        dValidClearance = 0
    End If
    
    ' 4. Create Clash
    CreateClashAnalysis colGroup1, colGroup2, bClearance, dValidClearance
End Sub

Sub QuickClash_SelectionVsSelection()
    MsgBox "Please select components for Group 1, then click OK.", vbInformation
    
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    
    If oSel.Count = 0 Then Exit Sub
    
    Dim colGroup1 As New Collection
    Dim i As Integer
    
    For i = 1 To oSel.Count
        If TypeName(oSel.Item(i).Value) = "Product" Then
            colGroup1.Add oSel.Item(i).Value
        End If
    Next
    
    If colGroup1.Count = 0 Then Exit Sub
    
    oSel.Clear
    MsgBox "Group 1 captured (" & colGroup1.Count & " items). Now select components for Group 2, then click OK.", vbInformation
    
    Dim colGroup2 As New Collection
    For i = 1 To oSel.Count
        If TypeName(oSel.Item(i).Value) = "Product" Then
            colGroup2.Add oSel.Item(i).Value
        End If
    Next
    
     If colGroup2.Count = 0 Then MsgBox "Group 2 is empty!", vbExclamation: Exit Sub
     
    ' Create Clash
    CreateClashAnalysis colGroup1, colGroup2, False, 0
End Sub

' Helper function to get the root product
Function GetRootProduct() As Product
    On Error Resume Next
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    If TypeName(oDoc) = "ProductDocument" Then
        Set GetRootProduct = oDoc.Product
    Else
        Set GetRootProduct = Nothing
    End If
    On Error GoTo 0
End Function

' Main function to create the clash
Sub CreateClashAnalysis(colGroup1 As Collection, colGroup2 As Collection, bClearance As Boolean, dClearanceVal As Double)
    Dim oRoot As Product
    Set oRoot = GetRootProduct()
    
    If oRoot Is Nothing Then
        MsgBox "Active document must be a Product.", vbCritical
        Exit Sub
    End If
    
    On Error Resume Next
    Dim cClashes As Clashes
    Set cClashes = oRoot.GetTechnologicalObject("Clashes")
    
    If cClashes Is Nothing Then
        MsgBox "Could not access Clashes collection. Ensure DMU Space Analysis license is available.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Create Groups
    Dim cGroups As Groups
    Set cGroups = oRoot.GetTechnologicalObject("Groups")
    
    Dim oGroup1 As Group
    Dim oGroup2 As Group
    
    ' Create Group 1
    Set oGroup1 = cGroups.Add()
    Dim vItem
    For Each vItem In colGroup1
        oGroup1.AddExplicit vItem
    Next
    
    ' Create Group 2
    Set oGroup2 = cGroups.Add()
    For Each vItem In colGroup2
        oGroup2.AddExplicit vItem
    Next
    
    ' Create Clash
    Dim oClash As Clash
    Set oClash = cClashes.Add()
    oClash.ComputationType = catClashComputationTypeBetweenTwoSelections
    
    oClash.FirstGroup = oGroup1
    oClash.SecondGroup = oGroup2
    
    If bClearance Then
        oClash.ComputationType = catClashComputationTypeClearancePlusContactPlusClash
        oClash.Clearance = dClearanceVal
    Else
        oClash.ComputationType = catClashComputationTypeContactPlusClash
    End If
    
    oClash.Compute
    
    MsgBox "Clash Analysis Created: " & oClash.Name, vbInformation
End Sub
