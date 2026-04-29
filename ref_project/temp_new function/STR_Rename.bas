Attribute VB_Name = "STR_Rename"
Option Explicit

' -------------------------------------------------------------------------
' Module: STR_Rename
' Description: Batch renaming utilities for CATIA V5 Products (Structural Engineering)
' Author: Google DeepMind
' -------------------------------------------------------------------------

' Entry point for Adding Prefix to Product Instance Names
Sub STR_AddPrefix_InstanceName()
    Dim sPrefix As String
    sPrefix = InputBox("Enter prefix to add to Instance Name:", "Add Prefix", "STR_")
    If sPrefix = "" Then Exit Sub
    
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    
    If oSel.Count = 0 Then
        MsgBox "Please select at least one Product.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Integer
    Dim oProd As Product
    Dim iCount As Integer
    iCount = 0
    
    On Error Resume Next
    For i = 1 To oSel.Count
        Set oProd = oSel.Item(i).Value
        If Err.Number = 0 And TypeName(oProd) = "Product" Then
            oProd.Name = sPrefix & oProd.Name
            iCount = iCount + 1
        End If
        Err.Clear
    Next
    On Error GoTo 0
    
    MsgBox "Renamed " & iCount & " items.", vbInformation
End Sub

' Entry point for Find and Replace in Product Instance Names
Sub STR_FindReplace_InstanceName()
    Dim sFind As String
    Dim sReplace As String
    
    sFind = InputBox("Enter string to find in Instance Name:", "Find String")
    If sFind = "" Then Exit Sub
    
    sReplace = InputBox("Enter replacement string:", "Replace With")
    
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    
    If oSel.Count = 0 Then
        MsgBox "Please select at least one Product.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Integer
    Dim oProd As Product
    Dim iCount As Integer
    iCount = 0
    
    On Error Resume Next
    For i = 1 To oSel.Count
        Set oProd = oSel.Item(i).Value
        If Err.Number = 0 And TypeName(oProd) = "Product" Then
            If InStr(oProd.Name, sFind) > 0 Then
                oProd.Name = Replace(oProd.Name, sFind, sReplace)
                iCount = iCount + 1
            End If
        End If
        Err.Clear
    Next
    On Error GoTo 0
    
    MsgBox "Renamed " & iCount & " items.", vbInformation
End Sub

' Entry point for Adding Prefix to Part Number
Sub STR_AddPrefix_PartNumber()
    Dim sPrefix As String
    sPrefix = InputBox("Enter prefix to add to Part Number:", "Add Part Number Prefix", "PN_")
    If sPrefix = "" Then Exit Sub
    
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    
    If oSel.Count = 0 Then
        MsgBox "Please select at least one Product.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Integer
    Dim oProd As Product
    Dim iCount As Integer
    iCount = 0
    
    On Error Resume Next
    For i = 1 To oSel.Count
        Set oProd = oSel.Item(i).Value
        If Err.Number = 0 And TypeName(oProd) = "Product" Then
            oProd.PartNumber = sPrefix & oProd.PartNumber
            iCount = iCount + 1
        End If
        Err.Clear
    Next
    On Error GoTo 0
    
    MsgBox "Renamed " & iCount & " items.", vbInformation
End Sub
