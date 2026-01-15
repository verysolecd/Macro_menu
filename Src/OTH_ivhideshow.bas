Attribute VB_Name = "OTH_ivhideshow"
Option Explicit

' Module-level state
Private isHidden As Boolean
Private hiddenElements As Collection

Sub CATMain()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub

    ' Ensure PDM instance
    If pdm Is Nothing Then Set pdm = New Cls_PDM

    Set osel = pdm.msel
    Dim oDoc As Document, cGroups As Object, oGroup As Object
    Set oDoc = CATIA.ActiveDocument
    Set cGroups = rootPrd.GetTechnologicalObject("Groups")

    If Not isHidden Then
        ' First execution: hide inverted selection and store elements
        Set oGroup = cGroups.AddFromSel
        oGroup.ExtractMode = 1
        oGroup.FillSelWithInvert
        cGroups.Remove 1
        Set cGroups = Nothing

        ' Store the elements that were selected (now hidden)
        Set hiddenElements = New Collection
        Dim sel As Selection
        Set sel = oDoc.Selection
        Dim i As Long
        For i = 1 To sel.Count
            hiddenElements.Add sel.Item(i)
        Next i

        ' Hide them
        sel.VisProperties.SetShow 1
        isHidden = True
    Else
        ' Second execution: restore visibility
        If Not hiddenElements Is Nothing Then
            Dim elem As Object
            For Each elem In hiddenElements
                elem.VisProperties.SetShow 0
            Next elem
        End If
        Set hiddenElements = Nothing
        isHidden = False
    End If
End Sub
