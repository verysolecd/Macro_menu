Attribute VB_Name = "Fix_Unfix_Position_3D"
Sub CATMain()
        
Fix_NoFix_Position.Show vbModeless

End Sub
'***************************************************************
'Macro to Delete existing constraints, and create FixInSpace constraints
'Operates recursively on all levels within the current document
'By Mark Forbes
'Original Version 0.1.0 May 19 2010
'Version 0.1.1 June 8 2010 Handles deactivated components
'Version 0.1.2 Sept 16 2010 Handles CATProducts only once, better error handling
'   on deactivated products, and prompts to skip Flexible Subassemblies
'Version 0.1.3 Feb 1 2011 Hide constraints, Late Bind oCurrentProd for R19
'Version 0.1.4 March 10 2011 : Adding more error handling
'***************************************************************

'Public oList As Variant
'Public oSelection As Selection
'Public oVisProp As VisPropertyset
'
'Option Explicit

'Sub CATMain()
'IntCATIA
''Declarations
'Dim oTopDoc As Document
'Dim oTopProd As ProductDocument
'Dim oCurrentProd As Object
''Dim oVisProp As VisPropertySet
'
'
'
''Check if the active document is an assembly, else exit
'Set oTopDoc = CATIA.ActiveDocument
'If Right(oTopDoc.Name, 7) <> "Product" Then
'    MsgBox "Active document should be a product"
'    Exit Sub
'End If
'
'Set oSelection = oTopDoc.Selection
'Set oVisProp = oSelection.VisProperties
'
'Set oCurrentProd = oTopDoc.Product
'
'Set oList = CreateObject("Scripting.dictionary")
'
''CATIA.StatusBar = "Working On" & " " & oCurrentProd.Name
'
'Call FixSingleLevel(oCurrentProd)    'Call the subroutine, it is a recursive loop
'
'
'CATIA.StatusBar = "Macro Done"
'MsgBox "Fixing Macro Finished"
'
'
'End Sub

Private Sub FixSingleLevel(ByRef oCurrentProd As Object)

On Error Resume Next

'More declarations
Dim ItemToFix As Product
Dim iProdCount As Integer
Dim i As Integer
Dim j As Integer
Dim oConstraints 'As Constraints
Dim oReference As Reference
Dim sItemName As String
Dim constraint1 'As Constraint
Dim pActivation 'As Parameter
Dim n, m As Integer
Dim sActivationName As String

Err.Clear
CATIA.StatusBar = "Working On" & " " & oCurrentProd.Name
Set oCurrentProd = oCurrentProd.ReferenceProduct
iProdCount = oCurrentProd.Products.Count
Set oConstraints = oCurrentProd.Connections("CATIAConstraints")

n = oConstraints.Count  'Remove Existing Constraints
m = n
For i = 1 To m
    oConstraints.Remove (n)
    n = n - 1
Next


For i = 1 To iProdCount                             'Cycle through the assembly's children

    Set ItemToFix = oCurrentProd.Products.Item(i)

CreateReference:
    
    sItemName = ItemToFix.Name
    CATIA.StatusBar = "Working On " & oCurrentProd.Name & " / " & sItemName
    
    Set oReference = oCurrentProd.CreateReferenceFromName(sItemName & "/!" & "/")
   
    Set constraint1 = oConstraints.AddMonoEltCst(0, oReference)
    constraint1.ReferenceType = 1 'catCstRefTypeFixInSpace
    
    oSelection.Add constraint1      'Set visibility to hidden
    oVisProp.SetShow catVisPropertyNoShowAttr
    oSelection.Clear

RecursionCall:
    If ItemToFix.Products.Count <> 0 Then        'Recursive Call
        If oList.exists(ItemToFix.PartNumber) Then GoTo Finish
        
        If ItemToFix.PartNumber = ItemToFix.ReferenceProduct.Parent.Product.PartNumber Then oList.Add ItemToFix.PartNumber, 1
        Call FixSingleLevel(ItemToFix)
    End If

Finish:

Next

GoTo End1:

'*****Error Handling
Err_Handling:

    
sActivationName = oCurrentProd.Name + "\" + ItemToFix.Name + "\Component Activation State"  'Build the reference Name
 Set pActivation = ItemToFix.Parameters.GetItem(sActivationName)
    If pActivation.ValueAsString = "false" Then
       CATIA.StatusBar = "Error, Try To Activate " & ItemToFix.Name 'Tell the user what is happening
       pActivation.ValuateFromString ("true")
       
     ElseIf pActivation.ValueAsString = "true" Then  'Assume this is a flexibe component
       j = MsgBox("Error on " & oCurrentProd.Name + "\" & ItemToFix.Name & vbCrLf _
            & "This element may be a flexible component, have an invalid" & vbCrLf _
            & "Instance Name, or other error" & vbCrLf _
            & vbCrLf _
            & "Skip component and continue?", vbOKCancel, "Error")
            Err.Clear
       If j = 1 Then Resume RecursionCall
       If j = 2 Then
        CATIA.StatusBar = "Fix All Aborted"
        End
       End If
       Else: Resume RecursionCall
      
    End If
'*****End of Error Handling

End2:
Resume

End1:
End Sub

