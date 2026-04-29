VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Fix_NoFix_Position 
   Caption         =   "全部固定|解除固定"
   ClientHeight    =   1800
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   2265
   OleObjectBlob   =   "Fix_NoFix_Position.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Fix_NoFix_Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oList As Variant
Dim oVisProp As VisPropertyset

Private Sub cmdFixAll_Click()

Set oActDoc = Nothing
IntCATIA
If oActDoc Is Nothing Then
MsgBox "只能在装配设计工作台下使用！"
Exit Sub
End If
Debug.Print TypeName(oActDoc)

If TypeName(oActDoc) <> "ProductDocument" Then
MsgBox "只能在装配设计工作台下使用！"
Exit Sub
End If

Dim oTopProd As ProductDocument
Dim oCurrentProd As Object
'Dim oVisProp As Object ' VisPropertySet

Set oSel = oActDoc.Selection
'Set oVisProp = oSel.VisProperties

Set oCurrentProd = oActDoc.Product

Set oList = CreateObject("Scripting.dictionary")

Call FixSingleLevel(oCurrentProd)    'Call the subroutine, it is a recursive loop
oList.RemoveAll
CATIA.StatusBar = "Macro Done"
MsgBox "所有零件已经固定！"
'Unload Me

End Sub

Private Sub FixSingleLevel(ByRef oCurrentProd As Object)

On Error Resume Next

'More declarations
Dim ItemToFix As Product
Dim iProdCount As Integer
Dim i As Integer
Dim j As Integer
Dim oConstraints 'As Constraints
Dim oReference 'As Reference
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
    'Debug.Print n & "/" & oConstraints.Item(n).Type & "/" & oConstraints.Item(n).ReferenceType
    If (oConstraints.Item(n).Type = 0) And (oConstraints.Item(n).ReferenceType = 1) Then
    oConstraints.Remove (n)
    End If
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
    
    oSel.Add constraint1      'Set visibility to hidden
    Set oVisProp = oSel.VisProperties
    oVisProp.SetShow 1 'catVisPropertyNoShowAttr
    oSel.Clear

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
Private Sub UnFixSingleLevel(ByRef oCurrentProd As Object)

On Error Resume Next

'More declarations
Dim ItemToFix As Product
Dim iProdCount As Integer
Dim i As Integer
Dim j As Integer
Dim oConstraints 'As Constraints
Dim oReference 'As Reference
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
    If (oConstraints.Item(n).Type = 0) And (oConstraints.Item(n).ReferenceType = 1) Then
    oConstraints.Remove (n)
    End If
    n = n - 1
Next


For i = 1 To iProdCount                             'Cycle through the assembly's children

    Set ItemToFix = oCurrentProd.Products.Item(i)

'CreateReference:
'
'    sItemName = ItemToFix.Name
'    CATIA.StatusBar = "Working On " & oCurrentProd.Name & " / " & sItemName
'
'    Set oReference = oCurrentProd.CreateReferenceFromName(sItemName & "/!" & "/")
'
'    Set constraint1 = oConstraints.AddMonoEltCst(0, oReference)
'    constraint1.ReferenceType = 1 'catCstRefTypeFixInSpace
'
'    oSel.Add constraint1      'Set visibility to hidden
'    Set oVisProp = oSel.VisProperties
'    oVisProp.SetShow 1 'catVisPropertyNoShowAttr
'    oSel.Clear

RecursionCall:
    If ItemToFix.Products.Count <> 0 Then        'Recursive Call
        If oList.exists(ItemToFix.PartNumber) Then GoTo Finish
        
        If ItemToFix.PartNumber = ItemToFix.ReferenceProduct.Parent.Product.PartNumber Then oList.Add ItemToFix.PartNumber, 1
        Call UnFixSingleLevel(ItemToFix)
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

Private Sub cmdNoFix_Click()
Set oActDoc = Nothing
IntCATIA
If oActDoc Is Nothing Then
MsgBox "只能在装配设计工作台下使用！"
Exit Sub
End If
Debug.Print TypeName(oActDoc)

If TypeName(oActDoc) <> "ProductDocument" Then
MsgBox "只能在装配设计工作台下使用！"
Exit Sub
End If

Dim oTopProd As ProductDocument
Dim oCurrentProd As Object
'Dim oVisProp As Object ' VisPropertySet

'Set oSel = oActDoc.Selection
'Set oVisProp = oSel.VisProperties

Set oCurrentProd = oActDoc.Product

Set oList = CreateObject("Scripting.dictionary")

Call UnFixSingleLevel(oCurrentProd)    'Call the subroutine, it is a recursive loop
oList.RemoveAll
CATIA.StatusBar = "Macro Done"
MsgBox "已取消所有零件固定！"
'Unload Me
End Sub

Private Sub RemoveFailueConst(ByRef oCurrentProd As Object)

On Error Resume Next

'More declarations
Dim ItemToFix As Product
Dim iProdCount As Integer
Dim i As Integer
Dim j As Integer
Dim oConstraints 'As Constraints
Dim oReference 'As Reference
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
    If (oConstraints.Item(n).Status <> 0) Then
    oConstraints.Remove (n)
    End If
    n = n - 1
Next


For i = 1 To iProdCount                             'Cycle through the assembly's children

    Set ItemToFix = oCurrentProd.Products.Item(i)

RecursionCall:
    If ItemToFix.Products.Count <> 0 Then        'Recursive Call
        If oList.exists(ItemToFix.PartNumber) Then GoTo Finish
        
        If ItemToFix.PartNumber = ItemToFix.ReferenceProduct.Parent.Product.PartNumber Then oList.Add ItemToFix.PartNumber, 1
        Call RemoveFailueConst(ItemToFix)
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
        CATIA.StatusBar = "Remove Constraint Aborted"
        End
       End If
       Else: Resume RecursionCall
      
    End If
'*****End of Error Handling

End2:
Resume

End1:
End Sub

Private Sub cmdRemoveFailure_Click()
Set oActDoc = Nothing
IntCATIA
If oActDoc Is Nothing Then
MsgBox "只能在装配设计工作台下使用！"
Exit Sub
End If
Debug.Print TypeName(oActDoc)

If TypeName(oActDoc) <> "ProductDocument" Then
MsgBox "只能在装配设计工作台下使用！"
Exit Sub
End If

Dim oTopProd As ProductDocument
Dim oCurrentProd As Object

Set oSel = oActDoc.Selection

Set oCurrentProd = oActDoc.Product

Set oList = CreateObject("Scripting.dictionary")

Call RemoveFailueConst(oCurrentProd)    'Call the subroutine, it is a recursive loop
oList.RemoveAll
CATIA.StatusBar = "Macro Done"
MsgBox "已移除错误约束！"
'Unload Me
End Sub
