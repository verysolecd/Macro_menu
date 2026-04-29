VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TitleBlockProperties 
   Caption         =   "Properties"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "TitleBlockProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TitleBlockProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOK_Click()

    ModScaniaPro CATIA.ActiveDocument.Product
    Unload Me

End Sub

Sub ModScaniaPro(oPrd As Object)
Dim crl
Dim j As Integer
Dim oUserPara As Object 'Parameters
Set oUserPara = oPrd.UserRefProperties
    For j = 1 To oUserPara.Count
    oUserPara.Remove (1)
    Next
    
    For Each crl In TitleBlockProperties.Controls
        If TypeName(crl) = "TextBox" Then
        '需检查不属于固有属性 "PartNumber", "Revision", "Definition", "Nomenclature", "Source", "DescriptionRef"
        Select Case crl.Name
               Case "PartNumber"
                    oPrd.PartNumber = crl.Text
               Case "Revision"
                    oPrd.Revision = crl.Text
               Case "Definition"
                    oPrd.Definition = crl.Text
               Case "Nomenclature"
                     oPrd.Nomenclature = crl.Text
               Case "DescriptionRef"
                    oPrd.DescriptionRef = crl.Text
               Case Else
                Set strParam = oUserPara.CreateString(crl.Name, crl.Text)
        End Select
        End If
    Next
End Sub

