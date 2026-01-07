Attribute VB_Name = "Export_Stp2_2D3D"
Sub CATMain()
       
IntCATIA
ExportStp3 oActDoc
End Sub
Sub ExportStp3(oDocStp)
On Error Resume Next
CATIA.DisplayFileAlerts = False
'*****************************
Dim settingControllers1 As Object
Dim stepSettingAtt1 As Object
Dim short1

Set settingControllers1 = CATIA.SettingControllers
Set stepSettingAtt1 = settingControllers1.Item("CATSdeStepSettingCtrl")
short1 = stepSettingAtt1.AttAP
'Debug.Print stepSettingAtt1.AttAP
If short1 <> 2 Then
stepSettingAtt1.AttAP = 2
'Debug.Print stepSettingAtt1.AttAP
End If
Dim sName1 As String
sName1 = oDocStp.Path & "\" & Replace(oDocStp.Name, ".", "_")

If (TypeName(oDocStp) = "PartDocument") Then
oDocStp.ExportData sName1, "stp"
End If

If TypeName(oDocStp) = "ProductDocument" Then
oDocStp.Product.ApplyWorkMode DESIGN_MODE
oDocStp.ExportData sName1, "stp"
ExportStp2 oDocStp.Product
End If


stepSettingAtt1.AttAP = short1
'Debug.Print stepSettingAtt1.AttAP
Set settingControllers1 = Nothing
Set stepSettingAtt1 = Nothing
'On Error GoTo 0
CATIA.DisplayFileAlerts = True
MsgBox "²Ù×÷Íê³É£¡"
End Sub

Sub ExportStp2(oProd2)
On Error Resume Next
Dim sName2 As String
If ProductIsComponent(oProd2) = False Then
    Dim objDoc2 As Object
    oProd2.ApplyWorkMode DESIGN_MODE
    Set objDoc2 = oProd2.ReferenceProduct.Parent
    sName2 = objDoc2.Path & "\" & Replace(objDoc2.Name, ".", "_")
    objDoc2.ExportData sName2, "stp"
End If
If oProd2.Products.Count <> 0 Then
'    Dim oProdx As Object
    Dim mi As Integer
    For mi = 1 To oProd2.Products.Count
    Call ExportStp2(oProd2.Products.Item(mi))
    Next
End If

End Sub

Function ProductIsComponent(iProduct) As Boolean
Dim objDoc As Object
Dim objParentDoc As Object
ProductIsComponent = False
On Error Resume Next
Set objDoc = iProduct.ReferenceProduct.Parent
Set objParentDoc = iProduct.Parent.Parent.ReferenceProduct.Parent
ProductIsComponent = (objDoc Is objParentDoc)

Set objDoc = Nothing
Set objParentDoc = Nothing
Err.Clear

End Function
