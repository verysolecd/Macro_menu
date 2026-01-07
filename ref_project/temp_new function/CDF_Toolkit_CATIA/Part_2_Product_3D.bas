Attribute VB_Name = "Part_2_Product_3D"
Option Explicit



Sub CATMain()

       
Part2Prod.Show vbModeless

End Sub

Sub Part2Product(oDocPart)
If TypeName(oDocPart) <> "PartDocument" Then
MsgBox "当前文档不是CATIA零件(*.CATPart)!" & vbCrLf & "请打开CATIA零件!"
Exit Sub
End If


Dim oDocPart1, oPart, oBodies, oBody, n, sn, oBody1, oSelection

Set oBodies = oDocPart.Part.Bodies
If oBodies.Count < 2 Then
    If MsgBox("当前零件只有一个实体, 确实需要再生成一个CATIA产品么? 该产品将只包含一个零件?", vbYesNo, "请确认") = vbNo Then
    Exit Sub
    End If
End If
Dim PreFix As String
Dim SufFix As String
PreFix = InputBox("部件名称采用 前缀+序列号+后缀 的形式" & vbCrLf & "请输入前缀(勿用中文)，不需要前缀则按取消（留空）: ", "CDF_Toolkit - Part2Product", "CDF-1688-")
SufFix = InputBox("请输入后缀(勿用中文)，不需要则按取消（留空）: ", "CDF_Toolkit - Part2Product")

'************************
Dim settingControllers1 As Object
Dim partInfrastructureSettingAtt1 As Object
Dim boolean9 As Boolean

Set settingControllers1 = CATIA.SettingControllers

Set partInfrastructureSettingAtt1 = settingControllers1.Item("CATMmuPartInfrastructureSettingCtrl")
boolean9 = partInfrastructureSettingAtt1.HybridDesignMode
'Debug.Print partInfrastructureSettingAtt1.HybridDesignMode
If partInfrastructureSettingAtt1.HybridDesignMode = 1 Then
   partInfrastructureSettingAtt1.HybridDesignMode = 0
'Debug.Print partInfrastructureSettingAtt1.HybridDesignMode
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
End If

'************************

Dim oDocProd As ProductDocument
Set oDocProd = CATIA.Documents.Add("Product")
oDocProd.Product.PartNumber = oDocPart.Product.PartNumber

Dim oProds As Products
Set oProds = oDocProd.Product.Products

On Error Resume Next
For n = 1 To oBodies.Count
    Set oBody = oBodies.Item(n)
    Set oSelection = oDocPart.Selection
    oSelection.Clear
    If oBodies.Item(n).InBooleanOperation = False Then
    oSelection.Add oBody
    'MsgBox oBody.Name
    oSelection.Copy
 
    
    Dim sDocPart1Nomenclature
    sDocPart1Nomenclature = Replace(Replace(Replace(oBody.Name, "\", "_"), "#", "_"), ".", "_")
    'MsgBox sDocPart1PN

    If n < 10 Then
    sn = "0" & CStr(n)
    Else
    sn = CStr(n)
    End If
    
    oProds.AddNewComponent "Part", PreFix & sn & SufFix
    Set oDocPart1 = CATIA.Documents.Item(PreFix & sn & SufFix & ".CATPart")
 

    
    oSelection.Add oDocPart1.Part
   
    oSelection.PasteSpecial "CATPrtResultWithOutLink"
    
    oDocPart1.Part.MainBody = oDocPart1.Part.Bodies.Item(oBody.Name)
    'MsgBox oDocPart1.Name
    
    oDocPart1.Product.PartNumber = Left(oDocPart1.Name, Len(oDocPart1.Name) - 8)
    oDocPart1.Product.Nomenclature = sDocPart1Nomenclature
    oDocPart1.Product.DescriptionInst = sDocPart1Nomenclature
    oDocPart1.Part.Update
    oProds.Item(n).Name = sDocPart1Nomenclature
    End If
    
Next

On Error GoTo 0
oDocProd.Activate
Set oBodies = Nothing
Set oProds = Nothing
Set oSelection = Nothing

partInfrastructureSettingAtt1.HybridDesignMode = boolean9
'Debug.Print partInfrastructureSettingAtt1.HybridDesignMode
Set settingControllers1 = Nothing
Set partInfrastructureSettingAtt1 = Nothing
End Sub
