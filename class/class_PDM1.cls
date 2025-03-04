VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_PDM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'version 2025.0303
' Class module for product operations
' Public Class Product_Module
'策划功能：

Option Explicit
Public Prd2wt
Public all_rows_data As Variant
Public counter As Long
Private i  As Integer
Private bdict  As Object
Private oSel
Private imsg
Private att(10)
Private aType(10)
Sub init()
    On Error Resume Next
        Set rootPrd = catia.ActiveDocument.Product
        Set oSel = catia.ActiveDocument.Selection
        oSel.Clear
        If Err.Number <> 0 Then
            MsgBox "请打开产品文档", vbCritical
            Exit Sub
        End If
    On Error GoTo 0
    ' 初始化属性定义
    att(1) = "iMaterial"
    att(2) = "iDensity"
    att(3) = "iMass"
    att(4) = "iThickness"
    aType(1) = "String"
    aType(2) = "Density"
    aType(3) = "Mass"
    aType(4) = "Length"
End Sub
Public Function catSel1(prompt As String) As Object
    oSel.Clear
    Dim iType(0)
    iType(0) = "Product"
    MsgBox prompt
    If oSel.SelectElement2(iType, prompt, False) = "Normal" Then
        If oSel.Count = 1 Then
            Set catSel1 = oSel.Item(1).LeafProduct
        End If
    End If
    oSel.Clear
End Function
Public Function catSelx(prompt As String) As collection
    Set catSelx = New collection
    oSel.Clear
    Dim iType(0)
    iType(0) = "Product"
    If oSel.SelectElement3(iType, prompt, True, 2, True) = "Normal" Then
        Dim i As Integer
        For i = 1 To oSel.Count
            catSelx.Add oSel.Item(i).LeafProduct
        Next
    End If
    oSel.Clear
End Function
Public Function whois2rv()    '获取要读取或修改的  全局变量 Prd2wt 定义
     imsg = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
        Select Case imsg
            Case 7: Set whois2rv = rootPrd '===选择“否”====
            Case 2: Exit Function '===选择“取消”====
            Case 6  '===选择“是”,进行产品选择====
                On Error Resume Next
                    Set whois2rv = selPrd()
                If Err.Number <> 0 Then
                    Err.Clear
                    Exit Function
                End If
        End Select
        If Not whois2rv Is Nothing Then
           imsg = "你选择的产品是" & whois2rv.PartNumber & "是否继续"
          MsgBox imsg
        End If
End Function
Public Function selPrd()
   Dim status, iType(0)
'   catia.ActiveWindow.WindowState = 0
    oSel.Clear
    iType(0) = "Product"
    If oSel.Count2 = 0 Then
        MsgBox "请选择要读取的产品"
        status = oSel.SelectElement2(iType, "请选择要读取的产品", False)
        'status = oSel.SelectElement3(iType, "请选择要读取的产品", True, 2, False)
    End If
    If status = "Cancel" Then
        Exit Function
    End If
    If status = "Normal" And oSel.Count2 = 1 Then
            Set selPrd = oSel.Item(1).LeafProduct.ReferenceProduct
            MsgBox "你选择的产品是" & selPrd.PartNumber
            oSel.Clear
    Else
        MsgBox "请只选择一个产品"
        Exit Function
    End If
End Function

    ' Calculate the total mass of a product and its sub-products
    Public Function Assmass(oPrd)
        Dim i
        Dim total, children: total = 0
        Set children = oPrd.Products
        If oPrd.Products.Count > 0 Then
            For i = 1 To children.Count
                Assmass = Assmass + Assmass(children.Item(i))
                total = total + children.Item(i).ReferenceProduct.UserRefProperties.Item("iMass").Value
            Next
            oPrd.ReferenceProduct.UserRefProperties.Item("iMass").Value = total
        Else
            total = oPrd.ReferenceProduct.UserRefProperties.Item("iMass").Value
        End If
        Assmass = total
    End Function
    ' Generate product information
    Public Function infoPrd(oPrd)
        Dim arr(10)
            With oPrd.ReferenceProduct
                arr(1) = .PartNumber
                arr(2) = .Nomenclature
                arr(3) = .Definition
                arr(4) = oPrd.Name
            End With
        Dim usrp
           Set usrp = oPrd.ReferenceProduct.UserRefProperties
                arr(5) = getAtt("iMass", usrp)(1)
                arr(6) = getAtt("iMaterial", usrp)(1)
                arr(7) = getAtt("iThickness", usrp)(1)
        On Error Resume Next
           Set usrp = oPrd.ReferenceProduct.Parent.part.parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                arr(8) = getAtt("iDensity", usrp)(1)
            If Error.Number <> 0 Then
                arr(8) = "__"
            End If
        On Error GoTo 0
                arr(9) = meQty(oPrd)
        infoPrd = arr()
    End Function
Private Function getAtt(itemName, collection)
    Dim arr(1) ' 正确声明数组
    On Error Resume Next
        Set arr(0) = collection.Item(itemName)
        If Err.Number = 0 Then ' 检查是否成功获取对象
            arr(1) = arr(0).Value
            getAtt = arr
        Else
            getAtt = Array(Nothing, "__")
        End If
    On Error GoTo 0
End Function
Function meQty(oPrd)
    Dim parentPrd, oDict
    Dim i, pn, Qty
    On Error Resume Next      ' 检查 oPrd 是否为根产品
        Set parentPrd = oPrd.Parent.Parent
        If (parentPrd Is Nothing) Or (TypeName(parentPrd) = "ProductDocument") Then
            Qty = 1 ' 是根产品，数量设为 1
        Else
            Set oDict = CreateObject("Scripting.Dictionary")        ' 不是根产品，获取兄弟字典
            For i = 1 To parentPrd.Products.Count
                pn = parentPrd.Products.Item(i).PartNumber
                If oDict.Exists(pn) Then
                    oDict(pn) = oDict(pn) + 1
                Else
                    oDict(pn) = 1
                End If
            Next
            Qty = oDict(oPrd.PartNumber) ' 从兄弟字典中获取 oPrd 的数量
        End If
    On Error GoTo 0
    meQty = Qty
End Function
' 全局变量声明
Public Function recurPrd(ByVal oPrd As Object, ByVal LV As Integer) As Variant
    If all_rows_data Is Nothing Then
        ReDim all_rows_data(1 To 1, 1 To 2)
        counter = 0
    End If
    counter = counter + 1
    ReDim Preserve all_rows_data(1 To counter, 1 To 2)
    all_rows_data(counter, 1) = counter
    all_rows_data(counter, 2) = LV
    Dim children As Products
    Set children = oPrd.Products
    If children.Count > 0 Then
        Dim i As Integer
        For i = 1 To children.Count
            recurPrd children.Item(i), LV + 1
        Next
    End If
    recurPrd = all_rows_data
End Function
