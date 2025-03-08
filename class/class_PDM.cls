VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_PDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'version 2025.0303
' Class module for product operations
' Public Class Product_Module
' #Folder:class 我现在希望class PDM类，实现以下多个功能：

' 1  初始化方法， 获取g_rootPrd为catia当前打开的根产品
' 2  catSel1方法  带提示让用户选择单个产品\part
' 3  catSelx方法 带提示让用户选择多个产品\part
' 4  infoPrd方法 获取传递的oPrd的一系列属性
' 5  getAtt方法  获取 属性或属性值，不存在时返回nothing 和“__”
' 6  meQty 方法  获取产品在当前子总成下的数量
' 7  recurPrd方法  调用infoPrd方法获取根产品及所有子产品的属性，返回为数组，数组第一列为编号，数组第二列为装配层级


Option Explicit
Public Prd2wt
Public bomdata As Variant
Public counter As Long
Private i  As Integer
Private bdict  As Object
Private oSel
Private imsg
Private att(10)
Private aType(10)

Private Sub Class_Initialize()
    iniarr
   On Error Resume Next
        If CATIA.Documents.Count = 0 Then
            Err.Raise 1001, , "未检测到打开的CATIA文档"
        End If
        Set rootPrd = CATIA.ActiveDocument.Product
        If Not rootPrd Is Nothing Then
         rootPrd.ApplyWorkMode (3)
     End If
        Set oSel = CATIA.ActiveDocument.Selection: oSel.Clear
    Exit Sub
    
    If Err.Number <> 0 Then
    MsgBox "初始化失败：" & Err.Description, vbCritical
    Err.Clear

    End If
   
End Sub

'Public Sub setmode(oprd)
'
'oprd.ApplyWorkMode (3)
'
'End Sub

Sub iniarr()
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
Public Function catchPrd()    '获取要读取或修改的  全局变量 Prd2wt 定义
     imsg = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
        Select Case imsg
            Case 7: Set catchPrd = rootPrd '===选择“否”====
            Case 2: Exit Function '===选择“取消”====
            Case 6  '===选择“是”,进行产品选择====
                On Error Resume Next
                    Set catchPrd = selPrd()
                If Err.Number <> 0 Then
                    Err.Clear
                    Exit Function
                End If
        End Select
        If Not catchPrd Is Nothing Then
           imsg = "你选择的产品是" & catchPrd.PartNumber & "是否继续"
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
        Dim arr(1 To 9)
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
                arr(9) = count_me(oPrd)
        infoPrd = arr()
    End Function
    
Private Function count_me(oPrd)  '获取兄弟字典
     Dim i, oDict, QTy, pn
         QTy = 1
     On Error Resume Next
     If TypeOf oPrd.Parent Is Products Then    '若有父级产品'获取兄弟字典
               Dim oParent: Set oParent = oPrd.Parent.Parent
         
              Set oDict = CreateObject("Scripting.Dictionary")
              For i = 1 To oParent.Products.Count
                     pn = oParent.Products.Item(i).PartNumber
                     If oDict.Exists(pn) = True Then
                         oDict(pn) = oDict(pn) + 1
                     Else
                         oDict(pn) = 1
                     End If
                 Next
        QTy = oDict(oPrd.PartNumber)       '获取oprd数量
     End If
     If Error.Number <> 0 Then
          QTy = 1
     End If
    count_me = QTy
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
    Dim i, pn, QTy
    On Error Resume Next      ' 检查 oPrd 是否为根产品
        Set parentPrd = oPrd.Parent.Parent
        If (parentPrd Is Nothing) Or (TypeName(parentPrd) = "ProductDocument") Then
            QTy = 1 ' 是根产品，数量设为 1
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
            QTy = oDict(oPrd.PartNumber) ' 从兄弟字典中获取 oPrd 的数量
        End If
    On Error GoTo 0
    meQty = QTy
End Function
' 全局变量声明
Public Function recurPrd(oPrd, LV)

     If counter = 0 Then
          ReDim bomdata(1 To 1000, 1 To 11) ' 扩展为11列：2列原有数据 + 9列产品属性
     '     If IsEmpty(bomdata) Then
    End If
    counter = counter + 1
    bomdata(counter, 1) = counter
    bomdata(counter, 2) = LV
    
    Dim prdInfo, j
     prdInfo = infoPrd(oPrd)
     For j = 1 To 9
         bomdata(counter, j + 2) = prdInfo(j)
     Next j
    
    Dim children As Products
    Set children = oPrd.Products
    If children.Count > 0 Then
        Dim i As Integer
        For i = 1 To children.Count
            recurPrd children.Item(i), LV + 1
        Next
    End If
      If counter > 0 Then
        Dim finalData() As Variant
        ReDim finalData(1 To counter, 1 To 11)
        finalData = Application.Index(bomdata, Application.Evaluate("ROW(1:" & counter & ")"), Application.Evaluate("ROW(1:11)"))
        recurPrd = finalData
    End If
End Function
