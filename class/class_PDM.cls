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
Private att(1 to 10)
Private aType(1 to 10)

Private Sub Class_Initialize()
    iniarr
    On Error GoTo ErrorHandler
        If CATIA.Documents.Count = 0 Then
            Err.Raise 1001, , "未检测到打开的CATIA文档"
            Exit Sub
        End If
        Set rootPrd = CATIA.Activedocument.product
        
        On Error GoTo 0
        If Not rootPrd Is Nothing Then
         rootPrd.ApplyWorkMode (3)
          End If
        Set oSel = CATIA.Activedocument.Selection: oSel.Clear
    Exit Sub    
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "初始化失败：" & Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
End Sub
Sub iniarr()   
    att(1) = "iMass"
    att(2) = "iMaterial"
    att(3) = "iThickness"
    att(4) = "iDensity"
    aType(1) = "Mass"
    aType(2) = "String"    
    aType(3) = "Length"
    aType(4) = "Density"
    iset="part_info"
    eklstr="let lst(list) set lst=part_info\iBodys  let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i))  i=i+1}"
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
Public Function defgprd()    '获取要读取或修改的  全局变量 Prd2wt 定义
     imsg = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
     Dim oPrd
        Select Case imsg
            Case 7 '===选择“否”====
            Set rootPrd = CATIA.Activedocument.product
            Set oPrd = rootPrd            
            Case 2: Exit Function '===选择“取消”====
            Case 6  '===选择“是”,进行产品选择====
                On Error Resume Next
                    Set oPrd = selPrd()
                If Err.Number <> 0 Then
                    Err.Clear
                    Exit Function
                End If
        End Select
        If Not oPrd Is Nothing Then
           imsg = "你选择的产品是" & oPrd.PartNumber & "是否继续"
          MsgBox imsg
        End If
        Set defgprd = oPrd
End Function

Public Sub freePrd()
     Set gprd = Nothing
End Sub
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
        Dim colls
           Set colls = oPrd.ReferenceProduct.UserRefProperties
                arr(5) = getAtt("iMass", colls)(1)
                arr(6) = getAtt("iMaterial", colls)(1)
                arr(7) = getAtt("iThickness", colls)(1)
        On Error Resume Next
           Set colls = oPrd.ReferenceProduct.Parent.part.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                arr(8) = getAtt("iDensity", colls)(1)
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
    recurPrd = bomdata
End Function

public modatt(oprd, oData)
    dim refprd : set refprd = oprd.referenceproduct
    dim colls
    dim currAtt(1 to 9)
    set currAtt() =infoPrd(oprd)
    for i=1 to 8
        if oData(i) <> "" and arr(i) <> currAtt(i) then
            on error resume next
            select case 
                case 1: refprd.partnumber=oData(i)
                case 2: refprd.nomenclature=oData(i)
                case 3: refprd.definition=oData(i)
                case 4: oprd.name=oData(i)                
                case 5 to 7
                    set colls= refprd.userrefproperties
                    colls.item(att(i-4)).value = oData(i)
                case 8
                    Set colls = refprd.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("part_info").DirectParameters
                    colls.item(att(i-4)).value = oData(i)                
            end select
            on error go to 0
        end if
    next
end sub

public sub initprd(oprd)
    dim refprd : set refprd = oprd.referenceproduct
    dim colls
    dim attObj(1 to 4)
    set colls = refprd.userrefproperties
    for i = 1 to 4
        if getAtt(att(i),colls)(0) is nothing then                
            select case
                case 1,3        
                    set attObj = colls.CreateDimension(att(i),aType(i),0#)       
                Case 2
                    set attobj(i)=colls.createstring(att(i),"tbd")
            end select
        else
            set attObj(i)=getAtt(att(i),colls)(0)
    Next
     Set colls = Nothing
    '=================进行发布======================
    Dim oprt 
    On Error Resume Next
        Set oPrt = refprd.Parent.Part
        Err.Clear
    On Error GoTo 0
    If Not oPrt Is Nothing Then
        Call iniPrt(oPrd, att_usp_Names())
    Else         
        Set colls = refprd.Publications       
        If getAtt(att(i),colls)(0) is nothing Then
            Dim oref, oPub
            i = 1
            Set oref = refprd.CreateReferenceFromName(attObj(i).Name)
            Set oPub = colls.Add(att(i)) ' 添加发布元素名称
            colls.SetDirect att(i), oref ' 设置发布元素参考
            Set colls = Nothing
        End If
    End If
    On Error GoTo 0
    Set colls = Nothing
End Sub

private sub initprt(oprd)
    Dim refprd：Set refprd = oPrd.ReferenceProduct
    dim oPrt：Set oPrt = refprd.Parent.Part
 '============创建参数集合=================
    Dim colls
    On Error Resume Next
        Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item(iset).DirectParameters
        If colls is nothing then
            Err.Clear
            Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.CreateSet(iset)
            Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item(iset).DirectParameters
        End If
    On Error GoTo 0
    '创建list，将mainbody加入list
    if gettAtt("ibodys",colls)(0) is nothing then
        Dim lst
        Set lst = colls.CreateList("ibodys")
    else
        Set lst = gettAtt("ibodys",colls)(0)
    end if
        if getAtt(oprt.mainbody.name,lst.valuelist)(0) is nothing then
            lst.valuelist.Add oprt.mainbody
        end if
        set lst = nothing
    '============创建参数和发布=================
    dim attObj
    For i = 1 to 4
        if i<>2 Then
            if getAtt(att(i),colls)(0) is nothing then     
                set attObj = colls.CreateDimension(att(i),aType(i),0#)       
            else          
                set attObj(i)=getAtt(att(i),colls)(0)         
            end if
        end if 
    next
        '=================进行发布======================    
        Dim Pubs
        Set pubs = refprd.Publications
            for i=1 to 4
                if getAtt(att(i),pubs)(0) is nothing then
                    Dim oref, oPub
                    Select Case i             
                        Case 2
                            Set attObj(i) = refprd.UserRefProperties.Item(att(i))                    
                    End Select
                Set oref = refprd.CreateReferenceFromName(attObj(i).Name)
                    Set oPub = pubs.Add(att(i)) ' 添加发布
                    pubs.SetDirect att(i), oref ' 设置发布元素
                End If
            next
    '================创建ekl==================== 
   set colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item(iset).DirectParameters
    if gettAtt("eklname",colls)(0) is nothing then
        Dim oRule
        dim oRL:set oRL=oprt.relations
        dim eklname, ekldesc
        eklname="cal_volume"
        ekldesc="calculate volume"
        Set oRule =oRL.CreateProgram(eklname, ekldesc, eklstr)
    else
        if  gettAtt("ekl",colls)(1)<>eklstr then
            gettAtt("ekl",colls)(0).modify eklstr
        end if
    end if