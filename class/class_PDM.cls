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
' 2  catSel1方法  带提示让用户选择单个产品\Part
' 3  catSelx方法 带提示让用户选择多个产品\Part
' 4  infoPrd方法 获取传递的oPrd的一系列属性
' 5  getPara方法  获取 属性或属性值，不存在时返回nothing 和“__”
' 6  meQty 方法  获取产品在当前子总成下的数量
' 7  recurPrd方法  调用infoPrd方法获取根产品及所有子产品的属性，返回为数组，数组第一列为编号，数组第二列为装配层级
' 对于产品发布谁？  质量mass来自usrp
' 对于零件发布谁？ usrp.mass part_info.thickness part_info.density usrp.material

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
    iset="Part_info"
    eklstr="let lst(list) set lst=Part_info\iBodys  let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i))  i=i+1}"
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
                arr(5) = getPara("Mass", colls)(1)
                arr(6) = getPara("Material", colls)(1)
                arr(7) = getPara("Thickness", colls)(1)
        On Error Resume Next
           Set colls = oPrd.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("part_info").DirectParameters
                arr(8) = getPara("Density", colls)(1)
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
                case 1: refprd.Partnumber=oData(i)
                case 2: refprd.nomenclature=oData(i)
                case 3: refprd.definition=oData(i)
                case 4: oprd.name=oData(i)                
                case 5 to 7
                    set colls= refprd.userrefproperties
                    colls.item(att(i-4)).value = oData(i)
                case 8
                    Set colls = refprd.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("Part_info").DirectParameters
                    colls.item(att(i-4)).value = oData(i)                
            end select
            on error go to 0
        end if
    next
end sub

public sub initPrd(oprd)
    Dim refprd: Set refprd = oPrd.ReferenceProduct   
    Dim NT As Variant
        NT = Array( _
            Array("Mass", "Mass"), _
            Array("Material", "String"), _
            Array("Thickness", "Length"), _
            Array("Density", "Density") _
        )
    Dim usrp(0 To 3)
    Set colls = refprd.UserRefProperties
    For i = 0 To 3
        Set usrp(i) = New Class_para
        usrp(i).SetNT NT(i)(0), NT(i)(1)
        checkPara usrp(i), colls
    Next
    on error resume next
         Dim oPrt: Set oPrt = refprd.Parent.Part
    If Err.Number = 0 Then
           Call iniPrt(oPrd, att_usp_Names())
        Else
           Err.Clear
    end if

Private sub iniPrt(oPrd)
    Dim refprd: Set refprd = oPrd.ReferenceProduct
    Dim oPrt: Set oPrt = refprd.Parent.Part
    '============创建参数集合=================
        Set infoset = New Class_para
        infoset.SetNT "Part_info", "ParameterSet"
        Set colls = oPrt.Parameters.RootParameterSet.ParameterSets
        checkPara infoset, colls
        Set colls = infoset.obj.DirectParameters
'============创建part_info内参数=================
    Dim iBodys
    Set iBodys = New Class_para
    iBodys.SetNT "iBodys", "list"
    checkPara iBodys, colls
    If getPara(oPrt.mainbody, iBodys.obj.valuelist)(0) Is Nothing Then
            iBodys.obj.valuelist.Add oPrt.mainbody
    End If
    
    Dim piset(0 To 2)
     PNT = Array( _
        Array("sumVol", "Volume"), _
        Array("Thickness", "Length"), _
        Array("Density", "Density") _
    )
    For i = 0 To 2
        Set piset(i) = New Class_para
        piset(i).SetNT PNT(i)(0), PNT(i)(1)
        checkPara piset(i), colls
    Next
    
Set colls = oPrt.relations
Dim oRule
Set oRule = New Class_para
       '---创建EKL
oRule.Reset
oRule.SetNT "sum_all_Vol", "Program"
oRule.str = eklstr
oRule.Desc = "汇总体积"
checkPara oRule, colls
        '---创建link_mass
oRule.Reset
oRule.SetNT "link_mass", "Formula"
oRule.Desc = "汇总体积"
oRule.Str = "Part_info\sumVol *Part_info\Density"
Set oRule.Target = refprd.UserRefProperties.item("Mass")
checkPara oRule, colls

        '---创建link_thickness
oRule.Reset
oRule.SetNT "link_thickness", "Formula"
oRule.Desc = "链接厚度"
oRule.Str = "Part_info\Thickness"
Set oRule.Target = refprd.UserRefProperties.item("Thickness")
checkPara oRule, colls

Private Function checkPara(thispara, colls)
    If  getPara(thispara, colls)(0) Is Nothing Then
        Debug.Print "需要创建" & thispara.Name
        paraCreatobj thispara, colls   '创建参数和公式时已经是默认值
        Debug.Print "已经创建" & thispara.Name
    else
       Set thispara.obj = getPara(thispara, colls)(0) '已有参数和公式，校验默认值
       Debug.Print "不需要创建" & thispara.obj.Name
End Function

Private Function getPara(thispara, collection)
    Dim arr(1) ' 正确声明数组
    On Error Resume Next
        Set arr(0) = collection.item(thispara.Name)
        If Err.Number = 0 Then ' 检查是否成功获取对象
            arr(1) = arr(0).Value
            getPara = arr
        Else
            getPara = Array(Nothing, "__")
        End If
    On Error GoTo 0
End Function

Private Function paraCreatobj(thispara, colls)
        Select Case thispara.iType
            Case "ParameterSet"
                Set thispara.obj = colls.CreateSet(thispara.Name)
            Case "list"
                Set thispara.obj = colls.CreateList(thispara.Name)
            Case "Mass", "Density", "Length", "Volume" '所有dimension参数
                Set thispara.obj = colls.CreateDimension(thispara.Name, thispara.iType, thispara.Str)
            Case "String"
                Set thispara.obj = colls.CreateString(thispara.Name, thispara.Str)
            Case "Formula"
                Set thispara.obj = colls.CreateFormula(thispara.Name, thispara.Desc, thispara.Target, thispara.Str)
            Case "Program"
                Set thispara.obj = colls.CreateProgram(thispara.Name, thispara.Desc, thispara.Str)
    End Select
End Function





  
