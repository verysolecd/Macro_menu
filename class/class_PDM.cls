VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_PDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'我现在希望class PDM类，实现以下多个功能：
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
Public bomdata As Variant
Private oSel
Private imsg
Private Att(1 To 10)
Private Const xx = "测试成功"
Private Const xy = "测试失败"
Private Const eklname = "sumVol"
Private Const ekldesc = "sum of vol of bodylist"
Private Const eklstr = "let lst(list) set lst=Part_info\iBodys   let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol =V"  '使用Const关键字定义常量
Private Sub Class_Initialize()
    iniarr
        On Error GoTo ErrorHandler
        If CATIA.Documents.Count = 0 Then
            Err.Raise 1001, , "未检测到打开的CATIA文档"
            Exit Sub
        End If
        Set rootPrd = CATIA.ActiveDocument.Product
        On Error GoTo 0
        If Not rootPrd Is Nothing Then
         rootPrd.ApplyWorkMode (3)
        End If
        Set oSel = CATIA.ActiveDocument.Selection: oSel.Clear
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "CATIA 程序错误：" & Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
End Sub
Sub iniarr()
    Att(1) = "Mass"
    Att(2) = "Material"
    Att(3) = "Thickness"
    Att(4) = "Density"
End Sub
    ' Calculate the total mass of a product and its sub-products
    Public Function Assmass(oPrd)
        Dim i
        Dim total, children: total = 0
        Set children = oPrd.Products
        If oPrd.Products.Count > 0 Then
            For i = 1 To children.Count
                Assmass = Assmass + Assmass(children.item(i))
                total = total + children.item(i).ReferenceProduct.UserRefProperties.item("iMass").Value
            Next
            oPrd.ReferenceProduct.UserRefProperties.item("iMass").Value = total
        Else
            total = oPrd.ReferenceProduct.UserRefProperties.item("iMass").Value
        End If
        Assmass = total
    End Function
    ' Generate product information
    Public Function infoPrd(oPrd)
        Dim arr(1 To 9)
            With oPrd.ReferenceProduct
                arr(1) = .PartNumber
                arr(2) = .nomenclature
                arr(3) = .definition
                arr(4) = oPrd.Name
            End With
        Dim colls
           Set colls = oPrd.ReferenceProduct.UserRefProperties
                arr(5) = getAtt("Material", colls)(1)
                arr(6) = getAtt("Density", colls)(1)
                arr(7) = getAtt("Mass", colls)(1)
                arr(8) = getAtt("Thickness", colls)(1)
        On Error GoTo 0
                arr(9) = count_me(oPrd)
        infoPrd = arr()
    End Function
Public Function recurPrd(oPrd, LV)
    Dim mapping As Variant
    Dim mapsize
    mapping = Array(0, 1, 2, 3, 4, 9, 7, 0, 5, 8, 0, 5, 6, 0, 0, 0)
    mapsize = UBound(mapping) + 2
    If counter = 1 Then
          ReDim bomdata(1 To 220, 1 To mapsize) ' 扩展列：2列原有数据 + XX列产品属性
    End If
    Dim data
    data = infoPrd(oPrd)
   '=====生成行数据 === bom行数据直接按最大列处理得到一行数据
    bomdata(counter, 1) = counter
    bomdata(counter, 2) = LV
    Dim i
          For i = 1 To UBound(bomdata, 2) - 2
                bomdata(counter, i + 2) = ""  '默认赋值为空，当对应列需要数据，
                If mapping(i) <= UBound(data) And mapping(i) >= LBound(data) Then
                    bomdata(counter, i + 2) = data(mapping(i))
                End If
          Next
    Dim children As Products
    Set children = oPrd.Products
    If children.Count > 0 Then
            Dim bdict
            Set bdict = CreateObject("Scripting.Dictionary")
        For i = 1 To children.Count
            If bdict.Exists(children.item(i).PartNumber) = False Then
               bdict(children.item(i).PartNumber) = 1
               counter = counter + 1
               recurPrd children.item(i), LV + 1
            End If
        Next
    End If
    recurPrd = bomdata
End Function
Public Function modatt(oPrd, odata)
    Dim refPrd: Set refPrd = oPrd.ReferenceProduct
    Dim colls, i
    Dim currAtt As Variant
    Set currAtt = infoPrd(oPrd)
    For i = 1 To 8
        If odata(i) <> "" And odata(i) <> currAtt(i) Then
            On Error Resume Next
            Select Case i
                Case 1: refPrd.PartNumber = odata(i)
                Case 2: refPrd.nomenclature = odata(i)
                Case 3: refPrd.definition = odata(i)
                Case 4: oPrd.Name = odata(i)
                Case 5:
                    Set colls = refPrd.UserRefProperties
                    colls.item(Att(i - 4)).Value = odata(i)
                Case 6
                    On Error Resume Next
                    Set colls = refPrd.Parent.Part.Parameters.RootParameterSet.ParameterSets.item("Part_info").DirectParameters
                    colls.item(Att(i - 4)).Value = odata(i)
                    Error.Clear
                    On Error GoTo 0
            End Select
            On Error GoTo 0
        End If
    Next
End Function
Public Sub initPrd(oPrd)
    Dim refPrd, colls
    Set refPrd = oPrd.ReferenceProduct
    Dim NT As Variant
        NT = Array( _
            Array("Mass", "Mass"), _
            Array("Material", "String"), _
            Array("Thickness", "Length"), _
            Array("Density", "Density") _
        )
    Dim usrp(0 To 0)
    Set colls = refPrd.UserRefProperties
    Dim i
    For i = 0 To 0
        Set usrp(i) = New Class_para
        usrp(i).SetNT NT(i)(0), NT(i)(1)
        checkPara usrp(i), colls
    Next
    On Error Resume Next
         Dim oPrt
         Set oPrt = refPrd.Parent.Part
        If Err.Number <> 0 Then
            Err.Clear
            Set oPrt = Nothing
        End If
    On Error GoTo 0
    If Not oPrt Is Nothing Then
            Call iniPrt(oPrd)
    Else
            Dim k
            k = 0
             Dim oref, pubid, oPub, pubs
           Set pubs = refPrd.Publications
            If getAtt(usrp(k).Name, pubs)(0) Is Nothing Then
                 Dim iName
                 iName = refPrd.PartNumber & "\" & "Properties" & "\" & usrp(k).Name
                Set oref = refPrd.CreateReferenceFromName(iName)
                Debug.Print usrp(k).Name
                Set oPub = pubs.Add(usrp(k).Name)
                pubs.SetDirect usrp(k).Name, oref
            End If
    End If
End Sub
Private Sub iniPrt(oPrd)
    Dim refPrd, colls
    Dim i
    Set refPrd = oPrd.ReferenceProduct
    Dim oPrt: Set oPrt = refPrd.Parent.Part
'============创建usrp属性=================
    Dim NT As Variant
        NT = Array( _
            Array("Mass", "Mass"), _
            Array("Material", "String"), _
            Array("Thickness", "Length"), _
            Array("Density", "Density") _
        )
    Dim usrp(0 To 3)
    Set colls = refPrd.UserRefProperties
    For i = 0 To 3
        Set usrp(i) = New Class_para
        usrp(i).SetNT NT(i)(0), NT(i)(1)
        checkPara usrp(i), colls
    Next
'============创建part_info内参数=================
    '============创建参数集合=================
        Dim infoset
        Set infoset = New Class_para
        infoset.SetNT "Part_info", "ParameterSet"
        Set colls = oPrt.Parameters.RootParameterSet.ParameterSets
        checkPara infoset, colls
        Set colls = infoset.obj.DirectParameters
     '====创建part_info内参数====
        Dim iBodys
        Set iBodys = New Class_para
        iBodys.SetNT "iBodys", "list"
        checkPara iBodys, colls
        If getPara(oPrt.mainbody, iBodys.obj.valuelist)(0) Is Nothing Then
                iBodys.obj.valuelist.Add oPrt.mainbody
        End If
        Dim infoPara(0 To 3)
        Dim PNT
         PNT = Array( _
            Array("sumVol", "Volume"), _
            Array("Thickness", "Length"), _
            Array("Density", "Density"), _
            Array("Mass", "Mass") _
        )
        For i = 0 To 3
            Set infoPara(i) = New Class_para
            infoPara(i).SetNT PNT(i)(0), PNT(i)(1)
            checkPara infoPara(i), colls
        Next
    '======创建零件发布======
        Dim oref, pubid, oPub, pubs
        Set pubs = refPrd.Publications
         Dim k
          For k = 1 To 3
               If getAtt(infoPara(k).Name, pubs)(0) Is Nothing Then
                   Dim iName
                   iName = refPrd.PartNumber & "\Parameters\Part_info\" & infoPara(k).Name
                   Set oref = refPrd.CreateReferenceFromName(iName)
                   Set oPub = pubs.Add(infoPara(k).Name)
                   pubs.SetDirect infoPara(k).Name, oref
               Else
               Debug.Print "不需要发布" & infoPara(k).Name
                End If
           Next
    '创建所有relations
            Set colls = oPrt.relations
            Dim oRule
            Set oRule = New Class_para
                   '---创建EKL
            oRule.Reset
            oRule.SetNT "sum_all_Vol", "Program"
            oRule.Str = eklstr
            oRule.Desc = "汇总体积"
            checkPara oRule, colls
                    '---创建Cal_mass
            oRule.Reset
            oRule.SetNT "cal_Mass", "Formula"
            oRule.Desc = "链接重量"
            oRule.Str = "Part_info\sumVol *Part_info\Density"
            Set oRule.Target = infoPara(3).obj
            checkPara oRule, colls
                    '---创建link_mass
            oRule.Reset
            oRule.SetNT "link_Mass", "Formula"
            oRule.Desc = "链接重量"
            oRule.Str = "Part_info\Mass"
            Set oRule.Target = refPrd.UserRefProperties.item("Mass")
            checkPara oRule, colls
                    '---创建link_thickness
            oRule.Reset
            oRule.SetNT "link_Thickness", "Formula"
            oRule.Desc = "链接厚度"
            oRule.Str = "Part_info\Thickness"
            Set oRule.Target = refPrd.UserRefProperties.item("Thickness")
            checkPara oRule, colls
                    '---创建link_Density
            oRule.Reset
            oRule.SetNT "link_Density", "Formula"
            oRule.Desc = "链接密度"
            oRule.Str = "Part_info\Density"
            Set oRule.Target = refPrd.UserRefProperties.item("Density")
            checkPara oRule, colls
            oRule.Reset
End Sub
Private Function checkPara(thispara, colls)
    If getPara(thispara, colls)(0) Is Nothing Then
'        Debug.Print "需要创建" & thispara.Name
        paraCreatobj thispara, colls   '创建参数和公式时已经是默认值'
    Else
'        Debug.Print "不需要创建" & thispara.obj.Name
        Set thispara.obj = getPara(thispara, colls)(0) '已有参数和公式，校验默认值'
        Select Case thispara.iType  '只有formula和rule才需要校核其内容
            Case "Program", "Formula"
                If thispara.obj.Value <> thispara.Str Then
                 thispara.obj.Modify thispara.Str
'                  Debug.Print "已经修改" & thispara.obj.Name
                End If
                thispara.obj.Hidden = True
        End Select
    End If
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
Private Function getAtt(iName, collection)
    Dim arr(1) ' 正确声明数组
    On Error Resume Next
        Set arr(0) = collection.item(iName)
        If Err.Number = 0 Then ' 检查是否成功获取对象
            arr(1) = arr(0).Value
            getAtt = arr
        Else
            getAtt = Array(Nothing, "__")
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
                 thispara.obj.Hidden = True
            Case "Program"
                  On Error Resume Next
                    Set thispara.obj = colls.CreateProgram(thispara.Name, thispara.Desc, thispara.Str)
                  If Error.Number <> 0 Then
                    Debug.Print "缺少KWA lisence, 无法创建rule单位"
                  End If
                    Error.Clear
                  On Error GoTo 0
    End Select
    Debug.Print "已经创建" & thispara.Name
End Function
Private Function count_me(oPrd)  '获取兄弟字典
     Dim i, oDict, QTy, pn
         QTy = 1
     On Error Resume Next
          If TypeOf oPrd.Parent Is Products Then    '若有父级产品'获取兄弟字典
                    Dim oParent: Set oParent = oPrd.Parent.Parent
                   Set oDict = CreateObject("Scripting.Dictionary")
                   For i = 1 To oParent.Products.Count
                          pn = oParent.Products.item(i).PartNumber
                            If oDict.Exists(pn) = True Then
                                oDict(pn) = oDict(pn) + 1
                            Else
                                oDict(pn) = 1
                            End If
                      Next
            QTy = oDict(oPrd.PartNumber)
          End If
    count_me = QTy
End Function
Public Function catSel1(prompt As String) As Object
    oSel.Clear
    Dim iType(0)
    iType(0) = "Product"
        MsgBox prompt
    If oSel.SelectElement2(iType, prompt, False) = "Normal" Then
        If oSel.Count = 1 Then
            Set catSel1 = oSel.item(1).LeafProduct
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
            catSelx.Add oSel.item(i).LeafProduct
        Next
    End If
    oSel.Clear
End Function
Public Function defgprd()    '获取要读取或修改的  全局变量 gPrd定义
     imsg = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
     Dim oPrd
        Select Case imsg
            Case 7 '===选择“否”====
            Set rootPrd = CATIA.ActiveDocument.Product
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
'           imsg = "你选择的产品是" & oPrd.PartNumber & "是否继续"
'          MsgBox imsg
        End If
        Set gPrd = oPrd
        Set defgprd = oPrd
End Function
Public Sub freePrd()
     Set gPrd = Nothing
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
            Set selPrd = oSel.item(1).LeafProduct.ReferenceProduct
            MsgBox "你选择的产品是" & selPrd.PartNumber
            oSel.Clear
    Else
        MsgBox "请只选择一个产品"
        Exit Function
    End If
End Function

