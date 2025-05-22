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
Public bigbom As Variant
Private osel
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
        Set osel = CATIA.ActiveDocument.Selection: osel.Clear
ErrorHandler:
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "CATIA 程序错误：" & Err.Description, vbCritical
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
Public Function Assmass(oprd)

        Dim i
        Dim total, children: total = 0
        Set children = oprd.Products
        If oprd.Products.Count > 0 Then
            For i = 1 To children.Count
                Assmass (children.item(i)) 'Assmass = Assmass +
                total = total + children.item(i).ReferenceProduct.UserRefProperties.item("Mass").Value
            Next
            oprd.ReferenceProduct.UserRefProperties.item("Mass").Value = total
        Else
            total = oprd.ReferenceProduct.UserRefProperties.item("Mass").Value
        End If
        Assmass = total

        MsgBox "重量已经计算"
    End Function
' Generate product information
    Public Function infoPrd(oprd)
        Dim arr(1 To 9)
            With oprd.ReferenceProduct
                arr(1) = .PartNumber
                arr(2) = .nomenclature
                arr(3) = .definition
                arr(4) = oprd.Name
            End With
        Dim colls
           Set colls = oprd.ReferenceProduct.UserRefProperties
                arr(5) = getAtt("Material", colls)(1)
                arr(6) = getAtt("Density", colls)(1)
                arr(7) = getAtt("Mass", colls)(1)
                arr(8) = getAtt("Thickness", colls)(1)
        On Error GoTo 0
                arr(9) = count_me(oprd)
        infoPrd = arr
    End Function
Public Function recurPrd(oprd, LV)
    Dim mapping As Variant
    mapping = Array(0, 0, 0, 1, 2, 3, 4, 9, 7, 0, 5, 8, 0, 5, 6, 0, 0, 0)
    
    Dim bomsize
        bomsize = UBound(mapping)
        
    Dim i, j
        If counter = 1 Then
        ReDim bigbom(1 To fn, 1 To bomsize)
            For i = 1 To fn
                For j = 1 To bomsize
                    bigbom(i, j) = ""
                Next j
            Next i
    End If
   
    Dim tempdata
    tempdata = infoPrd(oprd)
    
   '=====生成行数据 === bom行数据直接按最大列处理得到一行数据
   
        For i = 1 To UBound(bigbom, 2)
            bigbom(counter, i) = ""  '默认赋值为空
            Select Case i
                Case 1: bigbom(counter, i) = counter
                Case 2: bigbom(counter, i) = LV
                Case Else
                    If mapping(i) <= UBound(tempdata) And mapping(i) >= LBound(tempdata) Then '当对应列map值在tempdata元素范围内，
                        bigbom(counter, i) = tempdata(mapping(i))
                    End If
            End Select
           Next
           
            Dim children As Products
            Set children = oprd.Products
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
            
  ReDim finalbom(1 To counter, 1 To 17)
  
  For i = 1 To UBound(finalbom, 1)
    For j = 1 To UBound(finalbom, 2)
    finalbom(i, j) = bigbom(i, j)
    Next
    Next
            
            
        recurPrd = finalbom

End Function
Public Function modatt(oprd, odata)
    Dim refPrd: Set refPrd = oprd.ReferenceProduct
    Dim colls, i
    Dim currAtt As Variant
     currAtt = infoPrd(oprd)
    For i = 1 To 6
        If odata(i) <> "" Then   'And odata(i) <> currAtt(i)
            On Error Resume Next
                Select Case i
                    Case 1: refPrd.PartNumber = odata(i)
                    Case 2: refPrd.nomenclature = odata(i)
                    Case 3: refPrd.definition = odata(i)
                    Case 4: oprd.Name = odata(i)
                    Case 5:
                        Set colls = refPrd.UserRefProperties
                        colls.item("Material").Value = odata(i)
                    Case 6
                        On Error Resume Next
                            Set colls = refPrd.Parent.Part.Parameters.RootParameterSet.ParameterSets.item("Part_info").DirectParameters
                            colls.item("Density").Value = odata(i)
                            Error.Clear
                        On Error GoTo 0
                End Select
                
'                Error.Clear
            On Error GoTo 0
            
        End If
    Next
 End Function
 
Public Sub initPrd(oprd)
    Dim refPrd, colls
    Set refPrd = oprd.ReferenceProduct
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
         Dim oprt
         Set oprt = refPrd.Parent.Part
        If Err.Number <> 0 Then
            Err.Clear
            Set oprt = Nothing
        End If
    On Error GoTo 0
    If Not oprt Is Nothing Then
            Call iniPrt(oprd)
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

Private Sub iniPrt(oprd)
    Dim refPrd, colls
    Dim i
    Set refPrd = oprd.ReferenceProduct
    Dim oprt: Set oprt = refPrd.Parent.Part
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
        Set colls = oprt.Parameters.RootParameterSet.ParameterSets
        checkPara infoset, colls
        Set colls = infoset.obj.DirectParameters
     '====创建part_info内参数====
        Dim iBodys
        Set iBodys = New Class_para
        iBodys.SetNT "iBodys", "list"
        checkPara iBodys, colls
        If getPara(oprt.mainbody, iBodys.obj.valuelist)(0) Is Nothing Then
                iBodys.obj.valuelist.Add oprt.mainbody
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
            Set colls = oprt.relations
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
Private Function count_me(oprd)  '获取兄弟字典
     Dim i, oDict, QTy, pn
         QTy = 1
     On Error Resume Next
          If TypeOf oprd.Parent Is Products Then    '若有父级产品'获取兄弟字典
                    Dim oParent: Set oParent = oprd.Parent.Parent
                   Set oDict = CreateObject("Scripting.Dictionary")
                   For i = 1 To oParent.Products.Count
                          pn = oParent.Products.item(i).PartNumber
                            If oDict.Exists(pn) = True Then
                                oDict(pn) = oDict(pn) + 1
                            Else
                                oDict(pn) = 1
                            End If
                      Next
            QTy = oDict(oprd.PartNumber)
          End If
    count_me = QTy
End Function
Public Function catSel1(prompt As String) As Object
    osel.Clear
    Dim iType(0)
    iType(0) = "Product"
        MsgBox prompt
    If osel.SelectElement2(iType, prompt, False) = "Normal" Then
        If osel.Count = 1 Then
            Set catSel1 = osel.item(1).LeafProduct
        End If
    End If
    osel.Clear
End Function
Public Function catSelx(prompt As String) As collection
    Set catSelx = New collection
    osel.Clear
    Dim iType(0)
    iType(0) = "Product"
    If osel.SelectElement3(iType, prompt, True, 2, True) = "Normal" Then
        Dim i As Integer
        For i = 1 To osel.Count
            catSelx.Add osel.item(i).LeafProduct
        Next
    End If
    osel.Clear
End Function
Public Function defgprd()    '获取要读取或修改的  全局变量 gPrd定义
     imsg = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")
     Dim oprd
     
     Set oprd = Nothing
     
        Select Case imsg
            Case 7 '===选择“否”====
            Set rootPrd = CATIA.ActiveDocument.Product
            Set oprd = rootPrd
            Case 2:
            
            
           Set oprd = Nothing
            Case 6  '===选择“是”,进行产品选择====
                On Error Resume Next
                    Set oprd = selPrd()
                If Err.Number <> 0 Then
                    Err.Clear
                    Exit Function
                End If
        End Select
        
        Set defgprd = oprd
        Set oprd = Nothing
End Function
Public Sub freePrd()
     Set gPrd = Nothing
End Sub
Public Function selPrd()
    osel.Clear
   Dim status, iType(0)
'   catia.ActiveWindow.WindowState = 0
    osel.Clear
    iType(0) = "Product"
    If osel.Count2 = 0 Then
         ' MsgBox "请在CATIA结构树选择要读取的产品"
        status = osel.SelectElement2(iType, "请选择要读取的产品", False)
        'status = oSel.SelectElement3(iType, "请选择要读取的产品", True, 2, False)
    End If
    If status = "Cancel" Then
        Exit Function
    End If
    If status = "Normal" And osel.Count2 = 1 Then
            Set selPrd = osel.item(1).LeafProduct.ReferenceProduct
          '  MsgBox "你选择的产品是" & selPrd.PartNumber
            osel.Clear
    Else
        MsgBox "请只选择一个产品"
        Exit Function
    End If
End Function

