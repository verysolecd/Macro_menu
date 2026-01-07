VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_PDM"
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

Public Event ProductChanged()  ' 事件声明
Private mProduct As Object  ' 私有变量存储当前产品

Public bigbom As Variant
Public msel
Private imsg
Private Att(1 To 10)
Private Const xx = "测试成功"
Private Const xy = "测试失败"
Private Const eklname = "sumVol"
Private Const ekldesc = "sum of vol of bodylist"
Private Const eklstr = "let lst(list) set lst=Part_info\iBodys   let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol =V"  '使用Const关键字定义常量
Private Sub Class_Initialize()
     Set allPN = KCL.InitDic
     Set mProduct = Nothing
      iniarr
        On Error GoTo errorhandler
        
        
        
        If CATIA.Documents.count = 0 Then
          MsgBox "当前未打开CATIA文档"
          
          
          Else
          
          Set rootDoc = CATIA.ActiveDocument
            Set rootprd = CATIA.ActiveDocument.Product
            If Not rootDoc Is Nothing Then
            Set msel = CATIA.ActiveDocument.Selection: msel.Clear
            End If
        
        
        If Not rootprd Is Nothing Then
            rootprd.ApplyWorkMode (3)
        End If
          
        End If
           
             
        On Error GoTo 0
        
        
       
        
        
        
errorhandler:
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
        Dim i, total, children
            total = 0
        Set children = oprd.Products
        If oprd.Products.count > 0 Then
            For i = 1 To children.count
                Assmass (children.item(i)) 'Assmass = Assmass +
                total = total + children.item(i).ReferenceProduct.UserRefProperties.item("Mass").value
            Next
                oprd.ReferenceProduct.UserRefProperties.item("Mass").value = total
        Else
                total = oprd.ReferenceProduct.UserRefProperties.item("Mass").value
        End If
        Assmass = total
 End Function
' Generate product information
    Public Function infoPrd(oprd)
        Dim arr(1 To 9)
            With oprd.ReferenceProduct
                arr(1) = .PartNumber
                arr(2) = .Nomenclature
                arr(3) = .Definition
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
Public Function attLv2Prd(oprd)
    Dim tempdata, i, j
    tempdata = infoPrd(oprd)
    If counter = 1 Then
        ReDim bigbom(1 To gfn, 0 To UBound(tempdata)) '初始化所有bigbom元素为""
            For i = 1 To gfn
                For j = 0 To UBound(tempdata)
                    bigbom(i, j) = ""
                Next j
            Next i
    End If
   For i = 1 To UBound(tempdata)
       bigbom(counter, i) = tempdata(i)
   Next
  
   Dim children As Products: Set children = oprd.Products
            If children.count > 0 Then
                For j = 1 To children.count
                    counter = counter + 1
                    tempdata = infoPrd(children.item(j))
                   For i = 1 To UBound(tempdata)
                        bigbom(counter, i) = tempdata(i)
                   Next
                Next
            End If
  ReDim finalbom(1 To counter, 1 To UBound(tempdata))
      For i = 1 To UBound(finalbom, 1)
        For j = 1 To UBound(finalbom, 2)
          
          finalbom(i, j) = bigbom(i, j)
      
        Next
      Next
        attLv2Prd = finalbom
        counter = 1
End Function

Public Function recurInfoPrd(oprd, LV)

    Dim tempdata, i, j
    tempdata = infoPrd(oprd) '此数组是1-9
    
    If counter = 1 Then  '初始化所有bigbom元素为""
        ReDim bigbom(1 To gfn, 0 To UBound(tempdata))
            For i = 1 To gfn
                For j = 0 To UBound(tempdata)
                    bigbom(i, j) = ""
                Next j
            Next i
    End If
    '===================
   For i = 0 To UBound(tempdata)   'tempdata 数组是1-9
        Select Case i
             Case 0: bigbom(counter, i) = LV
             Case Else:
                 bigbom(counter, i) = tempdata(i)
             End Select
   Next
       Dim children: Set children = oprd.Products
            If children.count > 0 Then
                    Dim bdict: Set bdict = KCL.InitDic
                For i = 1 To children.count
                    If bdict.Exists(children.item(i).PartNumber) = False Then
                       bdict(children.item(i).PartNumber) = 1
                       counter = counter + 1
                       recurInfoPrd children.item(i), LV + 1
                    End If
                Next
            End If
            
    ReDim finalbom(1 To counter, 0 To UBound(tempdata))
      For i = 1 To UBound(finalbom, 1)
        For j = 0 To UBound(finalbom, 2)
            finalbom(i, j) = bigbom(i, j)
        Next
      Next
    recurInfoPrd = finalbom
End Function


Public Function recurPrd(oprd, LV)
    Dim mapping As Variant
    mapping = Array(0, 0, 0, 1, 2, 3, 4, 9, 7, 0, 5, 8, 0, 5, 6, 0, 0, 0)
    
    Dim bomsize
        bomsize = UBound(mapping)
    Dim i, j
        If counter = 1 Then
        ReDim bigbom(1 To gfn, 1 To bomsize)
            For i = 1 To gfn
                For j = 1 To bomsize
                    bigbom(i, j) = ""
                Next j
            Next i
    End If   '初始化所有bigbom元素为""
    Dim tempdata
    tempdata = infoPrd(oprd)
   '=====生成行数据 === bom行数据直接按最大列处理得到一行数据
        For i = 1 To UBound(bigbom, 2)  '从1到最大列
            bigbom(counter, i) = ""  '默认赋值为空
            Select Case i
                Case 1: bigbom(counter, i) = counter  'counter是全局变量,遍历BOM时为1
                Case 2: bigbom(counter, i) = LV
                Case Else
                    If mapping(i) <= UBound(tempdata) And mapping(i) >= LBound(tempdata) Then '当对应列map值在tempdata元素范围内，
                        bigbom(counter, i) = tempdata(mapping(i))
                    End If
            End Select
           Next
            Dim children As Products
            Set children = oprd.Products
            If children.count > 0 Then
                    Dim bdict
                    Set bdict = KCL.InitDic
                For i = 1 To children.count
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
        If odata(i) <> "" And odata(i) <> currAtt(i) Then
            On Error Resume Next
                Select Case i
                    Case 1: refPrd.PartNumber = odata(i)
                    Case 2: refPrd.Nomenclature = odata(i)
                    Case 3: refPrd.Definition = odata(i)
                    Case 4: oprd.Name = odata(i)
                    Case 5:
                        Set colls = refPrd.UserRefProperties
                        colls.item("Material").value = odata(i)
                    Case 6
                        On Error Resume Next
                            Set colls = refPrd.Parent.part.Parameters.RootParameterSet.ParameterSets.item("Part_info").DirectParameters
                            colls.item("Density").value = odata(i)
                            Error.Clear
                        On Error GoTo 0
                End Select
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
        Set usrp(i) = New Cls_Para
        usrp(i).SetNT NT(i)(0), NT(i)(1)
        checkPara usrp(i), colls
    Next
    
    On Error Resume Next
         Dim oprt
         Set oprt = refPrd.Parent.part
        If Err.Number <> 0 Then
            Err.Clear
            Set oprt = Nothing
        End If
    On Error GoTo 0
    
    If Not oprt Is Nothing Then
        On Error Resume Next
                Call iniPrt(oprd)
        On Error GoTo 0
    Else
            Dim k
            k = 0
             Dim oref, pubid, oPub, pubs
           Set pubs = refPrd.Publications
            If getAtt(usrp(k).Name, pubs)(0) Is Nothing Then
                 Dim iName
                 iName = refPrd.PartNumber & "\" & "Properties" & "\" & usrp(k).Name
                Set oref = refPrd.CreateReferenceFromName(iName)
'                Debug.Print usrp(k).Name
                Set oPub = pubs.Add(usrp(k).Name)
                pubs.SetDirect usrp(k).Name, oref
            End If
    End If
End Sub

Private Sub iniPrt(oprd)
    Dim refPrd, colls
    Dim i
    Set refPrd = oprd.ReferenceProduct
    Dim oprt: Set oprt = refPrd.Parent.part

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
            Set usrp(i) = New Cls_Para
            usrp(i).SetNT NT(i)(0), NT(i)(1)
            checkPara usrp(i), colls
        Next
'============创建part_info内参数=================
    '============创建参数集合=================
        Dim infoset
        Set infoset = New Cls_Para
        infoset.SetNT "Part_info", "ParameterSet"
        Set colls = oprt.Parameters.RootParameterSet.ParameterSets
        checkPara infoset, colls
        Set colls = infoset.obj.DirectParameters
     '====创建part_info内参数====
        Dim iBodys
        Set iBodys = New Cls_Para
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
            Set infoPara(i) = New Cls_Para
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
            Set oRule = New Cls_Para
                   '---创建EKL
            oRule.Reset
            oRule.SetNT "sum_all_Vol", "Program"
            oRule.str = eklstr
            oRule.Desc = "计算体积"
            checkPara oRule, colls
                    '---创建Cal_mass
            oRule.Reset
            oRule.SetNT "cal_Mass", "Formula"
            oRule.Desc = "计算重量"
    
            oRule.str = "round(Part_info\sumVol *Part_info\Density,""kg"",3)"   ''''"Part_info\sumVol *Part_info\Density"
            Set oRule.Target = infoPara(3).obj
            checkPara oRule, colls
                    '---创建link_mass
            oRule.Reset
            oRule.SetNT "link_Mass", "Formula"
            oRule.Desc = "链接重量"
            oRule.str = "Part_info\Mass"
            Set oRule.Target = refPrd.UserRefProperties.item("Mass")
            checkPara oRule, colls
                    '---创建link_thickness
            oRule.Reset
            oRule.SetNT "link_Thickness", "Formula"
            oRule.Desc = "链接厚度"
            oRule.str = "Part_info\Thickness"
            Set oRule.Target = refPrd.UserRefProperties.item("Thickness")
            checkPara oRule, colls
                    '---创建link_Density
            oRule.Reset
            oRule.SetNT "link_Density", "Formula"
            oRule.Desc = "链接密度"
            oRule.str = "Part_info\Density"
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
                If thispara.obj.value <> thispara.str Then
                 thispara.obj.Modify thispara.str
'                  Debug.Print "已经修改" & thispara.obj.Name
                End If
                thispara.obj.Hidden = True
        End Select
    End If
End Function
Private Function paraCreatobj(thispara, colls)
        Select Case thispara.iType
            Case "ParameterSet"
                Set thispara.obj = colls.CreateSet(thispara.Name)
            Case "list"
                Set thispara.obj = colls.CreateList(thispara.Name)
            Case "Mass", "Density", "Length", "Volume" '所有dimension参数
                Set thispara.obj = colls.CreateDimension(thispara.Name, thispara.iType, thispara.str)
            Case "String"
                Set thispara.obj = colls.CreateString(thispara.Name, thispara.str)
            Case "Formula"
                Set thispara.obj = colls.CreateFormula(thispara.Name, thispara.Desc, thispara.Target, thispara.str)
                 thispara.obj.Hidden = True
            Case "Program"
                  On Error Resume Next
                    Set thispara.obj = colls.CreateProgram(thispara.Name, thispara.Desc, thispara.str)
                    thispara.obj.Hidden = True
                  If Error.Number <> 0 Then
'                    Debug.Print "缺少KWA lisence, 无法创建rule单位"
                  End If
                    Error.Clear
                  On Error GoTo 0
    End Select
   Debug.Print "已经创建" & thispara.Name
End Function
Private Function getPara(thispara, collection)
    Dim arr(1) ' 正确声明数组
    On Error Resume Next
        Set arr(0) = collection.item(thispara.Name)
        If Err.Number = 0 Then ' 检查是否成功获取对象
            arr(1) = arr(0).value
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
            arr(1) = arr(0).value
            getAtt = arr
        Else
            getAtt = Array(Nothing, "__")
        End If
    On Error GoTo 0
End Function

Private Function count_me(oprd)  '获取兄弟字典
     Dim i, oDict, QTy, pn
         QTy = 1
     On Error Resume Next
          If TypeOf oprd.Parent Is Products Then    '若有父级产品'获取兄弟字典
                   Dim oParent: Set oParent = oprd.Parent.Parent
                   Set oDict = KCL.InitDic
                   For i = 1 To oParent.Products.count
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

Public Function catSelx(prompt As String) As collection
    Set catSelx = New collection
    msel.Clear
    Dim iType(0)
    iType(0) = "Product"
    If msel.SelectElement3(iType, prompt, True, 2, True) = "Normal" Then
        Dim i As Integer
        For i = 1 To msel.count
            catSelx.Add msel.item(i).LeafProduct
        Next
    End If
    msel.Clear
End Function
Public Function getiPrd()    '获取要读取或修改的  全局变量 gPrd定义
  Dim oprd: Set oprd = Nothing
    Dim btn, bTitle, bResult
        bTitle = "请选择产品"
      imsg = "“是”选择产品，“否”读取根产品，“取消退出”"
      btn = vbYesNoCancel + vbDefaultButton2 'vbExclamation
      bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
           Select Case bResult
              Case 7 '===选择“否”====
                Set oprd = CATIA.ActiveDocument.Product
              Case 6  '===选择“是”,进行产品选择====
                 On Error Resume Next
                   Set oprd = KCL.SelectItem("请选择产品", "Product")
                If Err.Number <> 0 Then
                    Err.Clear
                    Exit Function
                End If
          End Select
        Set getiPrd = oprd
        Set oprd = Nothing
End Function
Public Sub freePrd()
     Set gPrd = Nothing

End Sub

Public Function gxBom(oprd, LV)
    Dim mapping As Variant
    mapping = Array(0, 0, 0, 0, 0, 0, 1, 0, 3, 2, 0, 9, 0, 7, 0, 5, 0)
    Dim bomsize
        bomsize = UBound(mapping)
    Dim i, j
        If counter = 1 Then
        ReDim bigbom(1 To gfn, 1 To bomsize)
            For i = 1 To gfn
                For j = 1 To bomsize
                    bigbom(i, j) = ""
                Next j
            Next i
    End If   '初始化所有bigbom元素为""
    Dim tempdata
    tempdata = infoPrd(oprd)
   '=====生成行数据 === bom行数据直接按最大列处理得到一行数据
        For i = 1 To UBound(bigbom, 2)  '从1到最大列
            bigbom(counter, i) = ""  '默认赋值为空
            Select Case i
                Case 1: bigbom(counter, i) = counter  'counter是全局变量,遍历BOM时为1
                Case 2, 3, 4, 5:
                    If LV <= 4 Then
                        If LV = i - 1 Then
                            bigbom(counter, i) = LV
                        End If
                     Else
                     bigbom(counter, i) = "error"
                    End If
            
                Case Else
                    If mapping(i) <= UBound(tempdata) And mapping(i) >= LBound(tempdata) Then '当对应列map值在tempdata元素范围内，
                        bigbom(counter, i) = tempdata(mapping(i))
                    End If
            End Select
           Next
            Dim children As Products
            Set children = oprd.Products
            If children.count > 0 Then
                    Dim bdict
                    Set bdict = KCL.InitDic
                For i = 1 To children.count
                    If bdict.Exists(children.item(i).PartNumber) = False Then
                       bdict(children.item(i).PartNumber) = 1
                       counter = counter + 1
                       gxBom children.item(i), LV + 1
                    End If
                Next
            End If
  ReDim finalbom(1 To counter, 1 To 16)
  For i = 1 To UBound(finalbom, 1)
    For j = 1 To UBound(finalbom, 2)
        finalbom(i, j) = bigbom(i, j)
    Next
  Next
        gxBom = finalbom
End Function


' 属性：当前产品
Public Property Get CurrentProduct() As Object
    Set CurrentProduct = mProduct
End Property
Public Property Set CurrentProduct(ByVal newProduct As Object)
    ' 设置新产品并触发事件
    Set mProduct = newProduct
    RaiseEvent ProductChanged
End Property
