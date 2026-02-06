VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_PDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=======================================================================================
' 类名: Cls_PDM (UTF-8)
' 描述: CATIA 产品数据 management 类，负责属性获取、BOM 生成、机电协同等。
' 优化点:
'   1. 使用 Collection 蓄积行数据，避免递归中的数组重定义和复制开销。
'   2. 优化 count_me 逻辑，通过预先统计同级零件数量减少冗余扫描。
'   3. 统一属性获取接口，强化类型声明。
'   4. 移除对全局变量 g_counter, gfn 的依赖，增强类的独立性。
'=======================================================================================

Option Explicit

' 事件声明
Public Event ProductChanged()

' 私有变量
Private mProduct As Object          ' 存储当前产品
Private mSelection As Object        ' 存储当前选择集
Private mRows As Collection         ' 用于递归收集 BOM 行数据的临时集合

' 变量及辅助数组
Private bigbom As Variant           ' 内部中间变量
Private Att(1 To 10) As Variant

' 常量定义
Private Const MDL_NAME As String = "Cls_PDM"
Private Const EKL_STR As String = "let lst(list) set lst=Part_info\iBodys let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol =V"

' 初始化
Private Const mdlname As String = "Cls_PDM"
Private Sub Class_Initialize()
    On Error Resume Next
    Set g_allPN = KCL.InitDic ' 保持与全局变量同步
    Set mProduct = Nothing
    
    If CATIA.Documents.count > 0 Then
        Set mSelection = CATIA.ActiveDocument.Selection
        mSelection.Clear
        ' 根对象初始化
        Set rootDoc = CATIA.ActiveDocument
        Set rootPrd = CATIA.ActiveDocument.Product
        If Not rootPrd Is Nothing Then rootPrd.ApplyWorkMode (3)
    End If
    
    ' 初始化 Att 数组 (保持旧版兼容)
    iniarr
    On Error GoTo 0
End Sub

Public Sub iniarr()
    Att(1) = "Mass"
    Att(2) = "Material"
    Att(3) = "Thickness"
    Att(4) = "Density"
End Sub

'---------------------------------------------------------------------------------------
' 基础属性获取及辅助函数
'---------------------------------------------------------------------------------------

' 属性接口
Public Property Get CurrentProduct() As Object
    Set CurrentProduct = mProduct
End Property

Public Property Set CurrentProduct(ByVal newProduct As Object)
    Set mProduct = newProduct
    RaiseEvent ProductChanged
End Property

' 统一获取属性/参数的值
' 返回：Array(对象, 值)；失败返回 Array(Nothing, "__")
Private Function GetPropertyValue(ByVal Collection As Object, ByVal ItemIdentifier As Variant) As Variant
    On Error Resume Next
    Dim objItem As Object
    Set objItem = Collection.item(ItemIdentifier)
    If Err.Number = 0 Then
        GetPropertyValue = Array(objItem, objItem.value)
    Else
        GetPropertyValue = Array(Nothing, "__")
    End If
    On Error GoTo 0
End Function

' 获取测量工具
Public Function GetMeasurable(ByVal itm As Object) As Object
    Set GetMeasurable = Nothing
    If Not itm Is Nothing Then
        On Error Resume Next
        Dim spa As Object: Set spa = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
        Set GetMeasurable = spa.GetMeasurable(itm)
        On Error GoTo 0
    End If
End Function

' 获取长度
Public Function getlength(ByVal itm As Object) As Double
    getlength = 0
    Dim meas As Object: Set meas = GetMeasurable(itm)
    If Not meas Is Nothing Then getlength = meas.Length
End Function

' 旧版方法兼容性封装
Public Function getmeas(itm)
    Set getmeas = GetMeasurable(itm)
End Function


'---------------------------------------------------------------------------------------
' BOM 生成核心逻辑 (性能优化版)
'---------------------------------------------------------------------------------------

' 获取产品基本属性信息 (返回 1-9 索引数组)
Public Function infoPrd(ByVal oPrd As Object) As Variant
    Dim arr(1 To 9) As Variant
    Dim refPrd As Object: Set refPrd = oPrd.ReferenceProduct
    Dim userProps As Object: Set userProps = refPrd.UserRefProperties
    
    arr(1) = refPrd.PartNumber
    arr(2) = refPrd.Nomenclature
    arr(3) = refPrd.Definition
    arr(4) = oPrd.name
    arr(5) = GetPropertyValue(userProps, "Material")(1)
    arr(6) = GetPropertyValue(userProps, "Density")(1)
    arr(7) = GetPropertyValue(userProps, "Mass")(1)
    arr(8) = GetPropertyValue(userProps, "Thickness")(1)
    arr(9) = CountSamePartNumbers(oPrd)
    
    infoPrd = arr
End Function

' 递归计算并更新产品及其子产品的总质量 (恢复并优化版)
Public Function Assmass(ByVal oPrd As Object) As Double
    Dim children As Products: Set children = oPrd.Products
    Dim totalMass As Double: totalMass = 0
    Dim i As Long
    
    On Error Resume Next
    If children.count > 0 Then
        ' 汇总子产品质量
        For i = 1 To children.count
            totalMass = totalMass + Assmass(children.item(i))
        Next
        ' 更新当前产品的 Mass 属性
        oPrd.ReferenceProduct.UserRefProperties.item("Mass").value = totalMass
    Else
        ' 叶子节点：直接获取其质量
        totalMass = oPrd.ReferenceProduct.UserRefProperties.item("Mass").value
    End If
    On Error GoTo 0
    
    Assmass = totalMass
End Function

' 优化计数逻辑：按需统计同级 PN
Private Function CountSamePartNumbers(ByVal oPrd As Object) As Long
    On Error Resume Next
    If Not TypeOf oPrd.Parent Is Products Then
        CountSamePartNumbers = 1: Exit Function
    End If
    
    Dim productsColl As Products: Set productsColl = oPrd.Parent
    Dim targetPN As String: targetPN = oPrd.PartNumber
    Dim i As Long, cnt As Long: cnt = 0
    
    For i = 1 To productsColl.count
        If productsColl.item(i).PartNumber = targetPN Then cnt = cnt + 1
    Next
    CountSamePartNumbers = cnt
    On Error GoTo 0
End Function

' 高效 BOM 生成 (主入口)
Public Function ProduceBOM(ByVal rootPrd As Object, Optional ByVal filterDuplicate As Boolean = True) As Variant
    Set mRows = New Collection
    
    ' 内部递归执行
    TraverseProductTree rootPrd, 1, filterDuplicate
    
    If mRows.count = 0 Then ProduceBOM = Empty: Exit Function
    
    ' 转换为 2D 数组输出
    Dim rowCount As Long: rowCount = mRows.count
    Dim colCount As Long: colCount = UBound(mRows(1))
    Dim result() As Variant: ReDim result(1 To rowCount, 0 To colCount)
    
    Dim r As Long, c As Long, rowData As Variant
    For r = 1 To rowCount
        rowData = mRows(r)
        For c = 0 To colCount: result(r, c) = rowData(c): Next
    Next
    ProduceBOM = result
End Function

' 递归辅助函数
Private Sub TraverseProductTree(ByVal oPrd As Object, ByVal level As Long, ByVal filterDuplicate As Boolean)
    Dim row(0 To 9) As Variant
    Dim data As Variant: data = infoPrd(oPrd)
    
    row(0) = level
    Dim i As Long
    For i = 1 To 9: row(i) = data(i): Next
    mRows.Add row
    
    Dim children As Products: Set children = oPrd.Products
    If children.count > 0 Then
        Dim dict As Object: Set dict = KCL.InitDic
        For i = 1 To children.count
            Dim child As Object: Set child = children.item(i)
            If filterDuplicate Then
                If Not dict.Exists(child.PartNumber) Then
                    dict.Add child.PartNumber, 1
                    TraverseProductTree child, level + 1, filterDuplicate
                End If
            Else
                TraverseProductTree child, level + 1, filterDuplicate
            End If
        Next
    End If
End Sub

' 旧接口兼容逻辑
Public Function recurInfoPrd(oPrd, Lv) As Variant
    ' 此处原本依赖 g_counter 和 bigbom，现统一使用 ProduceBOM
    recurInfoPrd = ProduceBOM(oPrd, True)
End Function

Public Function attLv2Prd(oPrd) As Variant
    ' 此方法仅取两层信息
    Set mRows = New Collection
    TraverseTwoLevels oPrd
    
    If mRows.count = 0 Then attLv2Prd = Empty: Exit Function
    
    Dim result() As Variant: ReDim result(1 To mRows.count, 1 To 9)
    Dim r As Long, c As Long, rowData As Variant
    For r = 1 To mRows.count
        rowData = mRows(r)
        For c = 1 To 9: result(r, c) = rowData(c): Next
    Next
    attLv2Prd = result
End Function

Private Sub TraverseTwoLevels(ByVal oPrd As Object)
    ' 根产品
    mRows.Add infoPrd(oPrd)
    ' 子产品
    Dim children As Products: Set children = oPrd.Products
    Dim i As Long
    For i = 1 To children.count
        mRows.Add infoPrd(children.item(i))
    Next
End Sub

' gxBom 逻辑重构 (特殊格式 BOM)
Public Function gxBom(ByVal oPrd As Object, ByVal level As Long) As Variant
    Set mRows = New Collection
    Dim mapping As Variant: mapping = Array(0, 0, 0, 0, 0, 0, 1, 0, 3, 2, 0, 9, 0, 7, 0, 5, 0)
    
    TraverseGXFormat oPrd, level, mapping
    
    If mRows.count = 0 Then gxBom = Empty: Exit Function
    
    Dim result() As Variant: ReDim result(1 To mRows.count, 1 To 16)
    Dim r As Long, c As Long, rowData As Variant
    For r = 1 To mRows.count
        rowData = mRows(r)
        For c = 1 To 16: result(r, c) = rowData(c): Next
    Next
    gxBom = result
End Function

Private Sub TraverseGXFormat(ByVal oPrd As Object, ByVal level As Long, ByVal mapping As Variant)
    Dim row(1 To 16) As Variant
    Dim data As Variant: data = infoPrd(oPrd)
    Dim i As Long, currentIdx As Long: currentIdx = mRows.count + 1
    
    For i = 1 To 16
        Select Case i
            Case 1: row(i) = currentIdx
            Case 2, 3, 4, 5: If level = i - 1 Then row(i) = level Else row(i) = ""
            Case Else
                Dim mVal As Long: mVal = mapping(i)
                If mVal > 0 And mVal <= 9 Then row(i) = data(mVal) Else row(i) = ""
        End Select
    Next
    mRows.Add row
    
    Dim children As Products: Set children = oPrd.Products
    If children.count > 0 Then
        Dim dict As Object: Set dict = KCL.InitDic
        For i = 1 To children.count
            If Not dict.Exists(children.item(i).PartNumber) Then
                dict.Add children.item(i).PartNumber, 1
                TraverseGXFormat children.item(i), level + 1, mapping
            End If
        Next
    End If
End Sub

'---------------------------------------------------------------------------------------
' 属性修改与初始化逻辑
'---------------------------------------------------------------------------------------

Public Function modatt(ByVal oPrd As Object, ByVal oData As Variant)
    Dim refPrd As Object: Set refPrd = oPrd.ReferenceProduct
    Dim currAtt As Variant: currAtt = infoPrd(oPrd)
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To 6
        If oData(i) <> "" And oData(i) <> currAtt(i) Then
            Select Case i
                Case 1: refPrd.PartNumber = oData(i)
                Case 2: refPrd.Nomenclature = oData(i)
                Case 3: refPrd.Definition = oData(i)
                Case 4: oPrd.name = oData(i)
                Case 5: refPrd.UserRefProperties.item("Material").value = oData(i)
                Case 6: refPrd.Parent.part.Parameters.RootParameterSet.ParameterSets.item("Part_info").DirectParameters.item("Density").value = oData(i)
            End Select
        End If
    Next
    On Error GoTo 0
End Function

Public Sub initPrd(ByVal oPrd As Object)
    Dim refPrd As Object: Set refPrd = oPrd.ReferenceProduct
    
    ' 初始化基本的 Mass 参数
    EnsureParamExists refPrd.UserRefProperties, "Mass", "Mass"
    
    ' 检查是否含有零件
    Dim oPart As Object
    On Error Resume Next
    Set oPart = refPrd.Parent.part
    On Error GoTo 0
    
    If Not oPart Is Nothing Then
        iniPrt oPrd
    Else
        ' 仅产品：发布属性
        EnsurePublication refPrd, "Mass", "Properties\Mass"
    End If
End Sub

Private Sub iniPrt(ByVal oPrd As Object)
    Dim refPrd As Object: Set refPrd = oPrd.ReferenceProduct
    Dim oPart As Object: Set oPart = refPrd.Parent.part
    
    ' 1. 用户自定义属性
    Dim names As Variant: names = Array("Mass", "Material", "Thickness", "Density")
    Dim types As Variant: types = Array("Mass", "String", "Length", "Density")
    Dim i As Long
    For i = 0 To 3: EnsureParamExists refPrd.UserRefProperties, names(i), types(i): Next
    
    ' 2. 参数集合 Part_info
    Dim infoSet As Object
    On Error Resume Next
    Set infoSet = oPart.Parameters.RootParameterSet.ParameterSets.item("Part_info")
    If infoSet Is Nothing Then Set infoSet = oPart.Parameters.RootParameterSet.ParameterSets.CreateSet("Part_info")
    On Error GoTo 0
    
    ' 3. 零件参数 (包含 iBodys 列表)
    Dim infoNames As Variant: infoNames = Array("sumVol", "Thickness", "Density", "Mass")
    Dim infoTypes As Variant: infoTypes = Array("Volume", "Length", "Density", "Mass")
    For i = 0 To 3: EnsureParamExists infoSet.DirectParameters, infoNames(i), infoTypes(i): Next
    
    ' 特别处理 iBodys 列表
    Dim iBodys As Object
    
    On Error Resume Next
        Set iBodys = infoSet.DirectParameters.item("iBodys")
        If iBodys Is Nothing Then Set iBodys = infoSet.DirectParameters.CreateList("iBodys")
        
        ' 将 MainBody 同步到 iBodys 列表
        If Not oPart.MainBody Is Nothing Then
            Dim bodyInList As Boolean: bodyInList = False
            Dim j As Long
            For j = 1 To iBodys.ValueList.count
                If iBodys.ValueList.item(j).name = oPart.MainBody.name Then
                    bodyInList = True
                    Exit For
                End If
            Next
            If Not bodyInList Then iBodys.ValueList.Add oPart.MainBody
        End If
        Call nosamebdy_bdylst(oPart)
    On Error GoTo 0
    
    
    ' 4. 发布
    For i = 1 To 3: EnsurePublication refPrd, infoNames(i), "Parameters\Part_info\" & infoNames(i): Next
    
    ' 5. 公式与规则
    ApplyEngineeringLogic oPart, refPrd, infoSet
End Sub
Sub nosamebdy_bdylst(oprt)
    Dim lstPara, lstbdys, colls, oDic, keeplst, currobj, objkey, itm, bdy
    Set lstPara = oprt.Parameters.RootParameterSet.ParameterSets.item("Part_info")
    Set lstbdys = lstPara.DirectParameters.item("iBodys")
  Set colls = lstbdys.ValueList
    Set oDic = KCL.InitDic
    Set keeplst = KCL.InitDic
    For Each currobj In colls
        objkey = KCL.GetInternalName(currobj)
        If Not oDic.Exists(objkey) Then
          oDic(objkey) = 1
          keeplst(objkey) = 1
          End If
     Next
     For Each itm In colls
      colls.Remove itm.name
        Next itm
    For Each bdy In oprt.bodies
           If keeplst.Exists(KCL.GetInternalName(bdy)) Then colls.Add bdy
    Next
End Sub



Private Sub EnsureParamExists(ByVal coll As Object, ByVal name As String, ByVal pType As String)
    On Error Resume Next
    Dim p As Object: Set p = coll.item(name)
    If p Is Nothing Then
        Select Case pType
            Case "String": coll.CreateString name, ""
            Case "Mass", "Density", "Length", "Volume": coll.CreateDimension name, pType, 0
        End Select
    End If
    On Error GoTo 0
End Sub

Private Sub EnsurePublication(ByVal refPrd As Object, ByVal pubName As String, ByVal refPath As String)
    On Error Resume Next
    Dim pubs As Object: Set pubs = refPrd.Publications
    Dim oPub As Object: Set oPub = pubs.item(pubName)
    If oPub Is Nothing Then
        Dim oRef As Object: Set oRef = refPrd.CreateReferenceFromName(refPrd.PartNumber & "\" & refPath)
        Set oPub = pubs.Add(pubName)
        pubs.SetDirect pubName, oRef
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyEngineeringLogic(ByVal oPart As Object, ByVal refPrd As Object, ByVal infoSet As Object)
    Dim rels As Object: Set rels = oPart.Relations
    On Error Resume Next
    
    ' 规则：体积计算
    Dim rule As Object: Set rule = rels.item("sum_all_Vol")
    If rule Is Nothing Then Set rule = rels.CreateProgram("sum_all_Vol", "计算体积", EKL_STR)
    rule.Hidden = True
    
    ' 公式：质量计算
    CreateFormula rels, "cal_Mass", infoSet.DirectParameters.item("Mass"), "round(Part_info\sumVol * Part_info\Density, ""kg"", 3)"
    
    ' 链接到 UserRefProperties
    CreateFormula rels, "link_Mass", refPrd.UserRefProperties.item("Mass"), "Part_info\Mass"
    CreateFormula rels, "link_Thickness", refPrd.UserRefProperties.item("Thickness"), "Part_info\Thickness"
    CreateFormula rels, "link_Density", refPrd.UserRefProperties.item("Density"), "Part_info\Density"
    On Error GoTo 0
End Sub

Private Sub CreateFormula(ByVal rels As Object, ByVal fName As String, ByVal target As Object, ByVal source As String)
    Dim f As Object: Set f = Nothing
    On Error Resume Next
    Set f = rels.item(fName)
    If f Is Nothing Then Set f = rels.CreateFormula(fName, "", target, source)
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' 交互与工具
'---------------------------------------------------------------------------------------

' 单选产品/零件 (catSel1)
Public Function catSel1(ByVal prompt As String, Optional ByVal filter As String = "Product") As Object
    On Error Resume Next
    Set catSel1 = KCL.SelectItem(prompt, filter)
    On Error GoTo 0
End Function

Public Function catSelx(ByVal prompt As String) As Collection
    Set catSelx = New Collection
    If mSelection Is Nothing Then Exit Function
    mSelection.Clear
    
    Dim filters(0) As Variant: filters(0) = "Product"
    If mSelection.SelectElement3(filters, prompt, True, 2, True) = "Normal" Then
        Dim i As Long
        For i = 1 To mSelection.count
            catSelx.Add mSelection.item(i).LeafProduct
        Next
    End If
    mSelection.Clear
End Function


Public Sub freePrd()
    Set mProduct = Nothing
End Sub

Public Sub SelectAll(ByVal iQuery As String)
    If mSelection Is Nothing Or mProduct Is Nothing Then Exit Function
    On Error Resume Next
    mSelection.Clear
    mSelection.Add mProduct
    mSelection.Search iQuery & ",sel"
    On Error GoTo 0
End Sub

Public Sub DeleteAll(ByVal iQuery As String)
    If mSelection Is Nothing Then Exit Function
    On Error Resume Next
    mSelection.Clear
    mSelection.Search iQuery & ",sel"
    If mSelection.Count2 <> 0 Then mSelection.Delete
    On Error GoTo 0
End Sub

