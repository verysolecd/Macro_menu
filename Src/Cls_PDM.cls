VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_PDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type PropResult
    obj As Object
    Value As Variant
    IsValid As Boolean
End Type
Public Event ProductChanged()
' 私有变量
Private mProduct As Object          ' 存储当前产品
Private mSelection As Object        ' 存储当前选择集
Private mRows As Collection         ' 用于递归收集 BOM 行数据的临时集合
' 变量及辅助数组 (Updated for Array Storage)
Private mItems() As Bomline       ' Dynamic array for Bomlines
Private mItemCount As Long        ' Current count of items
Private Const CHUNK_SIZE As Long = 150 ' Array growth chunk size
'-------------------------------------------------------
Private EKL_STR As String 'As String '= "let lst(list) set lst=Part_info\iBodys let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol =V"
Private Const ekl1 As String = "let lst(list) set lst=Part_info\iBodys let V (Volume) V=0 let i(integer) i=1"
Private Const ekl2 As String = " for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol =V"
Private Const mdlname As String = "Cls_PDM"
Private Sub Class_Initialize()
    On Error Resume Next
    Set g_allPN = KCL.InitDic ' 保持与全局变量同步
    Set mProduct = Nothing
    mItemCount = 0
    ReDim mItems(1 To CHUNK_SIZE)
    If CATIA.Documents.count > 0 Then
        Set mSelection = CATIA.ActiveDocument.Selection
        mSelection.Clear
        Set rootDoc = CATIA.ActiveDocument
        Set rootPrd = CATIA.ActiveDocument.Product
        If Not rootPrd Is Nothing Then rootPrd.ApplyWorkMode (3)
    End If
    On Error GoTo 0
End Sub

Public Property Get CurrentProduct() As Object
    Set CurrentProduct = mProduct
End Property

Public Property Set CurrentProduct(ByVal newProduct As Object)
    Set mProduct = newProduct
    RaiseEvent ProductChanged
End Property

' 统一获取属性/参数的值
' 返回：PropResult 类型 (Obj, Value, IsValid)
Private Function GetPropertyValue(ByVal Collection As Object, ByVal ItemIdentifier As Variant) As PropResult
    On Error Resume Next
    Dim Result As PropResult
    Set Result.obj = Collection.item(ItemIdentifier)
    If Err.Number = 0 And Not Result.obj Is Nothing Then
        Result.Value = Result.obj.Value
        Result.IsValid = True
    Else
        Set Result.obj = Nothing
        Result.Value = "__"
        Result.IsValid = False
    End If
    GetPropertyValue = Result
    On Error GoTo 0
End Function
'---------------------------------------------------------------------------------------
' BOM 生成核心逻辑 (性能优化版)
'---------------------------------------------------------------------------------------
' Helper to add item to array with auto-resize
Private Sub pushBOM(item As Bomline)
    mItemCount = mItemCount + 1
    If mItemCount > UBound(mItems) Then
        ReDim Preserve mItems(1 To mItemCount + CHUNK_SIZE)
    End If
    mItems(mItemCount) = item
End Sub
' 获取产品基本属性信息 (返回 Bomline)
Public Function getBomLine(ByVal oPrd As Object) As Bomline
    Dim item As Bomline
    Dim refPrd: Set refPrd = oPrd.ReferenceProduct
    Dim userProps As Object: Set userProps = refPrd.UserRefProperties
    
    With item
        .partNumber = refPrd.partNumber
        .Nomenclature = refPrd.Nomenclature
        .Definition = refPrd.Definition
        .InstanceName = oPrd.name
        .Material = GetPropertyValue(userProps, "Material").Value
        .Density = val(GetPropertyValue(userProps, "Density").Value)
        .Mass = val(GetPropertyValue(userProps, "Mass").Value)
        .Thickness = val(GetPropertyValue(userProps, "Thickness").Value)
        .Quantity = QtyinLV(oPrd)
    End With
    
    getBomLine = item
End Function

' 递归计算并更新产品及其子产品的总质量 (恢复并优化版)
Public Function Assmass(ByVal oPrd As Object) As Double
    Dim children As Products: Set children = oPrd.Products
    Dim totalMass As Double: totalMass = 0
    Dim i As Long
    On Error Resume Next
    If children.count > 0 Then
        For i = 1 To children.count
            totalMass = totalMass + Assmass(children.item(i))
        Next
        oPrd.ReferenceProduct.UserRefProperties.item("Mass").Value = totalMass
    Else
        totalMass = oPrd.ReferenceProduct.UserRefProperties.item("Mass").Value
    End If
    On Error GoTo 0
    Assmass = totalMass
End Function

' 优化计数逻辑：按需统计同级 PN
Private Function QtyinLV(ByVal oPrd As Object) As Long
    On Error Resume Next
    If Not TypeOf oPrd.Parent Is Products Then
        QtyinLV = 1: Exit Function
    End If
    Dim productsColl As Products: Set productsColl = oPrd.Parent
    Dim targetPN As String: targetPN = oPrd.ReferenceProduct.partNumber
    Dim i As Long, cnt As Long: cnt = 0
    For i = 1 To productsColl.count
        If productsColl.item(i).ReferenceProduct.partNumber = targetPN Then cnt = cnt + 1
    Next
    QtyinLV = cnt
    On Error GoTo 0
End Function

' 高效 BOM 生成 (主入口) - 返回 Typed Array
Public Function ProduceBOM(ByVal rootPrd As Object, Optional ByVal filterDuplicate As Boolean = True) As Bomline()
    ' Reset storage
    mItemCount = 0
    ReDim mItems(1 To CHUNK_SIZE)
    ' 内部递归执行
    TraverseProductTree rootPrd, 1, filterDuplicate
    If mItemCount = 0 Then Exit Function
    ' Trim array to actual size and return
    Dim Result() As Bomline
    ReDim Result(1 To mItemCount)
    Dim i As Long
    For i = 1 To mItemCount
        Result(i) = mItems(i)
    Next
    ProduceBOM = Result
End Function

' 递归辅助函数
Private Sub TraverseProductTree(ByVal oPrd As Object, ByVal level As Long, ByVal filterDuplicate As Boolean)
    Dim item As Bomline
    item = getBomLine(oPrd)
    item.level = level
    
    ' Use Array instead of Collection
    pushBOM item
    
    Dim children As Products: Set children = oPrd.Products
    If children.count > 0 Then
        Dim dict As Object: Set dict = KCL.InitDic
        Dim i As Long
        For i = 1 To children.count
            Dim child As Object: Set child = children.item(i)
            If filterDuplicate Then
                If Not dict.Exists(child.ReferenceProduct.partNumber) Then ' Fixed: use ReferenceProduct
                    dict.Add child.ReferenceProduct.partNumber, 1
                    TraverseProductTree child, level + 1, filterDuplicate
                End If
            Else
                TraverseProductTree child, level + 1, filterDuplicate
            End If
        Next
    End If
End Sub


Public Function attLv2Prd(oPrd) As Variant
    ' 此方法仅取两层信息
    mItemCount = 0
    ReDim mItems(1 To CHUNK_SIZE)
    TraverseTwoLevels oPrd
    If mItemCount = 0 Then attLv2Prd = Empty: Exit Function
    Dim Result() As Variant: ReDim Result(1 To mItemCount, 1 To 9)
    Dim R As Long, C As Long
    ' Convert Bomline array to Variant array for legacy compatibility
    ' Note: This mapping matches the legacy expected format (approx)
    For R = 1 To mItemCount
        With mItems(R)
            Result(R, 1) = .partNumber
            Result(R, 2) = .Nomenclature
            Result(R, 3) = .Definition
            Result(R, 4) = .InstanceName
            Result(R, 5) = .Material
            Result(R, 6) = .Density
            Result(R, 7) = .Mass
            Result(R, 8) = .Thickness
            Result(R, 9) = .Quantity
        End With
    Next
    attLv2Prd = Result
End Function

' 获取用于 RW_Revise 的强类型数据 (替代 attLv2Prd)
Public Function GetReviseItems(ByVal oPrd As Object) As Bomline()
    mItemCount = 0
    ReDim mItems(1 To CHUNK_SIZE)
    TraverseTwoLevels oPrd
    
    If mItemCount = 0 Then Exit Function
    
    Dim Result() As Bomline
    ReDim Result(1 To mItemCount)
    Dim i As Long
    For i = 1 To mItemCount
        Result(i) = mItems(i)
    Next
    GetReviseItems = Result
End Function

Private Sub TraverseTwoLevels(ByVal oPrd As Object)
    pushBOM getBomLine(oPrd)     ' 根产品
    Dim children As Products: Set children = oPrd.Products
    Dim i As Long
    For i = 1 To children.count
        pushBOM getBomLine(children.item(i))    ' 子产品
    Next
End Sub
Public Function modatt(ByVal oPrd As Object, ByVal oData As Variant)
    Dim refPrd As Object: Set refPrd = oPrd.ReferenceProduct
    Dim currAtt As Bomline: currAtt = getBomLine(oPrd)
    Dim i As Long
    On Error Resume Next
    For i = 1 To 6
        ' Check if data is present and different from current attribute
        If oData(i) <> "" Then
            Dim isChanged As Boolean: isChanged = False
            
            Select Case i
                Case 1: If oData(1) <> currAtt.partNumber Then isChanged = True
                Case 2: If oData(2) <> currAtt.Nomenclature Then isChanged = True
                Case 3: If oData(3) <> currAtt.Definition Then isChanged = True
                Case 4: If oData(4) <> currAtt.InstanceName Then isChanged = True
                Case 5: If oData(5) <> currAtt.Material Then isChanged = True
                Case 6: If val(oData(6)) <> currAtt.Density Then isChanged = True
            End Select
            If isChanged Then
                Select Case i
                    Case 1: refPrd.partNumber = oData(i)
                    Case 2: refPrd.Nomenclature = oData(i)
                    Case 3: refPrd.Definition = oData(i)
                    Case 4: oPrd.name = oData(i)
                    Case 5: refPrd.UserRefProperties.item("Material").Value = oData(i)
                    Case 6: refPrd.Parent.part.Parameters.RootParameterSet.ParameterSets.item("Part_info").DirectParameters.item("Density").Value = oData(i)
                End Select
            End If
        End If
    Next
    On Error GoTo 0
End Function

Public Sub initPrd(ByVal oPrd As Object)
    Dim refPrd As Object: Set refPrd = oPrd.ReferenceProduct
    EnsureParamExists refPrd.UserRefProperties, "Mass", "Mass"
    Dim oPart As Object
    On Error Resume Next
    Set oPart = refPrd.Parent.part
    On Error GoTo 0
    
    If Not oPart Is Nothing Then
        iniPrt oPrd
    Else ' 仅产品：发布属性
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
        
        If Not oPart.MainBody Is Nothing Then  ' 将 MainBody 同步到 iBodys 列表
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
    Dim pubs As Object
    Dim oPub As Object:
'    Set pubs = refPrd.Publications
'    Set oPub = pubs.item(pubName)
'    Dim isupdated As Boolean: isupdated = True
'    If IsNothing(oPub) Then
'        Set oPub = pubs.Add(pubName)
'        isupdated = False
'    Else
'        Dim valObj As Object: Set valObj = Nothing
'            Set valObj = oPub.Valuation
'        If valObj Is Nothing Then
'            isupdated = False
''            pubs.Remove (pubName)
''            Set oPub = pubs.Add(pubName)
'        End If
'    End If
'    If isupdated = False Then
        Dim oRef As Object: Set oRef = refPrd.CreateReferenceFromName("." & "\" & refPath)
        pubs.SetDirect pubName, oRef
        
        If Err.Number <> 0 Then
            Debug.Print "发布失败：" & pubName & " → SetDirect执行错误"
            Exit Sub
        End If
    End If
    refPrd.Update
        Set oPub = pubs.item(pubName)
        If oPub Is Nothing Or oPub.Valuation Is Nothing Then
            Debug.Print "发布失败：" & pubName & " → 发布估值无效"
            Exit Sub
        End If
    On Error GoTo 0
End Sub

Private Sub ApplyEngineeringLogic(ByVal oPart As Object, ByVal refPrd As Object, ByVal infoSet As Object)
    Dim rels As Object: Set rels = oPart.Relations
    On Error Resume Next
    ' 规则：体积计算
    Dim rule As Object: Set rule = rels.item("sum_all_Vol")
    EKL_STR = ekl1 & ekl2
    If rule Is Nothing Then
        Set rule = rels.CreateProgram("sum_all_Vol", "计算体积", EKL_STR)
    Else
         rule.Modify EKL_STR
    End If
    rule.Hidden = True
    ' 公式：质量计算
    CreateFormula rels, "cal_Mass", infoSet.DirectParameters.item("Mass"), "round(Part_info\sumVol * Part_info\Density, ""kg"", 4)"
    CreateFormula rels, "link_Mass", refPrd.UserRefProperties.item("Mass"), "Part_info\Mass"   ' 链接到 UserRefProperties
    CreateFormula rels, "link_Thickness", refPrd.UserRefProperties.item("Thickness"), "Part_info\Thickness"
    CreateFormula rels, "link_Density", refPrd.UserRefProperties.item("Density"), "Part_info\Density"
    On Error GoTo 0
End Sub

Private Sub CreateFormula(ByVal rels As Object, ByVal fName As String, ByVal targetPara As Object, ByVal F_value As String)
    Dim f As Object: Set f = Nothing
    On Error Resume Next
        Set f = rels.item(fName)
        If f Is Nothing Then
            Set f = rels.CreateFormula(fName, "", targetPara, F_value)
        Else
             If StrComp(f.Value, F_value, vbTextCompare) <> 0 Then f.Modify (F_value)
        End If
    f.Hidden = True
    On Error GoTo 0
End Sub


Public Sub DelQuery(ByVal iQuery As String, Optional ByVal iRange = Nothing)
    If mSelection Is Nothing Then Exit Sub
On Error Resume Next
  If iRange Is Nothing Then
    mSelection.Search iQuery & ",all"
  Else
     mSelection.Add iRange
     mSelection.Search iQuery & ",sel"
  End If
    If mSelection.Count2 <> 0 Then mSelection.Delete
    Error.Clear
On Error GoTo 0
End Sub

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


