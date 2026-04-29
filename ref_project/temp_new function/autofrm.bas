' ==============================================================================
' 模块：AutoAssembleByDiameter
' 功能：点击零件 -> 输入/设定孔径 -> 自动识别所有符合该孔径的孔 -> 批量装配紧固件
' ==============================================================================
Option Explicit

Sub AutoAssembleByDiameter()
    ' ==========================================
    ' 【配置区域】
    ' ==========================================
    Dim FastenerPath As String
    ' 1. 设置紧固件绝对路径
    ' 请务必修改为你本地实际存在的零件路径
    FastenerPath = "C:\CATIA_Standard_Parts\Bolts\M8_Bolt.CATPart"
    
    ' 2. 设置目标孔径 (单位: mm)
    '    这里你可以写死 (例如 8)，或者改成 InputBox 让用户运行时输入
    Dim TargetDia As Double
    Dim UserInput As String
    UserInput = InputBox("请输入目标孔的直径 (mm)：", "自动装配设置", "8")
    If UserInput = "" Then Exit Sub
    TargetDia = CDbl(UserInput)
    ' ==========================================

    Dim oProdDoc As ProductDocument
    Set oProdDoc = CATIA.ActiveDocument
    
    Dim oRootProd As Product
    Set oRootProd = oProdDoc.Product
    
    Dim oSel As Selection
    Set oSel = oProdDoc.Selection

    ' 1. 让用户选择目标零件 (Product Instance)
    oSel.Clear
    Dim InputObjectType(0)
    InputObjectType(0) = "Product" ' 限制只能选产品/零件节点
    
    Dim Status As String
    Status = oSel.SelectElement2(InputObjectType, "请点击包含孔的目标零件 (Part Instance)", False)
    
    If Status = "Cancel" Then Exit Sub
    
    Dim oTargetProd As Product
    Set oTargetProd = oSel.Item(1).LeafProduct ' 获取选中的具体零件实例
    
    ' 获取该实例在装配树上的全路径名称 (用于后面生成Reference)
    ' 获取 SelectedElement 对象的完整路径是做上下文关联的关键
    Dim oTargetProdInstanceName As String
    oTargetProdInstanceName = oTargetProd.Name ' 例如 "Plate.1"
    
    ' 如果是多级装配，这里可能需要更复杂的路径处理，
    ' 为简单起见，假设选中零件直接位于当前ActiveDocument的下一级，或者我们只处理Part内部几何。
    
    ' 2. 进入零件文档进行几何搜索
    '    我们需要深入到 ReferenceProduct (即 PartDocument) 去扫描面
    Dim oPartDoc As PartDocument
    On Error Resume Next
    Set oPartDoc = oTargetProd.ReferenceProduct.Parent
    If Err.Number <> 0 Then
        MsgBox "选中的对象不是一个Part，无法扫描几何。", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 3. 搜索符合直径的圆柱面
    '    为了不干扰当前装配的选择集，我们创建一个临时选择集或者利用PartDoc的Selection
    Dim oPartSel As Selection
    Set oPartSel = oPartDoc.Selection
    oPartSel.Clear
    
    ' 搜索所有面
    ' 使用 Search 是极其快速的
    oPartSel.Search "Topology.CGMFace,all"
    
    If oPartSel.Count = 0 Then
        MsgBox "在该零件中未找到任何面。", vbExclamation
        Exit Sub
    End If
    
    ' 准备测量工具
    Dim oSPA As SPAWorkbench
    Set oSPA = oPartDoc.GetWorkbench("SPAWorkbench")
    Dim oMeas As Measurable
    Dim matchedFaces As Collection
    Set matchedFaces = New Collection
    
    Dim i As Integer
    Dim detectedRadius As Double
    Dim TargetRadius As Double
    TargetRadius = TargetDia / 2
    Dim Tolerance As Double
    Tolerance = 0.01 ' 允许的误差范围 (mm)
    
    ' 遍历搜索结果
    For i = 1 To oPartSel.Count
        Set oMeas = oSPA.GetMeasurable(oPartSel.Item(i).Reference)
        
        If oMeas.GeometryName = CatGeometryName.CatCylindrical Then
            detectedRadius = oMeas.Radius
            ' 检查半径是否匹配 (注意单位转换，CATIA内部无论如何通常是mm，但API有时候是m，取决于设置)
            ' SPAWorkbench通常返回文档单位(mm)，保险起见我们认为是文档单位
            
            If Abs(detectedRadius - TargetRadius) < Tolerance Then
                ' 找到了！
                ' 关键点：我们需要保留这个面的 Reference，但由于它是在PartDoc里的，
                ' 我们需要在 Assembly Context 下使用它。
                ' 使用 BRep Name 是最通用的方法。
                matchedFaces.Add oPartSel.Item(i).Reference
            End If
        End If
    Next i
    
    oPartSel.Clear ' 清理Part内的选择
    
    If matchedFaces.Count = 0 Then
        MsgBox "未找到直径为 " & TargetDia & "mm 的孔。", vbInformation
        Exit Sub
    End If
    
    Dim Confirm As Integer
    Confirm = MsgBox("检测到 " & matchedFaces.Count & " 个匹配的孔。是否全部装配？", vbYesNo + vbQuestion)
    If Confirm = vbNo Then Exit Sub


    ' 4. 开始批量装配
    CATIA.RefreshDisplay = False ' 冻结屏幕，加速
    
    Dim oConstraints As Constraints
    Set oConstraints = oRootProd.Connections("CATIAConstraints")
    
    Dim faceItem As Reference
    Dim oFormatRef As Reference
    Dim oNewBolt As Product
    Dim oBoltAxis As Reference
    Dim successCount As Integer
    successCount = 0
    
    For Each faceItem In matchedFaces
        ' 4.1 构造装配上下文中的 Reference
        ' 我们手里的 faceItem 是 Part 里的原始引用。
        ' 我们需要基于当前选中的 Product Instance (oTargetProd) 创建一个装配引用。
        ' 利用 CreateReferenceFromName 和 BRep 路径
        
        ' 获取面的内部名称 (类似于 Selection_RSur:(Face:(Brp:....)))
        Dim faceName As String
        faceName = faceItem.DisplayName
        
        ' 组合路径： ProductInstanceName/!FaceName
        ' 注意：如果你的 oTargetProd 不是根节点的直接子节点，这里需要完整的实例路径链
        ' 为了代码鲁棒性，这里假设是直接子节点。如果是多层级，需要拼接路径。
        On Error Resume Next
        Set oFormatRef = oRootProd.CreateReferenceFromName(oTargetProd.Name & "/!" & faceName)
        
        If Err.Number = 0 Then
            ' 4.2 插入紧固件
            Set oNewBolt = InsertFastenerQuick(oRootProd, FastenerPath)
            
            If Not oNewBolt Is Nothing Then
                ' 4.3 找螺栓轴线
                Set oBoltAxis = GetFastenerAxisRefQuick(oNewBolt)
                
                If Not oBoltAxis Is Nothing Then
                    ' 4.4 约束
                    Dim oCst As Constraint
                    Set oCst = oConstraints.AddBiEltCst(catCstTypeOn, oFormatRef, oBoltAxis)
                    oCst.Mode = catCstModeDrivingDimension
                    successCount = successCount + 1
                End If
            End If
        End If
        On Error GoTo 0
    Next faceItem
    
    oRootProd.Update
    CATIA.RefreshDisplay = True
    
    MsgBox "完成！已在 " & successCount & " 个位置插入了紧固件。", vbInformation

End Sub

' --- 快速插入辅助函数 ---
Function InsertFastenerQuick(rootProd As Product, path As String) As Product
    On Error Resume Next
    Dim arr(0)
    arr(0) = path
    Set InsertFastenerQuick = rootProd.Products.AddComponentsFromFiles(arr, "All")(1)
    On Error GoTo 0
End Function

' --- 快速找轴线辅助函数 ---
Function GetFastenerAxisRefQuick(iProd As Product) As Reference
    On Error Resume Next
    ' 1. 尝试直接读取发布
    Set GetFastenerAxisRefQuick = iProd.ReferenceProduct.Publications.Item("Axis").Valuation
    If Err.Number = 0 Then Exit Function
    
    ' 2. 也是盲搜第一个圆柱面
    Dim pDoc As PartDocument
    Set pDoc = iProd.ReferenceProduct.Parent
    Dim sel As Selection
    Set sel = pDoc.Selection
    sel.Search "Topology.CGMFace,all"
    
    Dim i As Integer
    Dim spa As SPAWorkbench
    Set spa = pDoc.GetWorkbench("SPAWorkbench")
    For i = 1 To sel.Count
        If spa.GetMeasurable(sel.Item(i).Reference).GeometryName = CatGeometryName.CatCylindrical Then
            Set GetFastenerAxisRefQuick = iProd.CreateReferenceFromName(iProd.Name & "/!" & sel.Item(i).Reference.DisplayName)
            sel.Clear
            Exit Function
        End If
    Next i
    sel.Clear
    Set GetFastenerAxisRefQuick = Nothing
End Function
